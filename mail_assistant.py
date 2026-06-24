"""
PST 信件 AI 摘要助手 (v0.2)
- 預過濾系統信/自動回覆/廣告 -> 不打 API
- 短信跳過 API
- 大量信件用 Batch API (5 折)
- 內文截短預設 2000 字
"""

import os
import re
import sys
import json
import time
from pathlib import Path

import win32com.client
import pythoncom


def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


APP_DIR = app_dir()
CONFIG_PATH = APP_DIR / "config.json"
STATE_FILE = APP_DIR / ".processed.json"

SUMMARY_PROMPT = """你是一個郵件助理。請閱讀以下郵件,輸出 JSON,格式:
{{
  "category": "<簡短分類,例如:客戶詢價、內部公告、系統通知、會議邀請、廣告、其他>",
  "summary": "<3-5 行重點摘要,中文>",
  "action_items": ["<待辦1>", "<待辦2>"],
  "priority": "<high|medium|low>"
}}

只輸出 JSON,不要其他文字。

寄件人: {sender}
主旨: {subject}
日期: {date}

內文:
{body}
"""


def load_config():
    if not CONFIG_PATH.exists():
        print(f"找不到 config.json,請複製 config.example.json 為 config.json 並修改。")
        sys.exit(1)
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def load_state():
    if STATE_FILE.exists():
        return set(json.loads(STATE_FILE.read_text(encoding="utf-8")))
    return set()


def save_state(processed):
    STATE_FILE.write_text(json.dumps(list(processed)), encoding="utf-8")


def find_pst_folder(ns, pst_name, folder_name):
    for store in ns.Stores:
        if pst_name.lower() in store.DisplayName.lower():
            root = store.GetRootFolder()
            for folder in root.Folders:
                if folder.Name.lower() == folder_name.lower() or folder.Name in ("收件匣", "收件箱", "Inbox"):
                    return folder
            return root.Folders.Item(1)
    raise RuntimeError(f"找不到 PST: {pst_name}")


def safe_filename(s, maxlen=80):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", s)
    return s[:maxlen].strip() or "untitled"


def should_skip(mail, cfg):
    """規則過濾 — 不打 API 直接分類"""
    sender = (getattr(mail, "SenderEmailAddress", "") or "").lower()
    sender_name = (mail.SenderName or "").lower()
    subject = mail.Subject or ""

    for kw in cfg.get("skip_senders", []):
        if kw.lower() in sender or kw.lower() in sender_name:
            return "系統通知"
    for kw in cfg.get("skip_subject_keywords", []):
        if kw.lower() in subject.lower():
            return "自動郵件"

    if getattr(mail, "MessageClass", "").startswith("IPM.Schedule"):
        return "會議邀請"

    body = mail.Body or ""
    if len(body.strip()) < cfg.get("min_body_chars", 200):
        return "簡短訊息"

    return None


def build_prompt(mail, max_chars):
    body = (mail.Body or "")[:max_chars]
    return SUMMARY_PROMPT.format(
        sender=mail.SenderName or "(unknown)",
        subject=mail.Subject or "(no subject)",
        date=str(mail.ReceivedTime),
        body=body,
    )


def parse_result(text):
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.MULTILINE)
    data = json.loads(text)
    # 防呆:summary 偶爾被回成 list
    s = data.get("summary")
    if isinstance(s, list):
        data["summary"] = "\n".join(str(x) for x in s)
    elif s is not None and not isinstance(s, str):
        data["summary"] = str(s)
    # action_items 確保是 list of str
    ai = data.get("action_items", [])
    if not isinstance(ai, list):
        data["action_items"] = [str(ai)] if ai else []
    else:
        data["action_items"] = [str(x) for x in ai]
    # category / priority 確保 string
    for k in ("category", "priority"):
        if k in data and not isinstance(data[k], str):
            data[k] = str(data[k])
    return data


def write_markdown(mail, result, output_dir):
    date = mail.ReceivedTime
    date_dir = output_dir / date.strftime("%Y-%m-%d")
    date_dir.mkdir(parents=True, exist_ok=True)

    sender = safe_filename(mail.SenderName or "unknown", 30)
    subject = safe_filename(mail.Subject or "no_subject", 50)
    fname = f"{date.strftime('%H%M')}_{sender}_{subject}.md"

    action_lines = "\n".join(f"- [ ] {a}" for a in result.get("action_items", [])) or "- (無)"

    content = f"""---
category: {result.get('category', '其他')}
priority: {result.get('priority', 'medium')}
from: {mail.SenderName}
date: {date.strftime('%Y-%m-%d %H:%M')}
subject: {mail.Subject}
---

# {mail.Subject}

**寄件人**: {mail.SenderName} <{getattr(mail, 'SenderEmailAddress', '')}>
**日期**: {date.strftime('%Y-%m-%d %H:%M')}
**分類**: {result.get('category', '其他')}
**優先度**: {result.get('priority', 'medium')}

## 摘要
{result.get('summary', '')}

## 待辦
{action_lines}
"""
    (date_dir / fname).write_text(content, encoding="utf-8")
    return date_dir / fname


def write_digest(digest, output_dir):
    """跑完總結:HTML 追蹤清單"""
    from datetime import datetime
    from html import escape
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    high = [(m, r) for m, r in digest if r.get("priority") == "high"]
    medium = [(m, r) for m, r in digest if r.get("priority") == "medium"]
    low = [(m, r) for m, r in digest if r.get("priority") == "low"]
    with_action = [(m, r) for m, r in digest if r.get("action_items")]
    total_actions = sum(len(r.get("action_items", [])) for _, r in with_action)

    by_cat = {}
    for m, r in digest:
        by_cat.setdefault(r.get("category", "其他"), []).append((m, r))

    def mail_card(m, r, priority_color):
        date = m.ReceivedTime.strftime("%m-%d %H:%M")
        sender = escape(m.SenderName or "")
        subject = escape(m.Subject or "(no subject)")
        summary = escape(r.get("summary", "")).replace("\n", "<br>")
        cat = escape(r.get("category", ""))
        actions_html = ""
        if r.get("action_items"):
            actions_html = "<ul class='actions'>" + "".join(
                f"<li><label><input type='checkbox'> {escape(a)}</label></li>"
                for a in r["action_items"]
            ) + "</ul>"
        return f"""
        <div class="card" style="border-left-color:{priority_color}">
          <div class="card-head">
            <span class="date">{date}</span>
            <span class="sender">{sender}</span>
            <span class="cat">{cat}</span>
          </div>
          <div class="subject">{subject}</div>
          <div class="summary">{summary}</div>
          {actions_html}
        </div>"""

    cat_stats = "".join(
        f"<tr><td>{escape(cat)}</td><td>{len(mails)}</td></tr>"
        for cat, mails in sorted(by_cat.items(), key=lambda x: -len(x[1]))
    )

    high_html = "".join(mail_card(m, r, "#e74c3c") for m, r in high) or "<p class='empty'>(無)</p>"
    medium_html = "".join(mail_card(m, r, "#f39c12") for m, r in medium) or "<p class='empty'>(無)</p>"
    low_html = "".join(mail_card(m, r, "#95a5a6") for m, r in low) or "<p class='empty'>(無)</p>"

    action_html = ""
    if with_action:
        for m, r in sorted(with_action, key=lambda x: {"high":0,"medium":1,"low":2}.get(x[1].get("priority"),3)):
            date = m.ReceivedTime.strftime("%m-%d")
            subject = escape((m.Subject or "")[:60])
            sender = escape(m.SenderName or "")
            pri = r.get("priority", "medium")
            pri_color = {"high":"#e74c3c","medium":"#f39c12","low":"#95a5a6"}.get(pri,"#95a5a6")
            items = "".join(f"<li><label><input type='checkbox'> {escape(a)}</label></li>" for a in r["action_items"])
            action_html += f"""
            <div class="action-block" style="border-left-color:{pri_color}">
              <div class="action-head"><b>[{date}] {subject}</b> <span class="meta">— {sender}</span></div>
              <ul>{items}</ul>
            </div>"""
    else:
        action_html = "<p class='empty'>(無)</p>"

    html = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<title>信件追蹤清單 — {now}</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{ font-family: -apple-system, "Segoe UI", "Microsoft JhengHei", sans-serif; max-width: 1100px; margin: 0 auto; padding: 24px; background: #f5f6fa; color: #2c3e50; line-height: 1.6; }}
  h1 {{ margin: 0 0 4px; }}
  .meta-bar {{ color: #7f8c8d; margin-bottom: 24px; font-size: 14px; }}
  .stats {{ display: flex; gap: 12px; margin-bottom: 28px; flex-wrap: wrap; }}
  .stat {{ background: white; padding: 14px 20px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); flex: 1; min-width: 140px; }}
  .stat .num {{ font-size: 28px; font-weight: bold; }}
  .stat .lbl {{ font-size: 12px; color: #7f8c8d; }}
  .stat.high .num {{ color: #e74c3c; }}
  .stat.action .num {{ color: #27ae60; }}
  h2 {{ margin-top: 32px; padding-bottom: 8px; border-bottom: 2px solid #ecf0f1; }}
  .card {{ background: white; padding: 14px 18px; margin-bottom: 10px; border-radius: 6px; border-left: 4px solid; box-shadow: 0 1px 2px rgba(0,0,0,0.04); }}
  .card-head {{ font-size: 12px; color: #7f8c8d; margin-bottom: 4px; }}
  .card-head .sender {{ margin-left: 12px; color: #34495e; }}
  .card-head .cat {{ float: right; background: #ecf0f1; padding: 2px 8px; border-radius: 10px; }}
  .subject {{ font-weight: 600; font-size: 15px; margin: 4px 0; }}
  .summary {{ font-size: 14px; color: #555; }}
  ul.actions {{ margin: 8px 0 0; padding-left: 0; list-style: none; }}
  ul.actions li {{ padding: 3px 0; }}
  .action-block {{ background: white; padding: 12px 16px; margin-bottom: 8px; border-radius: 6px; border-left: 4px solid; }}
  .action-block ul {{ margin: 6px 0 0; padding-left: 20px; }}
  .action-block .meta {{ color: #95a5a6; font-weight: normal; }}
  table {{ background: white; border-collapse: collapse; width: 100%; border-radius: 6px; overflow: hidden; }}
  th, td {{ padding: 10px 14px; text-align: left; border-bottom: 1px solid #ecf0f1; }}
  th {{ background: #34495e; color: white; }}
  .empty {{ color: #95a5a6; font-style: italic; }}
  details {{ background: white; padding: 8px 14px; border-radius: 6px; margin-bottom: 10px; }}
  summary {{ cursor: pointer; font-weight: 600; padding: 4px 0; }}
  input[type=checkbox] {{ margin-right: 6px; }}
  label {{ cursor: pointer; }}
</style>
</head>
<body>
  <h1>📬 信件追蹤清單</h1>
  <div class="meta-bar">產出時間: {now} · 總處理 {len(digest)} 封</div>

  <div class="stats">
    <div class="stat high"><div class="num">{len(high)}</div><div class="lbl">🔴 高優先</div></div>
    <div class="stat"><div class="num">{len(medium)}</div><div class="lbl">🟡 中優先</div></div>
    <div class="stat"><div class="num">{len(low)}</div><div class="lbl">⚪ 低優先</div></div>
    <div class="stat action"><div class="num">{total_actions}</div><div class="lbl">✅ 待辦事項</div></div>
  </div>

  <h2>🔴 高優先處理</h2>
  {high_html}

  <h2>✅ 待辦事項彙整</h2>
  {action_html}

  <details>
    <summary>🟡 中優先 ({len(medium)})</summary>
    {medium_html}
  </details>

  <details>
    <summary>⚪ 低優先 ({len(low)})</summary>
    {low_html}
  </details>

  <h2>📊 分類統計</h2>
  <table>
    <thead><tr><th>分類</th><th>數量</th></tr></thead>
    <tbody>{cat_stats}</tbody>
  </table>
</body>
</html>"""

    (output_dir / "_追蹤清單.html").write_text(html, encoding="utf-8")


def rule_based_result(mail, category):
    return {
        "category": category,
        "summary": (mail.Body or "")[:300].strip() or "(空)",
        "action_items": [],
        "priority": "low",
    }


class AnthropicProvider:
    def __init__(self, model):
        from anthropic import Anthropic
        self.client = Anthropic()
        self.model = model

    def call_one(self, prompt):
        resp = self.client.messages.create(
            model=self.model,
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
        return resp.content[0].text

    def call_batch(self, prompts):
        """custom_id -> prompt;回傳 custom_id -> text。Anthropic Batch API 5 折。"""
        requests = [
            {
                "custom_id": cid,
                "params": {
                    "model": self.model,
                    "max_tokens": 1024,
                    "messages": [{"role": "user", "content": p}],
                },
            }
            for cid, p in prompts.items()
        ]
        batch = self.client.messages.batches.create(requests=requests)
        print(f"Anthropic Batch ID: {batch.id}")
        while True:
            b = self.client.messages.batches.retrieve(batch.id)
            c = b.request_counts
            print(f"  完成 {c.succeeded}, 處理中 {c.processing}, 失敗 {c.errored}")
            if b.processing_status == "ended":
                break
            time.sleep(30)
        out = {}
        for entry in self.client.messages.batches.results(batch.id):
            if entry.result.type == "succeeded":
                out[entry.custom_id] = entry.result.message.content[0].text
        return out


class GeminiProvider:
    def __init__(self, model):
        from google import genai
        api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
        if not api_key:
            raise RuntimeError("請設 GEMINI_API_KEY 環境變數")
        self.client = genai.Client(api_key=api_key)
        self.model = model

    def call_one(self, prompt):
        from google.genai import types
        resp = self.client.models.generate_content(
            model=self.model,
            contents=prompt,
            config=types.GenerateContentConfig(
                http_options=types.HttpOptions(timeout=60000),
            ),
        )
        return resp.text

    def call_batch(self, prompts):
        """Gemini 沒有 Anthropic 式 batch,逐筆呼叫並顯示進度。"""
        out = {}
        total = len(prompts)
        for i, (cid, p) in enumerate(prompts.items(), 1):
            print(f"  [{i}/{total}] {cid}...", flush=True)
            try:
                out[cid] = self.call_one(p)
            except Exception as e:
                print(f"    失敗: {e}", flush=True)
        return out


def list_models(provider_name):
    """列出 provider 目前可用模型"""
    if provider_name == "gemini":
        from google import genai
        api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
        if not api_key:
            print("請先設 GEMINI_API_KEY")
            return
        client = genai.Client(api_key=api_key)
        print(f"\n=== Gemini 可用模型 ===")
        for m in client.models.list():
            methods = getattr(m, "supported_actions", None) or getattr(m, "supported_generation_methods", [])
            if not methods or "generateContent" in methods:
                name = m.name.replace("models/", "")
                display = getattr(m, "display_name", "") or ""
                print(f"  {name:50s} {display}")
    elif provider_name == "anthropic":
        from anthropic import Anthropic
        if not os.environ.get("ANTHROPIC_API_KEY"):
            print("請先設 ANTHROPIC_API_KEY")
            return
        client = Anthropic()
        print(f"\n=== Anthropic 可用模型 ===")
        for m in client.models.list(limit=50).data:
            print(f"  {m.id:45s} {m.display_name}")
    else:
        print(f"未知 provider: {provider_name}")


def make_provider(cfg):
    p = cfg.get("provider", "anthropic").lower()
    if p == "anthropic":
        return AnthropicProvider(cfg["model"])
    if p == "gemini":
        return GeminiProvider(cfg["model"])
    raise ValueError(f"未知 provider: {p}")


def process_sync(provider, mail_map, cfg):
    results = {}
    total = len(mail_map)
    for i, (cid, mail) in enumerate(mail_map.items(), 1):
        try:
            print(f"  [{i}/{total}] {(mail.Subject or '')[:50]}", flush=True)
            text = provider.call_one(build_prompt(mail, cfg["max_body_chars"]))
            results[cid] = parse_result(text)
            print(f"    -> {results[cid].get('category')}", flush=True)
        except Exception as e:
            print(f"    失敗: {e}", flush=True)
    return results


def process_batch(provider, mail_map, cfg):
    print(f"提交批次處理 ({len(mail_map)} 筆)...")
    prompts = {cid: build_prompt(mail, cfg["max_body_chars"]) for cid, mail in mail_map.items()}
    raw = provider.call_batch(prompts)
    results = {}
    for cid, text in raw.items():
        try:
            results[cid] = parse_result(text)
        except Exception as e:
            print(f"  解析失敗 {cid}: {e}")
    return results


def main():
    if len(sys.argv) > 1 and sys.argv[1] in ("--list-models", "-l"):
        provider_name = sys.argv[2] if len(sys.argv) > 2 else "gemini"
        list_models(provider_name.lower())
        return

    cfg = load_config()
    provider_name = cfg.get("provider", "anthropic").lower()
    if provider_name == "anthropic" and not os.environ.get("ANTHROPIC_API_KEY"):
        print("請設定環境變數 ANTHROPIC_API_KEY")
        sys.exit(1)
    if provider_name == "gemini" and not (os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")):
        print("請設定環境變數 GEMINI_API_KEY")
        sys.exit(1)

    output_dir = Path(cfg["output_dir"])
    if not output_dir.is_absolute():
        output_dir = APP_DIR / output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    inbox = find_pst_folder(ns, cfg["pst_display_name"], cfg.get("folder_name", "Inbox"))
    print(f"找到資料夾: {inbox.FolderPath} (共 {inbox.Items.Count} 封)")

    processed = load_state()
    try:
        inbox.Items.Sort("[ReceivedTime]", True)
    except Exception:
        pass
    items = []
    for it in inbox.Items:
        try:
            if getattr(it, "Class", 0) == 43:
                items.append(it)
        except Exception:
            continue

    skipped = 0
    rule_handled = 0
    to_api = {}
    mail_by_id = {}

    for mail in items:
        try:
            entry_id = mail.EntryID
        except Exception:
            continue
        if entry_id in processed:
            continue
        if getattr(mail, "Class", 0) != 43:
            continue

        skip_cat = should_skip(mail, cfg)
        if skip_cat:
            try:
                write_markdown(mail, rule_based_result(mail, skip_cat), output_dir)
                processed.add(entry_id)
                rule_handled += 1
            except Exception as e:
                print(f"規則寫入失敗: {e}")
                skipped += 1
            continue

        cid = f"mail_{len(to_api)}"
        to_api[cid] = mail
        mail_by_id[cid] = mail

    print(f"\n預過濾結果: 規則直接歸檔 {rule_handled} 封, 跳過 {skipped} 封, 送 API {len(to_api)} 封")

    if not to_api:
        save_state(processed)
        print("沒有需要 AI 摘要的新信。")
        return

    provider = make_provider(cfg)
    threshold = cfg.get("batch_threshold", 10)
    if len(to_api) >= threshold:
        results = process_batch(provider, to_api, cfg)
    else:
        results = process_sync(provider, to_api, cfg)

    written = 0
    digest = []
    for cid, result in results.items():
        mail = mail_by_id[cid]
        try:
            write_markdown(mail, result, output_dir)
            processed.add(mail.EntryID)
            written += 1
            digest.append((mail, result))
        except Exception as e:
            print(f"寫檔失敗 {cid}: {e}")

    save_state(processed)
    write_digest(digest, output_dir)
    print(f"\n完成。規則 {rule_handled} 封 + AI {written} 封,輸出: {output_dir}")
    print(f"追蹤清單: {output_dir / '_追蹤清單.html'} (用瀏覽器開)")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n錯誤: {e}")
        import traceback
        traceback.print_exc()
        input("按 Enter 結束...")
