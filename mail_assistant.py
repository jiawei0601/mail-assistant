"""
PST 信件 AI 摘要助手
讀 Outlook PST 收件匣,用 Claude 產摘要與分類,輸出 markdown。
設定見 config.json (從 config.example.json 複製修改)。
"""

import os
import re
import sys
import json
from pathlib import Path

import win32com.client
import pythoncom
from anthropic import Anthropic


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
        example = APP_DIR / "config.example.json"
        if example.exists():
            print(f"找不到 config.json,請複製 config.example.json 為 config.json 並修改。")
        else:
            print(f"找不到 config.json,路徑: {CONFIG_PATH}")
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
    raise RuntimeError(f"找不到 PST: {pst_name}。請確認 Outlook 已開啟並掛載此 PST。")


def safe_filename(s, maxlen=80):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", s)
    return s[:maxlen].strip() or "untitled"


def summarize(client, mail, model, max_chars):
    body = (mail.Body or "")[:max_chars]
    prompt = SUMMARY_PROMPT.format(
        sender=mail.SenderName or "(unknown)",
        subject=mail.Subject or "(no subject)",
        date=str(mail.ReceivedTime),
        body=body,
    )
    resp = client.messages.create(
        model=model,
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    )
    text = resp.content[0].text.strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.MULTILINE)
    return json.loads(text)


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


def main():
    cfg = load_config()

    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("錯誤:請先設定環境變數 ANTHROPIC_API_KEY")
        print('PowerShell: [Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY","sk-ant-...","User")')
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

    client = Anthropic()
    processed = load_state()
    items = list(inbox.Items)
    items.sort(key=lambda m: m.ReceivedTime, reverse=True)

    new_count = 0
    for mail in items:
        try:
            entry_id = mail.EntryID
        except Exception:
            continue
        if entry_id in processed:
            continue
        if getattr(mail, "Class", 0) != 43:
            continue
        try:
            print(f"[{new_count+1}] {mail.SenderName} - {(mail.Subject or '')[:50]}")
            result = summarize(client, mail, cfg["model"], cfg["max_body_chars"])
            path = write_markdown(mail, result, output_dir)
            print(f"    -> {path.name} [{result.get('category')}]")
            processed.add(entry_id)
            new_count += 1
            if new_count % 5 == 0:
                save_state(processed)
        except Exception as e:
            print(f"    !! 失敗: {e}")

    save_state(processed)
    print(f"\n完成。新處理 {new_count} 封。輸出: {output_dir}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n錯誤: {e}")
        input("按 Enter 結束...")
