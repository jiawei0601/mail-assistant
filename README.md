# MAIL 助手

讀 Outlook PST 收件匣,用 Claude 自動分類 + 產摘要 + 抽待辦,輸出 markdown。

## 下載 (一般使用者)

到 [Releases](../../releases) 下載最新版,解壓縮後:

1. 把 `config.example.json` 複製成 `config.json`,改 `pst_display_name` 為你的 PST 顯示名稱
2. 設環境變數 `ANTHROPIC_API_KEY`(去 https://console.anthropic.com/ 拿)
   ```powershell
   [Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY","sk-ant-...","User")
   ```
3. Outlook 開著 + PST 掛載
4. 雙擊 `MailAssistant.exe`

輸出在同層 `output\YYYY-MM-DD\` 資料夾。

## 從原始碼跑

```bash
pip install -r requirements.txt
cp config.example.json config.json
python mail_assistant.py
```

## 自己打包 EXE

```bash
pip install pyinstaller
build.bat
```

EXE 產出在 `dist\MailAssistant.exe`。

## 設定檔說明 (config.json)

| 欄位 | 說明 |
|---|---|
| `pst_display_name` | Outlook 左側 PST 顯示名稱(部分比對,不分大小寫) |
| `folder_name` | 要處理的資料夾,預設 Inbox |
| `output_dir` | 輸出路徑,相對或絕對 |
| `model` | Claude 模型,預設 `claude-haiku-4-5-20251001`(便宜快) |
| `max_body_chars` | 信件內文截斷長度 |

## 排程自動跑

Windows 工作排程器 → 新增工作 → 動作執行 `MailAssistant.exe`,觸發程序每 1 小時。

## 限制

- 僅 Windows(需 Outlook COM)
- Outlook 必須開啟
- 第一次跑會跳安全警告,允許即可
