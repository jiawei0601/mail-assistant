@echo off
REM 打包 EXE - 需先 pip install pyinstaller
pyinstaller --onefile --name MailAssistant ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import anthropic ^
  mail_assistant.py

echo.
echo EXE 在 dist\MailAssistant.exe
echo 記得把 config.example.json 複製成 config.json 放在 EXE 同層
pause
