# 1C Enterprise Automation Scripts

Daily sales reports via 1C COM → Telegram

## Structure

```
server1/          ← Сервер 1 (Гонгадзе / Vlada)
  1c_daily.ps1    ← daily Telegram report, run via SysWOW64 PowerShell (32-bit)
server2/          ← Сервер 2 (Гречко / Vlada2)
  1c_daily.ps1    ← daily Telegram report
retail_report/    ← Python COM report (Excel output)
  retail_report.py
  run.bat
  config/
    retail_config.json
```

## Key technical notes

- Must run with 32-bit PowerShell: `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`
- 1C COM: `New-Object -ComObject "V83.COMConnector"`
- Number format fix: `-replace [char]0xA0," " -replace ",","."`  (Ukrainian locale thousands/decimal)
- Query result row access: `$result.Получить($i, 0)` returns full row as string "Касса 2 95219.8"
- Schedule: daily via Windows Task Scheduler
