# 1C Daily Sales - yesterday
# Run via: C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe

# --- Settings ---
$SERVER   = "localhost"
$BASE     = "Vlada"
$USER     = "Админ"
$PASSWORD = "77556670"
$KASSAS   = @("Касса 2", "Касса 3", "Касса 4", "Касса 5")

$TG_TOKEN   = "8503330594:AAHAPN4KemJ6RAZS5ctAXbFhHKHHfIMVsMU"
$TG_CHAT_ID = "581120233"

$yesterday  = (Get-Date).AddDays(-1)
$FROM_DATE  = "$($yesterday.Year),$($yesterday.Month),$($yesterday.Day),0,0,0"
$TO_DATE    = "$($yesterday.Year),$($yesterday.Month),$($yesterday.Day),23,59,59"
$dateStr    = $yesterday.ToString("dd.MM.yyyy")

# --- Telegram ---
function Send-Telegram($text) {
    try {
        $utf8     = [System.Text.Encoding]::UTF8
        $escaped  = $text.Replace('\', '\\').Replace('"', '\"').Replace("`n", '\n').Replace("`r", '')
        $bodyStr  = '{"chat_id":"' + $TG_CHAT_ID + '","text":"' + $escaped + '"}'
        $bodyBytes = $utf8.GetBytes($bodyStr)
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("Content-Type", "application/json; charset=utf-8")
        $wc.UploadData("https://api.telegram.org/bot$TG_TOKEN/sendMessage", "POST", $bodyBytes) | Out-Null
    } catch { Write-Host "Telegram error: $_" }
}

# --- Connect ---
try {
    $connector = New-Object -ComObject "V83.COMConnector"
    $conn = $connector.Connect("Srvr=$SERVER;Ref=$BASE;Usr=$USER;Pwd=$PASSWORD")
    if ($null -eq $conn) { Write-Host "ERROR: Connection failed"; exit 1 }
} catch {
    Write-Host "ERROR: $_"; exit 1
}

# --- Query ---
$kassaFilter = ($KASSAS | ForEach-Object { '"' + $_ + '"' }) -join ", "
$queryText = "ВЫБРАТЬ Отчет.КассаККМ.Наименование КАК Касса, СУММА(Отчет.СуммаДокумента) КАК Сумма " +
             "ИЗ Документ.ОтчетОРозничныхПродажах КАК Отчет " +
             "ГДЕ Отчет.Дата >= ДАТАВРЕМЯ($FROM_DATE) " +
             "И Отчет.Дата <= ДАТАВРЕМЯ($TO_DATE) " +
             "И Отчет.КассаККМ.Наименование В ($kassaFilter) " +
             "И Отчет.Проведен = ИСТИНА " +
             "СГРУППИРОВАТЬ ПО Отчет.КассаККМ.Наименование"

try {
    $query = $conn.NewObject("Запрос")
    $query.Текст = $queryText
    $result = $query.Выполнить().Выгрузить()
} catch {
    Write-Host "ERROR: $_"; exit 1
}

# --- Output ---
Write-Host ""
Write-Host ("=" * 42)
Write-Host "  ПРОДАЖІ $dateStr"
Write-Host ("=" * 42)

$total  = 0
$tgMsg  = "Гонгадзе ПРОДАЖІ $dateStr`n"
$tgMsg += "--------------------`n"

for ($i = 0; $i -lt $result.Количество(); $i++) {
    $line   = "$($result.Получить($i, 0))"
    $parts  = $line -split " "
    $kassa  = ($parts[0..($parts.Length-2)] -join " ").TrimEnd()
    $sum    = [decimal]$parts[-1]
    $total += $sum
    $sumStr = ("{0:N2}" -f $sum) -replace [char]0xA0," " -replace ",","."
    Write-Host ("{0,-15} {1,15:N2} грн" -f $kassa, $sum)
    $tgMsg += "$kassa`: $sumStr грн`n"
}

Write-Host ("-" * 42)
Write-Host ("{0,-15} {1,15:N2} грн" -f "РАЗОМ", $total)
Write-Host ("=" * 42)

$totalStr = ("{0:N2}" -f $total) -replace [char]0xA0," " -replace ",","."
$tgMsg += "--------------------`n"
$tgMsg += "РАЗОМ: $totalStr грн"

Send-Telegram $tgMsg
Write-Host "Telegram: відправлено."
