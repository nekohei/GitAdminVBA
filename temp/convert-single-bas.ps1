# 単一VBAファイルをCP932・CRLF改行に変換
param([string]$FilePath)

# UTF-8でファイルを行ごとに読み込み、改行を制御
$lines = Get-Content -Path $FilePath -Encoding UTF8
$content = ($lines -join "`r`n") + "`r`n"
$bytes = [System.Text.Encoding]::GetEncoding(932).GetBytes($content)
[System.IO.File]::WriteAllBytes($FilePath, $bytes)
Write-Host "Converted: $FilePath" -ForegroundColor Green