# UTF-8からCP932（Shift-JIS）+ CRLF に変換するスクリプト
# VBAファイル用エンコーディング変換（OSに依存せずCRLFを強制）

param(
    [string]$Path = ".",
    [string[]]$Extensions = @("*.bas", "*.cls", "*.frm", "*.dcm")
)

Write-Host "VBAファイルのエンコーディング変換を開始します..." -ForegroundColor Green

foreach ($ext in $Extensions) {
    $files = Get-ChildItem -Path $Path -Filter $ext -Recurse

    foreach ($file in $files) {
        try {
            Write-Host "変換中: $($file.FullName)" -ForegroundColor Yellow

            # UTF-8として行単位で読み込み、CRLFで結合して末尾にもCRLFを付与
            $lines = Get-Content -Path $file.FullName -Encoding UTF8
            $content = ($lines -join "`r`n") + "`r`n"

            # CP932（Shift-JIS, codepage 932）でバイト列にして書き込み
            $bytes = [System.Text.Encoding]::GetEncoding(932).GetBytes($content)
            [System.IO.File]::WriteAllBytes($file.FullName, $bytes)

            Write-Host "完了: $($file.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "エラー: $($file.Name) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Write-Host "VBAファイルのエンコーディング変換が完了しました。" -ForegroundColor Green
