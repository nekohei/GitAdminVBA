# UTF-8からCP932（Shift-JIS）に変換するスクリプト
# VBAファイル用エンコーディング変換

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

            # UTF-8でファイルを読み込み
            $content = Get-Content -Path $file.FullName -Encoding UTF8 -Raw

            # CP932（Shift-JIS）で保存
            $content | Out-File -FilePath $file.FullName -Encoding Default -NoNewline

            Write-Host "完了: $($file.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "エラー: $($file.Name) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Write-Host "VBAファイルのエンコーディング変換が完了しました。" -ForegroundColor Green