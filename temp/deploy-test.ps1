# deploy-test.ps1
# テスト開始用：bin/Git管理.xlam をアドインフォルダに上書きする
# 上書き前に現在のアドインを bin/history/ にバックアップする
#
# 【手順】
#   1. Excel をすべて終了する
#   2. このスクリプトを実行する
#   3. Excel を起動して動作確認する
#   4. テスト終了後は restore-addin.ps1 で元に戻す

$addinPath  = "C:\Users\1784\AppData\Roaming\Microsoft\AddIns\Git管理.xlam"
$binPath    = "c:\Users\1784\claude\VBA\Excel\GitAdminVBA\bin\Git管理.xlam"
$historyDir = "c:\Users\1784\claude\VBA\Excel\GitAdminVBA\bin\history"
$backupPath = "$historyDir\Git管理_addin_backup.xlam"

# ---- Excel プロセス確認 ----
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Write-Host "Excel が起動中です。すべての Excel を閉じてから実行してください。" -ForegroundColor Red
    exit 1
}

# ---- バックアップ（上書きあり：常に最新の稼働バージョンを保持） ----
Copy-Item $addinPath $backupPath -Force
Write-Host "バックアップ完了: $backupPath" -ForegroundColor Cyan

# ---- アドインフォルダに上書きコピー ----
Copy-Item $binPath $addinPath -Force
Write-Host "デプロイ完了: $addinPath" -ForegroundColor Green
Write-Host "`nExcel を起動してテストしてください。" -ForegroundColor Yellow
Write-Host "テスト終了後は restore-addin.ps1 を実行して元に戻してください。" -ForegroundColor Yellow
