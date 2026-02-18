# restore-addin.ps1
# テスト終了用：バックアップからアドインフォルダを元に戻す
#
# 【手順】
#   1. Excel をすべて終了する
#   2. このスクリプトを実行する
#   3. Excel を起動して元のバージョンが読み込まれることを確認する

$addinPath  = "C:\Users\1784\AppData\Roaming\Microsoft\AddIns\Git管理.xlam"
$historyDir = "c:\Users\1784\claude\VBA\Excel\GitAdminVBA\bin\history"
$backupPath = "$historyDir\Git管理_addin_backup.xlam"

# ---- Excel プロセス確認 ----
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Write-Host "Excel が起動中です。すべての Excel を閉じてから実行してください。" -ForegroundColor Red
    exit 1
}

# ---- バックアップの存在確認 ----
if (-not (Test-Path $backupPath)) {
    Write-Host "バックアップが見つかりません: $backupPath" -ForegroundColor Red
    Write-Host "deploy-test.ps1 を実行した形跡がありません。" -ForegroundColor Red
    exit 1
}

# ---- 復元 ----
Copy-Item $backupPath $addinPath -Force
Write-Host "復元完了: $addinPath" -ForegroundColor Green
Write-Host "`nExcel を起動して元のバージョンが読み込まれることを確認してください。" -ForegroundColor Yellow
