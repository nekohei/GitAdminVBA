# import-vba.ps1
# src/ の VBAモジュールを bin/Git管理.xlam に反映する
# .dcm（ドキュメントモジュール）はスキップ
#
# 【前提】Excel をすべて終了した状態で実行すること
# （アドインが読み込まれていない状態でないと同名ファイルを開けないため）

$binPath    = "c:\Users\1784\claude\VBA\Excel\GitAdminVBA\bin\Git管理.xlam"
$historyDir = "c:\Users\1784\claude\VBA\Excel\GitAdminVBA\bin\history"
$srcPath    = "c:\Users\1784\claude\VBA\Excel\GitAdminVBA\src\"

# ---- 0. Excel プロセス確認 ----
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Write-Host "Excel が起動中です。すべての Excel を閉じてから実行してください。" -ForegroundColor Red
    exit 1
}

# ---- 1. 改修前バージョンをアーカイブ（日付付き、上書きなし） ----
$today       = Get-Date -Format "yyyyMMdd"
$archivePath = "$historyDir\Git管理_$today.xlam"
if (-not (Test-Path $archivePath)) {
    Copy-Item $binPath $archivePath
    Write-Host "アーカイブ完了: $archivePath" -ForegroundColor Cyan
} else {
    Write-Host "アーカイブ済み: $archivePath" -ForegroundColor Yellow
}

# ---- 2. Excel COM でインポート ----
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb   = $excel.Workbooks.Open($binPath)
    $proj = $wb.VBProject

    # 標準モジュール（Type=1）・クラスモジュール（Type=2）を削除
    $toRemove = @()
    foreach ($comp in $proj.VBComponents) {
        if ($comp.Type -eq 1 -or $comp.Type -eq 2) {
            $toRemove += $comp.Name
        }
    }
    foreach ($name in $toRemove) {
        $proj.VBComponents.Remove($proj.VBComponents.Item($name))
        Write-Host "削除: $name" -ForegroundColor DarkGray
    }

    # .bas をインポート
    Get-ChildItem $srcPath -Filter "*.bas" | ForEach-Object {
        $proj.VBComponents.Import($_.FullName) | Out-Null
        Write-Host "インポート: $($_.Name)" -ForegroundColor Green
    }

    # .cls をインポート
    Get-ChildItem $srcPath -Filter "*.cls" | ForEach-Object {
        $proj.VBComponents.Import($_.FullName) | Out-Null
        Write-Host "インポート: $($_.Name)" -ForegroundColor Green
    }

    Write-Host ".dcm はスキップ（ThisWorkbook / Sheet1 は変更なし）" -ForegroundColor Yellow

    $wb.Save()
    $wb.Close($false)
    Write-Host "`nインポート完了: $binPath" -ForegroundColor Green

} finally {
    $excel.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
