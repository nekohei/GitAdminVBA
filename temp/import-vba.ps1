# import-vba.ps1
# src/ の VBAモジュールを Git管理.xlam に反映する
# .dcm（ドキュメントモジュール）は今回の改修対象外のためスキップ

$binPath     = "c:\Users\1784\claude\VBA\Excel\GitAdminVBA\bin\Git管理.xlam"
$historyDir  = "c:\Users\1784\claude\VBA\Excel\GitAdminVBA\bin\history"
$archivePath = "$historyDir\Git管理_20260218.xlam"
$srcPath     = "c:\Users\1784\claude\VBA\Excel\GitAdminVBA\src\\"

# ---- 1. 改修前バージョンをアーカイブ ----
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

    # 標準モジュール・クラスモジュールを削除（Type 1=標準, 2=クラス）
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

    # .dcm（ドキュメントモジュール）は今回の改修対象外のためスキップ
    Write-Host ".dcm はスキップ（ThisWorkbook / Sheet1 は変更なし）" -ForegroundColor Yellow

    $wb.Save()
    $wb.Close($false)
    Write-Host "`nインポート完了: $binPath" -ForegroundColor Green

} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
