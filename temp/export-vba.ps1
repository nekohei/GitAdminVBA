# export-vba.ps1
# Git管理.xlam から VBAモジュールを src/ にエクスポートする

$xlamPath = "c:\Users\1784\claude\VBA\Excel\GitAdminVBA\bin\Git管理.xlam"
$srcPath  = "c:\Users\1784\claude\VBA\Excel\GitAdminVBA\src\"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($xlamPath)
    $proj = $wb.VBProject

    foreach ($comp in $proj.VBComponents) {
        $type = $comp.Type
        # 1=標準モジュール 2=クラス 3=UserForm 100=ドキュメント
        if ($type -eq 100) {
            $ext = ".dcm"
        } elseif ($type -eq 2) {
            $ext = ".cls"
        } elseif ($type -eq 3) {
            $ext = ".frm"
        } else {
            $ext = ".bas"
        }
        $outPath = $srcPath + $comp.Name + $ext
        $comp.Export($outPath)
        Write-Host ("エクスポート: " + $comp.Name + $ext)
    }
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Write-Host "完了" -ForegroundColor Green
