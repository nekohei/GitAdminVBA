$p = Join-Path $env:APPDATA "Microsoft\AddIns\Git管理.xlam"
if (Test-Path $p) {
    $f = Get-Item $p
    Write-Output ("Found: " + $f.FullName)
    Write-Output ("Size : " + $f.Length + " bytes")
    Write-Output ("Date : " + $f.LastWriteTime)
} else {
    Write-Output ("Not found: " + $p)
}
