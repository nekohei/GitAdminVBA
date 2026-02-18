param(
    [string]$Path = ".",
    [string[]]$Extensions = @("*.bas", "*.cls", "*.frm", "*.dcm")
)

function Get-BomType {
    param([byte[]]$Bytes)
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { return 'UTF-8 BOM' }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) { return 'UTF-16 LE BOM' }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) { return 'UTF-16 BE BOM' }
    return 'None'
}

function Get-NewlineStats {
    param([byte[]]$Bytes)
    $crlf = 0; $lf = 0
    $i = 0
    while ($i -lt $Bytes.Length) {
        if ($Bytes[$i] -eq 0x0D -and ($i + 1) -lt $Bytes.Length -and $Bytes[$i+1] -eq 0x0A) {
            $crlf++
            $i += 2
            continue
        }
        if ($Bytes[$i] -eq 0x0A) { $lf++ }
        $i++
    }
    [pscustomobject]@{ CRLF = $crlf; LF = $lf }
}

function Guess-Encoding {
    param([byte[]]$Bytes)
    $bom = Get-BomType -Bytes $Bytes
    if ($bom -ne 'None') { return $bom }
    try {
        $sUtf8 = [System.Text.Encoding]::UTF8.GetString($Bytes)
        $rUtf8 = [System.Text.Encoding]::UTF8.GetBytes($sUtf8)
        if ($rUtf8.Length -eq $Bytes.Length) {
            $same = $true
            for ($i=0; $i -lt $Bytes.Length; $i++) { if ($Bytes[$i] -ne $rUtf8[$i]) { $same = $false; break } }
            if ($same) { return 'UTF-8' }
        }
    } catch { }
    try {
        $enc932 = [System.Text.Encoding]::GetEncoding(932)
        $s932 = $enc932.GetString($Bytes)
        $r932 = $enc932.GetBytes($s932)
        if ($r932.Length -eq $Bytes.Length) {
            $same = $true
            for ($i=0; $i -lt $Bytes.Length; $i++) { if ($Bytes[$i] -ne $r932[$i]) { $same = $false; break } }
            if ($same) { return 'CP932' }
        }
    } catch { }
    return 'Unknown'
}

$results = @()
foreach ($ext in $Extensions) {
    $files = Get-ChildItem -Path $Path -Filter $ext -Recurse -File -ErrorAction SilentlyContinue
    foreach ($f in $files) {
        $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
        $nl = Get-NewlineStats -Bytes $bytes
        $enc = Guess-Encoding -Bytes $bytes
        $crlfOk = ($nl.LF -eq 0)
        $needsConv = ($enc -ne 'CP932' -or -not $crlfOk)
        $results += [pscustomobject]@{
            Path = $f.FullName
            Encoding = $enc
            CRLF_OK = $crlfOk
            NeedsConversion = $needsConv
            CRLF_Count = $nl.CRLf
            LF_Alone_Count = $nl.LF
        }
    }
}

# 出力を見やすい順で：変換が必要なもの→不要なもの
$results | Sort-Object -Property @{Expression='NeedsConversion';Descending=$true}, Path | Format-Table -AutoSize

