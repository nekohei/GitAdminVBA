# push-all.ps1
# sekinekh と nekohei の両アカウントに git push するスクリプト
# nekohei の PAT はレジストリから自動取得

$repoPath = $PSScriptRoot
$branch = "main"

# nekohei の PAT をレジストリから取得
$pat = (Get-ItemProperty "HKCU:\Software\VB and VBA Program Settings\GitHub\nekohei").Classic
if (-not $pat) {
    Write-Error "nekohei の PAT がレジストリに見つかりません。"
    exit 1
}

# sekinekh へ push
Write-Host "sekinekh へ push 中..." -ForegroundColor Cyan
git -C $repoPath push https://github.com/sekinekh/GitAdminVBA.git $branch
if ($LASTEXITCODE -ne 0) {
    Write-Error "sekinekh への push に失敗しました。"
    exit 1
}

# nekohei へ push（PAT をURLに埋め込んで使用）
Write-Host "nekohei へ push 中..." -ForegroundColor Cyan
git -C $repoPath push "https://$pat@github.com/nekohei/GitAdminVBA.git" $branch
if ($LASTEXITCODE -ne 0) {
    Write-Error "nekohei への push に失敗しました。"
    exit 1
}

Write-Host "両アカウントへの push が完了しました。" -ForegroundColor Green
