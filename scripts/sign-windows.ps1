param(
    [Parameter(Mandatory = $true)]
    [string]$CertificatePath,

    [Parameter(Mandatory = $true)]
    [string]$CertificatePassword,

    [Parameter(Mandatory = $true)]
    [string[]]$InputPaths,

    [string]$TimestampUrl = "http://timestamp.digicert.com",
    [string]$Description = "AutoWeChat",
    [string]$FileDigestAlgorithm = "SHA256"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-SignToolPath {
    $command = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $kitRoot = "C:\Program Files (x86)\Windows Kits\10\bin"
    if (Test-Path $kitRoot) {
        $x64Candidates = Get-ChildItem $kitRoot -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match "\\x64\\signtool\.exe$" } |
            Sort-Object FullName -Descending
        if ($x64Candidates) {
            return $x64Candidates[0].FullName
        }

        $allCandidates = Get-ChildItem $kitRoot -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending
        if ($allCandidates) {
            return $allCandidates[0].FullName
        }
    }

    throw "未找到 signtool.exe，请确认 Windows SDK 已安装。"
}

function Resolve-SignTargets {
    param([string[]]$Paths)

    $extensions = @(".exe", ".dll", ".pyd")
    $resolved = New-Object System.Collections.Generic.List[string]

    foreach ($inputPath in $Paths) {
        if ([string]::IsNullOrWhiteSpace($inputPath)) {
            continue
        }

        $fullPath = [System.IO.Path]::GetFullPath($inputPath)
        if (-not (Test-Path $fullPath)) {
            throw "待签名路径不存在: $fullPath"
        }

        $item = Get-Item $fullPath
        if ($item.PSIsContainer) {
            Get-ChildItem $fullPath -Recurse -File | Where-Object { $extensions -contains $_.Extension.ToLowerInvariant() } |
                Sort-Object FullName |
                ForEach-Object { $resolved.Add($_.FullName) }
            continue
        }

        if ($extensions -contains $item.Extension.ToLowerInvariant()) {
            $resolved.Add($item.FullName)
        }
    }

    return $resolved | Select-Object -Unique
}

$certificateFullPath = [System.IO.Path]::GetFullPath($CertificatePath)
if (-not (Test-Path $certificateFullPath)) {
    throw "证书文件不存在: $certificateFullPath"
}

$targets = Resolve-SignTargets -Paths $InputPaths
if (-not $targets -or $targets.Count -eq 0) {
    throw "没有找到可签名的 exe/dll/pyd 文件。"
}

$signTool = Get-SignToolPath
Write-Host "Using SignTool: $signTool"
Write-Host "Signing $($targets.Count) file(s)..."

foreach ($target in $targets) {
    Write-Host "Signing: $target"
    & $signTool sign `
        /fd $FileDigestAlgorithm `
        /td $FileDigestAlgorithm `
        /tr $TimestampUrl `
        /f $certificateFullPath `
        /p $CertificatePassword `
        /d $Description `
        $target

    if ($LASTEXITCODE -ne 0) {
        throw "签名失败: $target"
    }

    & $signTool verify /pa $target
    if ($LASTEXITCODE -ne 0) {
        throw "签名校验失败: $target"
    }
}

Write-Host "All signing operations completed successfully."
