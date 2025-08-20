#requires -Version 3.0
[CmdletBinding()]
param(
    [string]$Version = "1.14.0",
    [ValidateSet("gpu", "cpu")]
    [string]$Type = "gpu",
    [switch]$NoAria2c,
    [string]$BinPath = "$PSScriptRoot\..\..\..\bin",
    [Alias("h", "help")]
    [switch]$MyHelp,
    [switch]$CopyFromLibs
)

<#
.SYNOPSIS
    Downloads and extracts ONNX Runtime DLLs, then copies them to the project's bin directory.

.DESCRIPTION
    This script downloads the specified version of ONNX Runtime (CPU or GPU) for Windows,
    extracts the DLLs, and copies them to the openFrameworks project's bin directory.
    If aria2c is available, it will be used for faster parallel download; otherwise, Invoke-WebRequest is used.
    You can force not to use aria2c with the -NoAria2c option.
    The bin directory can be specified with -BinPath.
    If -CopyFromLibs is specified, copies DLLs from libs\onnxruntime\lib\vs\x64\ to BinPath (no download or extract).
    All operations are performed in a secure temporary directory, which is cleaned up even on error.

.PARAMETER Version
    The ONNX Runtime version to download. Default is "1.14.0".

.PARAMETER Type
    "gpu" or "cpu". Default is "gpu".

.PARAMETER NoAria2c
    If specified, aria2c will not be used even if available.

.PARAMETER BinPath
    The path to the bin directory. Default is the openFrameworks project bin directory.

.PARAMETER CopyFromLibs
    If specified, copies DLLs from libs\onnxruntime\lib\vs\x64\ to BinPath (no download or extract).

.PARAMETER MyHelp
    Shows this help message.

.EXAMPLE
    .\download_and_copy_onnxruntime.ps1 -CopyFromLibs
    Copies DLLs from libs\onnxruntime\lib\vs\x64\ to bin, no download or extract.
#>

if ($MyHelp) {
    Get-Help $MyInvocation.MyCommand.Path -Full
    exit 0
}

# Resolve and check bin path
$resolvedBinPath = Resolve-Path -Path $BinPath -ErrorAction SilentlyContinue
if (!$resolvedBinPath) {
    Write-Error "bin directory not found: $BinPath. Please specify app bin path using -BinPath"
    exit 1
}
$binPath = $resolvedBinPath.Path

$binPath = [System.IO.Path]::GetFullPath($binPath)
if ($binPath -eq "C:\" -or $binPath -eq "C:/") {
    Write-Error "Refusing to copy to root directory: $binPath. Please specify app bin path using -BinPath"
    exit 1
}

if ($Type -eq "gpu") {
    $onnxDirName = "onnxruntime-win-x64-gpu-$Version"
} else {
    $onnxDirName = "onnxruntime-win-x64-$Version"
}

# CopyFromLibs mode: copy DLLs from libs\onnxruntime\lib\vs\x64\ to BinPath
if ($CopyFromLibs) {
    
    $addonRoot = Split-Path -Parent $PSScriptRoot
    $libRoot = Join-Path $addonRoot "libs\onnxruntime\lib"
    $libX64 = Join-Path $libRoot "vs\x64"

    # Copy all DLLs from vs\x64 to bin
    $dlls = Get-ChildItem -Path $libX64\$onnxDirName\lib -Filter *.dll -ErrorAction SilentlyContinue
    if ($dlls.Count -eq 0) {
        Write-Warning "No DLLs found in $libX64\$onnxDirName\lib. Please download them first without using -CopyFromLibs"
    } else {
        Write-Host "Copying DLLs from $libX64\$onnxDirName\lib to $binPath ..."
        Copy-Item "$libX64\$onnxDirName\lib\*.dll" -Destination $binPath -Force
        Write-Host "Done."
    }
    exit 0
}

# Create a secure temporary working directory and file
$tmpRoot = [System.IO.Path]::GetTempPath()
$tmpDir = Join-Path $tmpRoot ([guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $tmpDir | Out-Null

$tmpZip = [System.IO.Path]::GetTempFileName()
$tmpZipZip = [System.IO.Path]::ChangeExtension($tmpZip, ".zip")
Rename-Item -Path $tmpZip -NewName $tmpZipZip
$downloadPath = $tmpZipZip
$extractPath = Join-Path $tmpDir "extract"

# Prepare file name and directory for aria2c
$downloadFileName = [System.IO.Path]::GetFileName($downloadPath)
$downloadDir = [System.IO.Path]::GetDirectoryName($downloadPath)

if ($Type -eq "gpu") {
    $zipName = "onnxruntime-win-x64-gpu-$Version.zip"
} else {
    $zipName = "onnxruntime-win-x64-$Version.zip"
}
$baseUrl = "https://github.com/microsoft/onnxruntime/releases/download/v$Version"
$downloadUrl = "$baseUrl/$zipName"

# Check for aria2c
$aria2c = Get-Command "aria2c.exe" -ErrorAction SilentlyContinue

try {
    Remove-Item $downloadPath -Force -ErrorAction SilentlyContinue

    Write-Host "Downloading $downloadUrl ..."
    if ($aria2c -and -not $NoAria2c) {
        Write-Host "aria2c found. Downloading with aria2c (4 connections)..."
        & $aria2c.Path "-x" "4" "-d" $downloadDir "-o" $downloadFileName $downloadUrl
        $aria2file = "$downloadPath.aria2"
        while (Test-Path $aria2file) { Start-Sleep -Seconds 1 }
    } else {
        if ($NoAria2c) { Write-Host "Forcing Invoke-WebRequest (aria2c not used by user option)." }
        else { Write-Host "aria2c not found. Downloading with Invoke-WebRequest..." }
        Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath
    }

    Start-Sleep -Seconds 1
    Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue

    if (!(Test-Path $downloadPath)) {
        Write-Error "Download failed: $downloadPath not found."
        exit 1
    }

    Write-Host "Extracting $zipName ..."
    Expand-Archive -Path $downloadPath -DestinationPath $extractPath -Force

    # Find the extracted folder (should be onnxruntime-win-x64(-gpu)-$Version)
    $onnxDir = Get-ChildItem -Path $extractPath | Where-Object { $_.PSIsContainer } | Select-Object -First 1
    if (!$onnxDir) {
        Write-Error "Extracted directory not found."
        exit 1
    }

    $dllSource = Join-Path $onnxDir.FullName "lib"
    if (!(Test-Path $dllSource)) {
        Write-Error "DLL source directory not found: $dllSource"
        exit 1
    }
    $dlls = Get-ChildItem -Path $dllSource -Filter *.dll -ErrorAction SilentlyContinue
    if ($dlls.Count -eq 0) {
        Write-Error "No DLLs found in $dllSource"
        exit 1
    }

    # Check and create libs\onnxruntime\lib\vs\x64
    $addonRoot = Split-Path -Parent $PSScriptRoot
    $libRoot = Join-Path $addonRoot "libs\onnxruntime\lib"
    $libX64 = Join-Path $libRoot "vs\x64"
    if (!(Test-Path $libRoot)) {
        Write-Error "libs\onnxruntime\lib does not exist. Please check your addon structure."
        exit 1
    }
    if (!(Test-Path $libX64)) {
        New-Item -ItemType Directory -Path $libX64 | Out-Null
        Write-Host "Created directory: $libX64"
    }

    Write-Host "Copying DLLs from $dllSource to $binPath ..."
    Copy-Item "$dllSource\*.dll" -Destination $binPath -Force

    $onnxDirFull = $onnxDir.FullName
    Write-Host "Copying $onnxDirFull into $libX64 ..."
    Copy-Item "$onnxDirFull" -Destination $libX64 -Force -Recurse

    Write-Host "Done."
}
finally {
    # Cleanup temporary files and directories
    if (Test-Path $downloadPath) { Remove-Item $downloadPath -Force -ErrorAction SilentlyContinue }
    if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $tmpDir) { Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
}