#Requires -Version 5.1
<#
.SYNOPSIS
    Builds WinUpdateChecker.exe from WinUpdateChecker.ps1 using ps2exe.
.DESCRIPTION
    Produces a windowed (-noConsole), self-elevating (-requireAdmin) executable
    with embedded version metadata and the app icon. The .exe is written to
    dist\ (gitignored) and published via GitHub releases, not committed.
.NOTES
    Installs the ps2exe module from PSGallery (CurrentUser scope) if missing.
#>
[CmdletBinding()]
param(
    [string]$InputFile,
    [string]$OutputFile,
    [string]$IconFile,
    [string]$Version = '1.0.0.0'
)

$ErrorActionPreference = 'Stop'

# Resolve the repo base robustly ($PSScriptRoot can be empty depending on how
# the script is invoked); callers may also override any path explicitly.
$base = if ($PSScriptRoot) { $PSScriptRoot }
        elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
        else { (Get-Location).Path }
if (-not $InputFile)  { $InputFile  = Join-Path $base 'WinUpdateChecker.ps1' }
if (-not $OutputFile) { $OutputFile = Join-Path $base 'dist\WinUpdateChecker.exe' }
if (-not $IconFile)   { $IconFile   = Join-Path $base 'assets\icon.ico' }
[Net.ServicePointManager]::SecurityProtocol =
    [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host 'Installing ps2exe (CurrentUser scope) from PSGallery...'
    Install-Module -Name ps2exe -Scope CurrentUser -Force
}
Import-Module ps2exe

$outDir = Split-Path $OutputFile -Parent
if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

Write-Host "Compiling`n  $InputFile`n-> $OutputFile"
Invoke-ps2exe -inputFile $InputFile -outputFile $OutputFile -iconFile $IconFile `
    -title     'Windows Update Checker' `
    -product   'Windows Update Checker' `
    -company   'Pablo Cartagena / TEC Building Systems LLC' `
    -copyright 'Copyright (c) 2026 Pablo Cartagena / TEC Building Systems LLC' `
    -version   $Version `
    -requireAdmin `
    -noConsole

if (-not (Test-Path $OutputFile)) { throw "Build failed: $OutputFile was not produced." }

$vi = (Get-Item $OutputFile).VersionInfo
Write-Host "`nBuilt OK: $OutputFile"
Write-Host ("  Product:   {0}" -f $vi.ProductName)
Write-Host ("  File desc: {0}" -f $vi.FileDescription)
Write-Host ("  Version:   {0}" -f $vi.FileVersion)
Write-Host ("  Company:   {0}" -f $vi.CompanyName)
Write-Host ("  Copyright: {0}" -f $vi.LegalCopyright)
