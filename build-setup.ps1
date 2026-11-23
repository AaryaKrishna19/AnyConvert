#Requires -Version 5.1

$ErrorActionPreference = 'Stop'
$project = $PSScriptRoot
$builder = Join-Path $project 'build-winrar-style-setup.ps1'

if (!(Test-Path -LiteralPath $builder -PathType Leaf)) {
    throw "Missing build script: $builder"
}

# Find Inno Setup, including installations outside the two original paths.
$candidates = @()
$command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
if ($command) {
    $candidates += $command.Source
}

$uninstallKeys = @(
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

foreach ($entry in Get-ItemProperty $uninstallKeys -ErrorAction SilentlyContinue) {
    if ($entry.DisplayName -like 'Inno Setup*' -and $entry.InstallLocation) {
        $candidates += Join-Path $entry.InstallLocation 'ISCC.exe'
    }
}

foreach ($folder in @(
    "${env:ProgramFiles(x86)}\Inno Setup 6"
    "$env:ProgramFiles\Inno Setup 6"
    "$env:LOCALAPPDATA\Programs\Inno Setup 6"
)) {
    $candidates += Join-Path $folder 'ISCC.exe'
}

$inno = $candidates |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    Select-Object -First 1

if (!$inno) {
    throw 'Install Inno Setup 6 from https://jrsoftware.org/isdl.php and run this script again.'
}
Write-Host "Found Inno Setup: $inno"

# Use Microsoft's locator to find an installation with the x64 C++ compiler.
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (!(Test-Path -LiteralPath $vswhere)) {
    throw 'Install Microsoft Build Tools with Desktop development with C++ and a Windows 11 SDK, then run this script again.'
}

$vsPath = & $vswhere -latest -products '*' `
    -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
    -property installationPath

if (!$vsPath) {
    throw 'In Visual Studio Installer, select Modify and install Desktop development with C++ and a Windows 11 SDK.'
}

$devCmd = Join-Path $vsPath 'Common7\Tools\VsDevCmd.bat'
if (!(Test-Path -LiteralPath $devCmd)) {
    throw "Missing Visual Studio environment script: $devCmd"
}

# The earlier builder chooses a fixed Inno path. Give this build a temporary
# copy that uses the compiler we actually found.
$source = Get-Content -LiteralPath $builder -Raw
$start = $source.IndexOf('$inno = Join-Path')
$end = $source.IndexOf('foreach ($file in @(')

if ($start -lt 0 -or $end -le $start) {
    throw 'The installer script differs from the version expected by this launcher.'
}

$source = $source.Substring(0, $start) +
    '$inno = $env:PROJECT_CONVERT_ISCC' + "`r`n`r`n" +
    $source.Substring($end)

$id = [guid]::NewGuid().ToString('N')
$tempBuilder = Join-Path $project "build-local-$id.ps1"
$tempBatch = Join-Path $project "build-local-$id.cmd"

try {
    Set-Content -LiteralPath $tempBuilder -Value $source -Encoding UTF8
    $env:PROJECT_CONVERT_ISCC = $inno
    $env:PROJECT_CONVERT_DEV_CMD = $devCmd
    $env:PROJECT_CONVERT_BUILDER = $tempBuilder

    @'
@echo off
call "%PROJECT_CONVERT_DEV_CMD%" -no_logo -arch=x64 -host_arch=x64
if errorlevel 1 exit /b 1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PROJECT_CONVERT_BUILDER%"
exit /b %errorlevel%
'@ | Set-Content -LiteralPath $tempBatch -Encoding ASCII

    & $tempBatch
    if ($LASTEXITCODE -ne 0) {
        throw "Build stopped. See the first error above."
    }

    $setup = Join-Path $project 'dist\AnyConvertSetup.exe'
    if (!(Test-Path -LiteralPath $setup -PathType Leaf)) {
        throw "The build finished without producing $setup"
    }
    Write-Host "Installer ready: $setup"
}
finally {
    foreach ($file in @($tempBuilder, $tempBatch)) {
        if (Test-Path -LiteralPath $file) {
            Remove-Item -LiteralPath $file -Force
        }
    }
}