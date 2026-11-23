#Requires -Version 5.1
#Requires -RunAsAdministrator

$ErrorActionPreference = 'Stop'

$project = $PSScriptRoot
$identity = 'ProjectConvert.Local'
$publisher = 'CN=Project Convert Local Development'
$commandId = 'A5F9E465-48E4-49EC-9806-542702D61F2C'
$version = '1.4.0.0'

foreach ($name in @(
    'ProjectConvertShell.cpp',
    'converter_picker.py',
    'converter_gui.py',
    'converter.py',
    '.venv\Scripts\pythonw.exe'
)) {
    if (!(Test-Path -LiteralPath (Join-Path $project $name) -PathType Leaf)) {
        throw "Missing file: $name"
    }
}

if ($env:VSCMD_ARG_TGT_ARCH -ne 'x64' -or !(Get-Command cl.exe -ErrorAction SilentlyContinue)) {
    throw 'Open an x64 Native Tools Command Prompt for Visual Studio as administrator.'
}

$sdkRoot = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'
$sdk = Get-ChildItem -LiteralPath $sdkRoot -Directory |
    Where-Object {
        $_.Name -match '^10\.\d+\.\d+\.\d+$' -and
        (Test-Path -LiteralPath (Join-Path $_.FullName 'x64\MakeAppx.exe')) -and
        (Test-Path -LiteralPath (Join-Path $_.FullName 'x64\SignTool.exe'))
    } |
    Sort-Object { [version]$_.Name } -Descending |
    Select-Object -First 1

if (!$sdk) {
    throw 'Install a Windows SDK that includes MakeAppx.exe and SignTool.exe.'
}

$makeAppx = Join-Path $sdk.FullName 'x64\MakeAppx.exe'
$signTool = Join-Path $sdk.FullName 'x64\SignTool.exe'
$build = Join-Path $project 'modern-menu-build'
$packageInput = Join-Path $build 'package'
$assets = Join-Path $project 'ModernMenuAssets'

New-Item -ItemType Directory -Path $build, $packageInput, $assets -Force | Out-Null

function Invoke-Checked {
    param([string]$Executable, [string[]]$Arguments)

    & $Executable @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Executable exited with code $LASTEXITCODE"
    }
}

Push-Location -LiteralPath $build
try {
    Invoke-Checked -Executable 'cl.exe' -Arguments @(
        '/nologo', '/LD', '/O2', '/MT', '/EHsc',
        (Join-Path $project 'ProjectConvertShell.cpp'),
        '/link', 'ole32.lib', 'shlwapi.lib', 'uuid.lib',
        "/OUT:$(Join-Path $project 'ProjectConvertShell.dll')"
    )
}
finally {
    Pop-Location
}

Add-Type -AssemblyName System.Drawing

foreach ($size in @(44, 50, 150)) {
    $bitmap = New-Object System.Drawing.Bitmap -ArgumentList $size, $size
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $font = New-Object System.Drawing.Font -ArgumentList 'Segoe UI', ([single]($size * 0.38))
    $center = New-Object System.Drawing.StringFormat
    try {
        $graphics.Clear([System.Drawing.Color]::FromArgb(22, 82, 146))
        $center.Alignment = [System.Drawing.StringAlignment]::Center
        $center.LineAlignment = [System.Drawing.StringAlignment]::Center
        $bounds = New-Object System.Drawing.RectangleF -ArgumentList 0, 0, $size, $size
        $graphics.DrawString('PC', $font, [System.Drawing.Brushes]::White, $bounds, $center)
        $bitmap.Save((Join-Path $assets "$size.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $center.Dispose()
        $font.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

$manifest = @"
<?xml version="1.0" encoding="utf-8"?>
<Package xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
         xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
         xmlns:uap10="http://schemas.microsoft.com/appx/manifest/uap/windows10/10"
         xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
         xmlns:com="http://schemas.microsoft.com/appx/manifest/com/windows10"
         xmlns:desktop4="http://schemas.microsoft.com/appx/manifest/desktop/windows10/4"
         xmlns:desktop5="http://schemas.microsoft.com/appx/manifest/desktop/windows10/5"
         IgnorableNamespaces="uap uap10 rescap com desktop4 desktop5">
  <Identity Name="$identity" Publisher="$publisher" Version="$version" ProcessorArchitecture="x64" />
  <Properties>
    <DisplayName>AnyConvert</DisplayName>
    <PublisherDisplayName>AnyConvert</PublisherDisplayName>
    <Logo>ModernMenuAssets\50.png</Logo>
    <uap10:AllowExternalContent>true</uap10:AllowExternalContent>
  </Properties>
  <Dependencies>
    <TargetDeviceFamily Name="Windows.Desktop" MinVersion="10.0.22000.0" MaxVersionTested="10.0.26100.0" />
  </Dependencies>
  <Resources><Resource Language="en-US" /></Resources>
  <Applications>
    <Application Id="ProjectConvert" Executable=".venv\Scripts\pythonw.exe"
                 uap10:RuntimeBehavior="win32App" uap10:TrustLevel="mediumIL">
      <uap:VisualElements DisplayName="AnyConvert" Description="Offline file conversion"
                          Square150x150Logo="ModernMenuAssets\150.png"
                          Square44x44Logo="ModernMenuAssets\44.png"
                          BackgroundColor="transparent" AppListEntry="none" />
      <Extensions>
        <com:Extension Category="windows.comServer">
          <com:ComServer>
            <com:SurrogateServer DisplayName="AnyConvert">
              <com:Class Id="$commandId" Path="ProjectConvertShell.dll" ThreadingModel="STA" />
            </com:SurrogateServer>
          </com:ComServer>
        </com:Extension>
        <desktop4:Extension Category="windows.fileExplorerContextMenus">
          <desktop4:FileExplorerContextMenus>
            <desktop5:ItemType Type="*">
              <desktop5:Verb Id="ProjectConvert" Clsid="$commandId" />
            </desktop5:ItemType>
          </desktop4:FileExplorerContextMenus>
        </desktop4:Extension>
      </Extensions>
    </Application>
  </Applications>
  <Capabilities>
    <rescap:Capability Name="runFullTrust" />
    <rescap:Capability Name="unvirtualizedResources" />
  </Capabilities>
</Package>
"@

$manifestPath = Join-Path $packageInput 'AppxManifest.xml'
Set-Content -LiteralPath $manifestPath -Value $manifest -Encoding UTF8

$package = Join-Path $build 'ProjectConvert.msix'
Invoke-Checked -Executable $makeAppx -Arguments @(
    'pack', '/d', $packageInput, '/p', $package, '/nv', '/o'
)

$certificate = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object {
        $_.Subject -eq $publisher -and $_.HasPrivateKey -and
        $_.NotAfter -gt (Get-Date).AddDays(1)
    } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

if (!$certificate) {
    $certificate = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject $publisher `
        -CertStoreLocation 'Cert:\CurrentUser\My'
}

Invoke-Checked -Executable $signTool -Arguments @(
    'sign', '/fd', 'SHA256', '/sha1', $certificate.Thumbprint, $package
)

$trustedPath = "Cert:\LocalMachine\TrustedPeople\$($certificate.Thumbprint)"
if (!(Test-Path -LiteralPath $trustedPath)) {
    $publicCert = Join-Path $build 'ProjectConvert.cer'
    Export-Certificate -Cert $certificate -FilePath $publicCert -Force | Out-Null
    Import-Certificate -FilePath $publicCert `
        -CertStoreLocation 'Cert:\LocalMachine\TrustedPeople' | Out-Null
}

$existing = Get-AppxPackage -Name $identity
if ($existing) {
    $existing | Remove-AppxPackage -ErrorAction Stop
}

Add-AppxPackage -Path $package -ExternalLocation $project -ErrorAction Stop
Write-Host 'Installed the Project Convert command for the Windows 11 menu.'
