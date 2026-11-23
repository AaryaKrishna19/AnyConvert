#Requires -Version 5.1

$ErrorActionPreference = 'Stop'

$project = $PSScriptRoot
$python = Join-Path $project '.venv\Scripts\python.exe'
$inno = Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'
if (!(Test-Path -LiteralPath $inno)) {
    $inno = Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe'
}

foreach ($file in @(
    $python,
    $inno,
    (Join-Path $project 'converter.py'),
    (Join-Path $project 'project_convert_app.py'),
    (Join-Path $project 'ProjectConvertShell.cpp'),
    (Join-Path $project 'install-modern-menu.ps1')
)) {
    if (!(Test-Path -LiteralPath $file -PathType Leaf)) {
        throw "Missing required file: $file"
    }
}

if ($env:VSCMD_ARG_TGT_ARCH -ne 'x64' -or
    !(Get-Command cl.exe -ErrorAction SilentlyContinue)) {
    throw 'Run this from the x64 Native Tools Command Prompt.'
}

$build = Join-Path $project 'setup-build'
$stage = Join-Path $build 'payload'
$spec = Join-Path $build 'spec'
$output = Join-Path $project 'dist'
New-Item -ItemType Directory -Force -Path $build, $stage, $spec, $output | Out-Null

function Replace-Required {
    param(
        [string]$Text,
        [string]$Old,
        [string]$New,
        [string]$File
    )

    if (!$Text.Contains($Old)) {
        throw "The expected code was not found in $File. Check that it matches the earlier complete file."
    }

    $Text.Replace($Old, $New)
}

function Invoke-Checked {
    param([string]$Executable, [string[]]$Arguments)

    & $Executable @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Executable exited with code $LASTEXITCODE"
    }
}

# Resolve release payloads on the build PC. Setup uses only embedded files.
Invoke-Checked $python @('-m','pip','install','-r',(Join-Path $project 'requirements.txt'))
Invoke-Checked $python @('-m','pip','check')
Invoke-Checked $python @('-c','import pymupdf, pdf2docx, PIL, openpyxl, docx, reportlab, pptx, cv2, numpy, tkinter, PIL.ImageTk')
$bundle = Join-Path $build ('offline-' + [guid]::NewGuid().ToString('N'))
$officeDownload = Join-Path $bundle 'office-download'
$vcDownload = Join-Path $bundle 'vc-download'
New-Item -ItemType Directory -Path $officeDownload,$vcDownload -Force | Out-Null
$winget = (Get-Command winget.exe -ErrorAction Stop).Source
foreach ($package in @(
    @{ Id='TheDocumentFoundation.LibreOffice'; Directory=$officeDownload },
    @{ Id='Microsoft.VCRedist.2015+.x64'; Directory=$vcDownload }
)) {
    Invoke-Checked $winget @('download','--id',$package.Id,'--exact','--source','winget',
        '--architecture','x64','--download-directory',$package.Directory,
        '--accept-package-agreements','--accept-source-agreements','--skip-dependencies')
}
$msi = @(Get-ChildItem $officeDownload -Recurse -Filter '*.msi')
$vc = @(Get-ChildItem $vcDownload -Recurse -Filter '*.exe')
if ($msi.Count -ne 1 -or $vc.Count -ne 1) { throw 'Missing or ambiguous offline installer payload.' }
$vcInstaller = $vc[0].FullName
$officeDir = Join-Path $bundle 'office'
$extract = Start-Process msiexec.exe -ArgumentList ('/a "{0}" /qn /norestart TARGETDIR="{1}"' -f $msi[0].FullName,$officeDir) -Wait -PassThru -WindowStyle Hidden
if ($extract.ExitCode -ne 0) { throw "Office engine staging failed: $($extract.ExitCode)" }
$engine = @(Get-ChildItem $officeDir -Recurse -Filter soffice.exe)
if ($engine.Count -ne 1) { throw 'Could not locate the staged Office engine.' }
$officeRoot = $engine[0].Directory.Parent.FullName

# The frozen app reads '<format> <path>' from Explorer itself, so it is used as written.
$appName = 'project_convert_app.py'
Copy-Item -LiteralPath (Join-Path $project 'anyconvert_formats.py') -Destination $build -Force
Copy-Item -LiteralPath (Join-Path $project $appName) -Destination $build -Force
$appPath = Join-Path $build $appName

# Generate one COM command per output format from the existing DLL source.
$routeJson = & $python (Join-Path $project 'export_routes.py')
if ($LASTEXITCODE -ne 0) { throw 'Could not read conversion routes.' }
$routes = $routeJson | ConvertFrom-Json
$titles = @{ pdf='A4 PDF'; docx='Word'; 'docx-fast'='Word (fast text)'; pptx='PowerPoint'; xlsx='Excel'; jpg='JPEG' }
$cppTemplate = Get-Content -Raw -LiteralPath (Join-Path $project 'ProjectConvertShell.cpp')
$classes = @()
$items = @()
$extensions = @($routes.PSObject.Properties.Name)
foreach ($index in 0..($extensions.Count - 1)) {
    $extension = $extensions[$index]
    $prefix = '{0:X8}' -f [uint32](2784625792 + $index)
    $clsid = "$prefix-48E4-49EC-9806-542702D61F2C"
    $dllName = "AnyConvertMenu$index.dll"
    $choices = foreach ($format in $routes.$extension) {
        $title = $titles[$format]
        if (!$title) { $title = $format.ToUpperInvariant() }
        '{{ L"{0}", L"Convert to {1}" }}' -f $format, $title
    }
    $definition = 'static const struct { PCWSTR format; PCWSTR title; } choices[] = {' + ($choices -join ', ') + '};'
    $cpp = $cppTemplate.Replace('static HMODULE moduleHandle', (('static PCWSTR registeredExtension = L"{0}";' -f $extension) + [Environment]::NewLine + $definition + [Environment]::NewLine + 'static HMODULE moduleHandle'))
    $cpp = $cpp.Replace('0xa5f9e465,', "0x$prefix,")
    $cppPath = Join-Path $build "AnyMenu$index.cpp"
    Set-Content -LiteralPath $cppPath -Value $cpp -Encoding UTF8
    Push-Location -LiteralPath $build
    try {
        Invoke-Checked 'cl.exe' @('/nologo', '/LD', '/O2', '/MT', '/EHsc', $cppPath,
            '/link', 'ole32.lib', 'shlwapi.lib', 'uuid.lib', "/OUT:$(Join-Path $stage $dllName)")
    } finally { Pop-Location }
    $classes += '<com:Class Id="{0}" Path="{1}" ThreadingModel="STA" />' -f $clsid, $dllName
    $items += '<desktop5:ItemType Type="{0}"><desktop5:Verb Id="AnyConvert{1}" Clsid="{2}" /></desktop5:ItemType>' -f $extension, $index, $clsid
}

# Reuse the existing manifest, SDK discovery, logo generation, and signing.
# The generated script only builds the signed package and exports its certificate.
$oldInstallName = 'install-modern-menu.ps1'
$pack = (Get-Content -Raw -LiteralPath (Join-Path $project $oldInstallName)) -replace "`r`n", "`n"
$pack = Replace-Required $pack '#Requires -RunAsAdministrator' '' $oldInstallName
$start = $pack.IndexOf('foreach ($name in @(')
$end = $pack.IndexOf('if ($env:VSCMD_ARG_TGT_ARCH', $start)
if ($start -lt 0 -or $end -le $start) { throw 'Could not locate the old input checks.' }
$pack = $pack.Remove($start, $end - $start)

$start = $pack.IndexOf('Push-Location -LiteralPath $build')
$end = $pack.IndexOf('Add-Type -AssemblyName System.Drawing', $start)
if ($start -lt 0 -or $end -le $start) { throw 'Could not locate the old DLL build block.' }
$pack = $pack.Remove($start, $end - $start)

$pack = Replace-Required $pack "'ProjectConvert.Local'" "'ProjectConvert.Setup'" $oldInstallName
$pack = Replace-Required $pack 'Executable=".venv\Scripts\pythonw.exe"' `
    'Executable="ProjectConvert.exe"' $oldInstallName
$pack = Replace-Required $pack `
    '<com:Class Id="$commandId" Path="ProjectConvertShell.dll" ThreadingModel="STA" />' `
    ($classes -join "`n") $oldInstallName

$oldItem = [regex]::Match(
    $pack, '(?s)<desktop5:ItemType Type="\*">.*?</desktop5:ItemType>'
).Value
if (!$oldItem) { throw 'Could not locate the old Explorer menu declaration.' }
$pack = Replace-Required $pack $oldItem ($items -join "`n") $oldInstallName

$start = $pack.IndexOf('$trustedPath =')
if ($start -lt 0) { throw 'Could not locate the old registration block.' }
$pack = $pack.Substring(0, $start) + @'

Export-Certificate -Cert $certificate -FilePath (Join-Path $project 'ProjectConvert.cer') -Force | Out-Null
Copy-Item -LiteralPath $package -Destination (Join-Path $project 'ProjectConvert.msix') -Force
'@

$packPath = Join-Path $stage 'build-package.ps1'
Set-Content -LiteralPath $packPath -Value $pack -Encoding UTF8

Invoke-Checked $python @(
    '-m', 'PyInstaller', '--noconfirm', '--clean', '--onedir', '--windowed',
    '--icon', (Join-Path $project 'AnyConvert.ico'), '--name', 'ProjectConvert', '--paths', $project,
    '--specpath', $spec, '--workpath', (Join-Path $build 'pyinstaller'),
    '--distpath', $stage, '--collect-all', 'pymupdf', '--collect-all', 'pdf2docx', '--collect-all', 'pptx',
    '--hidden-import', 'PIL.ImageTk', '--hidden-import', 'PIL._imagingtk',
    $appPath
)
Invoke-Checked 'powershell.exe' @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $packPath
)

foreach ($size in @(44, 50, 150)) {
    Copy-Item -LiteralPath (Join-Path $project "AnyConvert$size.png") -Destination (Join-Path $stage "ModernMenuAssets\$size.png") -Force
}
$registrationScript = @'
param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
$identity = 'ProjectConvert.Setup'

if ($Uninstall) {
    Get-AppxPackage -Name $identity | Remove-AppxPackage
    exit
}

$certificate = Join-Path $PSScriptRoot 'ProjectConvert.cer'
Import-Certificate -FilePath $certificate -CertStoreLocation 'Cert:\LocalMachine\TrustedPeople' | Out-Null

Get-AppxPackage -Name $identity | Remove-AppxPackage
Add-AppxPackage -Path (Join-Path $PSScriptRoot 'ProjectConvert.msix') -ExternalLocation $PSScriptRoot

# ConvertAny legacy menu migration
# Run after the replacement menu has registered successfully.
$oldFormats = @{
    '.pdf'  = @('png', 'jpg', 'docx', 'docx-fast', 'xlsx')
    '.png'  = @('pdf', 'jpg')
    '.jpg'  = @('pdf', 'png')
    '.jpeg' = @('pdf', 'png')
    '.xlsx' = @('pdf', 'docx')
}

foreach ($extension in $oldFormats.Keys) {
    foreach ($format in $oldFormats[$extension]) {
        $key = "HKCU:\Software\Classes\SystemFileAssociations\$extension\shell\ProjectConvert-$format"
        $commandKey = Get-Item -LiteralPath "$key\command" -ErrorAction SilentlyContinue
        if (!$commandKey) { continue }

        $oldCommand = [string]$commandKey.GetValue('')
        if ($oldCommand -match '(?i)\\(?:converter_gui|converter)\.py"') {
            Remove-Item -LiteralPath $key -Recurse -Force -ErrorAction Stop
        }
    }
}

# The early development installer registered a separate modern-menu package.
$developmentPackage = Get-AppxPackage -Name 'ProjectConvert.Local' |
    Where-Object { $_.Publisher -eq 'CN=Project Convert Local Development' }
foreach ($package in $developmentPackage) {
    Remove-AppxPackage -Package $package.PackageFullName -ErrorAction Stop
}
'@

$registerPath = Join-Path $stage 'register-menu.ps1'
Set-Content -LiteralPath $registerPath -Value $registrationScript -Encoding UTF8

$iss = @"
[Setup]
AppId=ProjectConvert.Setup
AppName=AnyConvert
AppVersion=1.4.0
SetupIconFile=$project\AnyConvert.ico
UninstallDisplayIcon={app}\AnyConvert.ico
DefaultDirName={autopf}\Project Convert
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.22000
OutputDir=$output
OutputBaseFilename=AnyConvertSetup
Compression=lzma2
SolidCompression=yes

[InstallDelete]
Type: files; Name: "{app}\ProjectConvertShell0.dll"
Type: files; Name: "{app}\ProjectConvertShell1.dll"
Type: files; Name: "{app}\ProjectConvertShell2.dll"
Type: files; Name: "{app}\ProjectConvertShell3.dll"
Type: files; Name: "{app}\ProjectConvertShell4.dll"
Type: files; Name: "{app}\ProjectConvertShell5.dll"

[Files]
Source: "$vcInstaller"; DestName: "vc_redist.x64.exe"; Flags: dontcopy
Source: "$officeRoot\*"; DestDir: "{app}\OfficeEngine"; Excludes: "*.msi"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "$project\AnyConvert.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "$stage\ProjectConvert\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "$stage\AnyConvertMenu*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "$stage\ProjectConvert.msix"; DestDir: "{app}"; Flags: ignoreversion
Source: "$stage\ProjectConvert.cer"; DestDir: "{app}"; Flags: ignoreversion
Source: "$stage\register-menu.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "$stage\ModernMenuAssets\*.png"; DestDir: "{app}\ModernMenuAssets"; Flags: ignoreversion

[Code]
var
  IsUpdate: Boolean;

function InitializeSetup(): Boolean;
var
  OldVersion: String;
begin
  IsUpdate := RegQueryStringValue(HKLM64,
      'Software\Microsoft\Windows\CurrentVersion\Uninstall\ProjectConvert.Setup_is1',
      'DisplayVersion', OldVersion);
  Result := True;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Code: Integer;
begin
  Result := '';
  ExtractTemporaryFile('vc_redist.x64.exe');
  if not Exec(ExpandConstant('{tmp}\vc_redist.x64.exe'), '/install /quiet /norestart',
              '', SW_HIDE, ewWaitUntilTerminated, Code) then begin
    Result := 'Could not launch the bundled Microsoft runtime.';
    Exit;
  end;
  if Code = 3010 then NeedsRestart := True
  else if (Code <> 0) and (Code <> 1638) then
    Result := 'Microsoft runtime installation failed: ' + IntToStr(Code);
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := IsUpdate and (PageID < wpInstalling);
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if IsUpdate and (CurPageID = wpInstalling) then begin
    WizardForm.PageNameLabel.Caption := 'Updating AnyConvert';
    WizardForm.PageDescriptionLabel.Caption := 'Applying the new version.';
    WizardForm.StatusLabel.Visible := False;
    WizardForm.FilenameLabel.Visible := False;
    WizardForm.FilenameLabel.Caption := '';
  end;
end;

procedure RegisterMenu(Uninstall: Boolean);
var
  Args: String;
  ResultCode: Integer;
begin
  Args := '-NoProfile -ExecutionPolicy Bypass -File "' +
          ExpandConstant('{app}\register-menu.ps1') + '"';
  if Uninstall then
    Args := Args + ' -Uninstall';

  if not Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
              Args, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    RaiseException('Could not start the Explorer registration script.');

  if ResultCode <> 0 then
    RaiseException('Explorer registration failed. PowerShell exit code: ' +
                   IntToStr(ResultCode));
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    RegisterMenu(False);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
    RegisterMenu(True);
end;
"@

$issPath = Join-Path $build 'ProjectConvert.iss'
Set-Content -LiteralPath $issPath -Value $iss -Encoding UTF8
Invoke-Checked $inno @($issPath)

Write-Host "Installer ready: $(Join-Path $output 'AnyConvertSetup.exe')"
