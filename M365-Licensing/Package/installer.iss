; Inno Setup script for the customer edition. Build with: ISCC /DAppVersion=1.0.0 installer.iss
; Run Build-CustomerPackage.ps1 first so ..\dist\M365TenantTidy exists.
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif
#ifndef Publisher
  #define Publisher "Your Company Name"
#endif
#define AppName "M365TenantTidy"

[Setup]
AppId={{6F1D2C58-9B4E-4C1A-8E73-5A2B7D9C41E6}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#Publisher}
VersionInfoVersion={#AppVersion}.0
VersionInfoCompany={#Publisher}
VersionInfoDescription={#AppName} setup
DefaultDirName={localappdata}\M365TenantTidy
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=M365TenantTidy-Setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#AppName}
#ifdef Sign
; Store rule: the installer AND every PE file it installs (including the uninstaller) must be signed
SignTool=store
SignedUninstaller=yes
#endif

[Files]
Source: "..\dist\M365TenantTidy\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Shortcuts:"

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\Start-M365LicenseGUI.cmd"; WorkingDir: "{app}"; Flags: runminimized
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\Start-M365LicenseGUI.cmd"; WorkingDir: "{app}"; Flags: runminimized; Tasks: desktopicon

[Run]
Filename: "{app}\Start-M365LicenseGUI.cmd"; Description: "Open {#AppName} now"; Flags: postinstall nowait skipifsilent runminimized
