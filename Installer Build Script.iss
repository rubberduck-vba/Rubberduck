#define BuildDir SourcePath + "RetailCoder.VBE\bin\release"
#define AppName "Rubberduck"
#define AddinDLL "Rubberduck.dll"
#define AppVersion GetFileVersion(SourcePath + "RetailCoder.VBE\bin\release\Rubberduck.dll")
#define AppPublisher "Rubberduck"
#define AppURL "http://rubberduckvba.com"
#define License SourcePath + "\License.rtf"
#define OutputDirectory SourcePath + "\Installers"
#define AddinProgId "Rubberduck.Extension"
#define AddinCLSID "8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66"

[Setup]
; TODO this CLSID should match the one used by the current installer.
AppId={{979AFF96-DD9E-4FC2-802D-9E0C36A60D09}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
; use the local appdata folder instead of the program files dir.
DefaultDirName={commonappdata}\{#AppName}
DefaultGroupName=Rubberduck
AllowNoIcons=yes
LicenseFile={#License}
OutputDir={#OutputDirectory}
OutputBaseFilename=Rubberduck.Setup.{#AppVersion}
Compression=lzma
SolidCompression=yes

ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
; TODO add additional installation languages here.
Name: "English"; MessagesFile: "compiler:Default.isl"

[Files]
; Install the correct bitness binaries.
Source: "{#BuildDir}\NativeBinaries\amd64\*"; DestDir: "{app}"; Flags: ignoreversion; Excludes: "{#AddinDLL}"; Check: Is64BitOfficeInstalled
Source: "{#BuildDir}\NativeBinaries\x86\*"; DestDir: "{app}"; Flags: ignoreversion; Excludes: "{#AddinDLL}"; Check: Is32BitOfficeInstalled

Source: "{#BuildDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs; Excludes: "{#AddinDLL},\NativeBinaries"
Source: "{#BuildDir}\{#AddinDLL}"; DestDir: "{app}"; Flags: ignoreversion; AfterInstall: RegisterAddin

[Run]
; http://stackoverflow.com/questions/5618337/how-to-register-a-net-dll-using-inno-setup
Filename: "{dotnet4032}\RegAsm.exe"; Parameters: "/codebase {#AddinDLL}"; WorkingDir: "{app}"; Flags: runascurrentuser runminimized; StatusMsg: "Registering Controls..."; Check: Is32BitOfficeInstalled
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/codebase {#AddinDLL}"; WorkingDir: "{app}"; Flags: runascurrentuser runminimized; StatusMsg: "Registering Controls..."; Check: Is64BitOfficeInstalled

[UninstallRun]
Filename: "{dotnet4032}\RegAsm.exe"; Parameters: "/u {#AddinDLL}"; WorkingDir: "{app}"; StatusMsg: "Unregistering Controls..."; Flags: runascurrentuser runminimized; Check: Is32BitOfficeInstalled
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/u {#AddinDLL}"; WorkingDir: "{app}"; StatusMsg: "Unregistering Controls..."; Flags: runascurrentuser runminimized; Check: Is64BitOfficeInstalled

[UninstallDelete]
Type: filesandordirs; Name: "{localappdata}\{#AppName}"

[CustomMessages]
; TODO add additional languages here.
English.NETFramework40NotInstalled=Microsoft .NET Framework 4.0 installation was not detected.

[Icons]
Name: "{group}\{cm:ProgramOnTheWeb,{#AppName}}"; Filename: "{#AppURL}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"

[Code]
// The following code is adapted from: http://stackoverflow.com/a/11651515/2301065
const
  SCS_32BIT_BINARY = 0;
  SCS_64BIT_BINARY = 6;
  // There are other values that GetBinaryType can return, but we're not interested in them.
  OfficeNotFound = -1;
  
var
  HasCheckedOfficeBitness: Boolean;
  OfficeIs64Bit: Boolean;

function GetBinaryType(lpApplicationName: AnsiString; var lpBinaryType: Integer): Boolean;
external 'GetBinaryTypeA@kernel32.dll stdcall';

function GetOfficeAppBitness(exeName: string): Integer;
var
  appPath: String;
  binaryType: Integer;
begin
  Result := OfficeNotFound;  // Default value.

  if RegQueryStringValue(HKEY_LOCAL_MACHINE,
    'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\' + exeName,
    '', appPath) then begin
    try
      if GetBinaryType(appPath, binaryType) then Result := binaryType;
    except
    end;
  end;
end;

function GetOfficeBitness(): Integer;
var
  appBitness: Integer;
  officeExeNames: array[0..6] of String;
  i: Integer;
begin
  officeExeNames[0] := 'excel.exe';
  officeExeNames[1] := 'msaccess.exe';
  officeExeNames[2] := 'winword.exe';
  officeExeNames[3] := 'outlook.exe';
  officeExeNames[4] := 'powerpnt.exe';
  officeExeNames[5] := 'mspub.exe';
  officeExeNames[6] := 'winproj.exe';

  for i := 0 to 4 do begin
    appBitness := GetOfficeAppBitness(officeExeNames[i]);
    if appBitness <> OfficeNotFound then begin
      Result := appBitness;
      exit;
    end;
  end;
  // Note if we get to here then we haven't found any Office versions.  Should
  // we fail the installation?
end;

function Is64BitOfficeInstalled(): Boolean;
begin
  if (not HasCheckedOfficeBitness) then 
    OfficeIs64Bit := (GetOfficeBitness() = SCS_64BIT_BINARY);
  Result := OfficeIs64Bit;
end;

function Is32BitOfficeInstalled(): Boolean;
begin
  Result := (not Is64BitOfficeInstalled());
end;

// http://kynosarges.org/DotNetVersion.html
function IsDotNetDetected(version: string; service: cardinal): boolean;
// Indicates whether the specified version and service pack of the .NET Framework is installed.
//
// version -- Specify one of these strings for the required .NET Framework version:
//    'v1.1.4322'     .NET Framework 1.1
//    'v2.0.50727'    .NET Framework 2.0
//    'v3.0'          .NET Framework 3.0
//    'v3.5'          .NET Framework 3.5
//    'v4\Client'     .NET Framework 4.0 Client Profile
//    'v4\Full'       .NET Framework 4.0 Full Installation
//    'v4.5'          .NET Framework 4.5
//
// service -- Specify any non-negative integer for the required service pack level:
//    0               No service packs required
//    1, 2, etc.      Service pack 1, 2, etc. required
var
    key: string;
    install, release, serviceCount: cardinal;
    check45, success: boolean;
begin
    // .NET 4.5 installs as update to .NET 4.0 Full
    if version = 'v4.5' then begin
        version := 'v4\Full';
        check45 := true;
    end else
        check45 := false;

    // installation key group for all .NET versions
    key := 'SOFTWARE\Microsoft\NET Framework Setup\NDP\' + version;

    // .NET 3.0 uses value InstallSuccess in subkey Setup
    if Pos('v3.0', version) = 1 then begin
        success := RegQueryDWordValue(HKLM, key + '\Setup', 'InstallSuccess', install);
    end else begin
        success := RegQueryDWordValue(HKLM, key, 'Install', install);
    end;

    // .NET 4.0/4.5 uses value Servicing instead of SP
    if Pos('v4', version) = 1 then begin
        success := success and RegQueryDWordValue(HKLM, key, 'Servicing', serviceCount);
    end else begin
        success := success and RegQueryDWordValue(HKLM, key, 'SP', serviceCount);
    end;

    // .NET 4.5 uses additional value Release
    if check45 then begin
        success := success and RegQueryDWordValue(HKLM, key, 'Release', release);
        success := success and (release >= 378389);
    end;

    result := success and (install = 1) and (serviceCount >= service);
end;

function InitializeSetup(): Boolean;
var
   iErrorCode: Integer;
begin
  // MS .NET Framework 4.5 must be installed for this application to work.
  if not IsDotNetDetected('v4.5', 0) then
  begin
    MsgBox(ExpandConstant('{cm:NETFramework40NotInstalled}'), mbCriticalError, mb_Ok);
    ShellExec('open', 'http://msdn.microsoft.com/en-us/netframework/aa731542', '', '', SW_SHOW, ewNoWait, iErrorCode) 
    Result := False;
  end
  else
    Result := True;
end;

procedure RegisterAddinForIDE(const iRootKey: Integer; const sAddinSubKey: String; const sProgIDConnect: String);
begin
   RegWriteStringValue(iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'FriendlyName', '{#AppName}');
   RegWriteStringValue(iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'Description' , '{#AppName}');
   RegWriteDWordValue (iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'LoadBehavior', 3);
end;

procedure UnregisterAddinForIDE(const iRootKey: Integer; const sAddinSubKey: String; const sProgIDConnect: String);
begin
   if RegKeyExists(iRootKey, sAddinSubKey + '\' + sProgIDConnect) then
      RegDeleteKeyIncludingSubkeys(iRootKey, sAddinSubKey + '\' + sProgIDConnect);
end;

procedure RegisterAddin();
begin
  if Is32BitOfficeInstalled() then
    RegisterAddinForIDE(HKCU32, 'Software\Microsoft\VBA\VBE\6.0\Addins', '{#AddinProgId}');

  if Is64BitOfficeInstalled() then 
    RegisterAddinForIDE(HKCU64, 'Software\Microsoft\VBA\VBE\6.0\Addins64', '{#AddinProgId}');
end;

procedure UnregisterAddin();
begin
  UnregisterAddinForIDE(HKCU32, 'Software\Microsoft\VBA\VBE\6.0\Addins', '{#AddinProgId}');
  if IsWin64() then 
    UnregisterAddinForIDE(HKCU64, 'Software\Microsoft\VBA\VBE\6.0\Addins64', '{#AddinProgId}');
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then UnregisterAddin();
end;
