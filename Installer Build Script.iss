#define BuildDir SourcePath + "RetailCoder.VBE\bin\Debug"
#define AppName "RubberDuck"
#define AddinDLL "Rubberduck.dll"
#define AppVersion GetFileVersion(SourcePath + "RetailCoder.VBE\bin\Debug\Rubberduck.dll")
#define AppPublisher "RubberDuck"
#define AppURL "http://rubberduck-vba.com"
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
DefaultDirName={localappdata}\{#AppName}
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
; We are taking everything from the Build directory and adding it to the installer.  This
; might not be what we want to do.
Source: "{#BuildDir}\*"; DestDir: "{app}"; Flags: ignoreversion; Excludes: "{#AddinDLL}"; Check: Is64BitOfficeInstalled
Source: "{#BuildDir}\NativeBinaries\amd64\*"; DestDir: "{app}"; Flags: ignoreversion; Excludes: "{#AddinDLL}"; Check: Is64BitOfficeInstalled
Source: "{#BuildDir}\{#AddinDLL}"; DestDir: "{app}"; Flags: ignoreversion; Check: Is64BitOfficeInstalled; AfterInstall: RegisterAddin

Source: "{#BuildDir}\*"; DestDir: "{app}"; Flags: ignoreversion; Excludes: "{#AddinDLL}"; Check: Is32BitOfficeInstalled
Source: "{#BuildDir}\NativeBinaries\x86\*"; DestDir: "{app}"; Flags: ignoreversion; Excludes: "{#AddinDLL}"; Check: Is32BitOfficeInstalled
Source: "{#BuildDir}\{#AddinDLL}"; DestDir: "{app}"; Flags: ignoreversion; Check: Is32BitOfficeInstalled; AfterInstall: RegisterAddin

[UninstallDelete]
; Removing all application files (except for configuration).
Name: "{app}\*.dll"; Type: filesandordirs
Name: "{app}\*.xml"; Type: filesandordirs  
Name: "{app}\*.pdb"; Type: filesandordirs

[Run]
; http://stackoverflow.com/questions/5618337/how-to-register-a-net-dll-using-inno-setup
Filename: "{dotnet4032}\RegAsm.exe"; Parameters: "/codebase {#AddinDLL}"; WorkingDir: "{app}"; Flags: runascurrentuser runminimized; StatusMsg: "Registering Controls..."; Check: Is32BitOfficeInstalled
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/codebase {#AddinDLL}"; WorkingDir: "{app}"; Flags: runascurrentuser runminimized; StatusMsg: "Registering Controls..."; Check: Is64BitOfficeInstalled

[UninstallRun]
Filename: "{dotnet4032}\RegAsm.exe"; Parameters: "/u {#AddinDLL}"; WorkingDir: "{app}"; StatusMsg: "Unregistering Controls..."; Flags: runascurrentuser runminimized; Check: Is32BitOfficeInstalled
Filename: "{dotnet4064}\RegAsm.exe"; Parameters: "/u {#AddinDLL}"; WorkingDir: "{app}"; StatusMsg: "Unregistering Controls..."; Flags: runascurrentuser runminimized; Check: Is64BitOfficeInstalled

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
  // There are other values that GetBinaryType can return, but we're
  // not interested in them.
var
  HasCheckedOfficeBitness: Boolean;
  OfficeIs64Bit: Boolean;

function GetBinaryType(lpApplicationName: AnsiString; var lpBinaryType: Integer): Boolean;
external 'GetBinaryTypeA@kernel32.dll stdcall';

// TODO this only checks for Excel's bitness, but what if they don't have it installed?
function Is64BitExcelFromRegisteredExe(): Boolean;
var
  excelPath: String;
  binaryType: Integer;
begin
  Result := False; // Default value - assume 32-bit unless proven otherwise.
  // RegQueryStringValue second param is '' to get the (default) value for the key
  // with no sub-key name, as described at
  // http://stackoverflow.com/questions/913938/
  if IsWin64() and RegQueryStringValue(HKEY_LOCAL_MACHINE,
      'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe',
      '', excelPath) then begin
    // We've got the path to Excel.
    try
      if GetBinaryType(excelPath, binaryType) then begin
        Result := (binaryType = SCS_64BIT_BINARY);
      end;
    except
      // Ignore - better just to assume it's 32-bit than to let the installation
      // fail.  This could fail because the GetBinaryType function is not
      // available.  I understand it's only available in Windows 2000
      // Professional onwards.
    end;
  end;
end;

function Is64BitOfficeInstalled(): Boolean;
begin
  if (not HasCheckedOfficeBitness) then 
    OfficeIs64Bit := Is64BitExcelFromRegisteredExe();
  Result := OfficeIs64Bit;
end;

function Is32BitOfficeInstalled(): Boolean;
begin
  Result := (not Is64BitOfficeInstalled());
end;

function InitializeSetup(): Boolean;
var
   iErrorCode: Integer;
begin
  // MS .NET Framework 4.0 must be installed for this application to work.
  if Not RegKeyExists(HKLM, 'SOFTWARE\Microsoft\.NETFramework\v4.0.30319') then
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
