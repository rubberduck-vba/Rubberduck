;The file must be encoded in UTF-8 BOM

#pragma include __INCLUDE__ + ";" + SourcePath + "Includes\"

#define protected
#define BuildDir ExtractFileDir(ExtractFileDir(SourcePath)) + "\bin\"
#define IncludesDir SourcePath + "Includes\"
#define GraphicsDir SourcePath + "Graphics\"
#define AppName "Rubberduck"
#define AddinDLL "Rubberduck.dll"
#define Tlb32bit "Rubberduck.x32.tlb"
#define Tlb64bit "Rubberduck.x64.tlb"
#define DllFullPath BuildDir + AddinDLL
#define Tlb32bitFullPath BuildDir + Tlb32bit
#define Tlb64bitFullPath BuildDir + Tlb64bit
#define AppVersion GetFileVersion(BuildDir + "Rubberduck.dll")
#define AppPublisher "Rubberduck"
#define AppURL "http://rubberduckvba.com"
#define License IncludesDir + "License.rtf"
#define OutputDirectory SourcePath + "Installers\"
#define AddinProgId "Rubberduck.Extension"
#define AddinCLSID "8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66"

; Output the defined constants to aid in verification
#pragma message "Include: " + __INCLUDE__
#pragma message "SourcePath: " + SourcePath
#pragma message "BuildDir: " + BuildDir
#pragma message "GraphicsDir: " + GraphicsDir
#pragma message "AppName: " + AppName
#pragma message "AddinDLL: " + AddinDLL
#pragma message "DllFullPath: " + DllFullPath
#pragma message "Tlb32bitFullPath: " + Tlb32bitFullPath
#pragma message "Tlb64bitFullPath: " + Tlb64bitFullPath
#pragma message "AppVersion: " + AppVersion
#pragma message "AppPublisher: " + AppPublisher
#pragma message "AppURL: " + AppURL
#pragma message "License: " + License
#pragma message "OutputDirectory: " + OutputDirectory
#pragma message "AddinProgId: " + AddinProgId
#pragma message "AddinCLSID: " + AddInCLSID

[Setup]
; The Previous language must be no when AppId uses constants
UsePreviousLanguage=no
AppId={code:GetAppId}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={code:GetDefaultDirName}
DefaultGroupName=Rubberduck
AllowNoIcons=yes
LicenseFile={#License}
OutputDir={#OutputDirectory}
OutputBaseFilename=Rubberduck.Setup
Compression=lzma
SolidCompression=yes

ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64

SetupLogging=yes
PrivilegesRequired=lowest

UninstallFilesDir={app}\Installers
UninstallDisplayName={code:GetUninstallDisplayName}
UninstallDisplayIcon={group}\Ducky.ico
Uninstallable=ShouldCreateUninstaller()
CreateUninstallRegKey=ShouldCreateUninstaller()

; should be at least a 55 x 55 bitmap
WizardSmallImageFile={#GraphicsDir}Rubberduck.Duck.Small.55x55.bmp, \
                     {#GraphicsDir}Rubberduck.Duck.Small.64x68.bmp, \
                     {#GraphicsDir}Rubberduck.Duck.Small.83x80.bmp, \
                     {#GraphicsDir}Rubberduck.Duck.Small.92x97.bmp, \
                     {#GraphicsDir}Rubberduck.Duck.Small.110x106.bmp, \
                     {#GraphicsDir}Rubberduck.Duck.Small.119x123.bmp, \
                     {#GraphicsDir}Rubberduck.Duck.Small.138x140.bmp

; should be at least a 164 x 314 bitmap
WizardImageFile={#GraphicsDir}Rubberduck.Duck.164x314.bmp, \
                {#GraphicsDir}Rubberduck.Duck.192x386.bmp, \
                {#GraphicsDir}Rubberduck.Duck.246x459.bmp, \
                {#GraphicsDir}Rubberduck.Duck.273x556.bmp, \
                {#GraphicsDir}Rubberduck.Duck.328x604.bmp, \
                {#GraphicsDir}Rubberduck.Duck.355x700.bmp

[Languages]
Name: "English"; MessagesFile: "compiler:Default.isl"
Name: "French"; MessagesFile: "compiler:Languages\French.isl"
Name: "German"; MessagesFile: "compiler:Languages\German.isl"
Name: "Czech"; MessagesFile: "compiler:Languages\Czech.isl"
Name: "Spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Dirs]
; Make folder "readonly" to support icons (it does not actually make folder readonly. A weird Windows quirk)
Name: {group}; Attribs: readonly

[Files]
; Install the correct bitness binaries.
Source: "{#BuildDir}*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs replacesameversion; Excludes: "Rubberduck.Deployment.*,olewoo.*,*.idl,*.dll.xml,*.tlb.xml,{#AddinDLL},\NativeBinaries"; Check: CheckShouldInstallFiles
Source: "{#BuildDir}{#AddinDLL}"; DestDir: "{app}"; Flags: ignoreversion replacesameversion; Check: CheckShouldInstallFiles;

; Used for customizing the Start menu folder appearance
Source: "desktop.ini"; DestDir: "{group}"; Attribs: hidden system; Flags: ignoreversion replacesameversion; Check: CheckShouldInstallFiles;
Source: "{#GraphicsDir}Ducky.ico"; DestDir: "{group}"; Attribs: hidden system; Flags: ignoreversion replacesameversion; Check: CheckShouldInstallFiles;

; Makes it easier to fix VBE registration issues
Source: "{#IncludesDir}Rubberduck.RegisterAddIn.bat"; DestDir: "{app}"; Flags: ignoreversion replacesameversion;
Source: "{#IncludesDir}Rubberduck.RegisterAddIn.reg"; DestDir: "{app}"; Flags: ignoreversion replacesameversion;

[Registry]
; DO NOT attempt to register VBE Add-In with this section. It doesn't work
; Use [Code] section (RegisterAddIn procedure) to register the entries instead.
#include <RegistryCleanup.reg.iss>
#include <Rubberduck.reg.iss>

; Commneted out because we don't want to delete users setting when they are just
; uninstalling to install another version of Rubberduck. Considered prompting to
; delete but not for right now.
; [UninstallDelete]
; Type: filesandordirs; Name: "{userappdata}\{#AppName}"

[CustomMessages]
; TODO add additional languages here by adding include files in \Includes folder
;      and uncomment or add lines to include the file.
#include <English.CustomMessages.iss>
#include <French.CustomMessages.iss>
#include <German.CustomMessages.iss>
#include <Czech.CustomMessages.iss>
#include <Spanish.CustomMessages.iss>

[Icons]
Name: "{group}\{cm:ProgramOnTheWeb,{#AppName}}"; Filename: "{#AppURL}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"
Name: "{group}\{cm:RegisterAddin, {#AppName}}"; Filename: "{app}\Rubberduck.RegisterAddIn.bat"; WorkingDir: "{app}";

[Code]
///<remarks>
/// The code section is divided into subsections:
///   global declarations of const and variables
///   external functions
///   helper functions
///   event functions
///</remarks>

// Global declarations Section

type
  HINSTANCE = THandle;

const
   ///<remarks>
   ///Identifiers for installation in everyone mode to support
   ///functions that needs to distinguish between install modes.
   ///</remarks>
   EveryoneAppMode = 'AllUsers';
   EveryoneAppId = '{979AFF96-DD9E-4FC2-802D-9E0C36A60D09}';
   PerUserAppMode = 'CurrentUser';
   PerUserAppId = '{DF0E0E6F-2CED-482E-831C-7E9721EB66AA}';

  ///<remarks>
  ///Pseudo-Bitwise enum to support functions for
  ///checking previous versions of different modes.
  ///Pascal scripting doesn't have bitwise operators
  ///so we improvise....
  ///<remarks>
  NoneIsInstalled = 0;
  PerUserIsInstalled = 1;
  EveryoneIsInstalled = 2;
  BothAreInstalled = 3;

var
  ///<remarks>
  ///Flag to indicate that we automatically skipped pages during an
  ///elevated install.
  ///
  ///The flag should be only changed within the <see cref="ShouldSkipPage" />
  ///event function. Treat it as read-only in all other contexts.
  ///</remarks>
  PagesSkipped: Boolean;

  ///<remarks>
  ///Custom page to select whether the installer should run for the
  ///current user only or for all users (requiring elevation).
  ///Configured in the WizardInitialize event function.
  ///</remarks>
  InstallForWhoOptionPage: TInputOptionWizardPage;

  ///<remarks>
  ///Custom page to indicate whether the installer should create per-user
  ///registry key to make VBE addin available to the application.
  ///<see cref="RegisterAddIn" />
  ///</remakrs>
  RegisterAddInOptionPage: TInputOptionWizardPage;

  ///<remarks>
  ///Set by the <see cref="RegisterAddInOptionPage" />; should be
  ///read-only normally. When set to true, it is used to check if
  ///elevation is needed for the installer.
  ///</remarks>
  ShouldInstallAllUsers: Boolean;

  ///<remarks>
  ///When the non-elevated installer launches the elevated installer
  ///via the <see cref="Elevate" />, it will pass in a switch. This
  ///helps us to know that the elevated installer was started in this
  ///manner rather than user right-clicking and choosing "Run As Administrator".
  ///This mainly influences whether we should skip the pages. Treat it
  ///as read-only in all contexts other than <see cref="InitializeWizard" />.
  ///<remarks>
  HasElevateSwitch : Boolean;

  ///<remarks>
  ///Indicates that the installer can only run in per-user only. This is typically
  ///when there is a previous version of Rubberduck and the user either won't or
  ///can't uninstall it. Therefore, we cannot safely install using all-user mode
  ///in this context. This is set in the <see cref="SetupInitialize" /> event
  ///function and should be read-only in all other contexts.
  ///</remarks>
  PerUserOnly : Boolean;

// External functions section

///<remarks>
///Used to select correct subtype of Win32 API
///based on the installer's encoding.
///<remarks>
#ifdef UNICODE
  #define AW "W"
#else
  #define AW "A"
#endif

///<remarks>
///Win32 API function used to launch a 2nd instance of the installer
///under elevated context, used by <see cref="Elevate" />.
///</remarks>
function ShellExecute(hwnd: HWND; lpOperation: string; lpFile: string;
  lpParameters: string; lpDirectory: string; nShowCmd: Integer): HINSTANCE;
  external 'ShellExecute{#AW}@shell32.dll stdcall';

// Helper functions section

///<remarks>
///Encapuslate the check whether the installer is running in an
///elevated context or not. This does not indicate whether the
///installer was launched by a non-elevated installer, however.
///<see cref="HasElevateSwitch" />
///</remarks>
function IsElevated: Boolean;
begin
  Result := IsAdminLoggedOn;
end;

///<remarks>
///Generic helper function to parse the command line arguments,
///and indicate whether a particular argument was passed in.
///</remarks>
function CmdLineParamExists(const Value: string): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 1 to ParamCount do
    if CompareText(ParamStr(I), Value) = 0 then
    begin
      Result := True;
      Exit;
    end;
end;

///<remarks>
///Used to determine whether a install directory that user
///selected is in fact writable by the user, especially
///for non-elevated installation.
///</remarks>
function HaveWriteAccessToApp: Boolean;
var
  PathName: string;
  FileName: string;
begin

  //DirExists() will return true if the path has a ending backslash
  //so we must take care to remove the backslash while we locate a
  //existent directory to avoid trying to create a file to a directory
  //that doesn't exists. We can assume that if we can create a file in
  //the first existent directory, we should be able to create directories
  PathName := RemoveBackslash(WizardDirValue);
  Log('Starting PathName: ' + PathName);

  while not DirExists(PathName) do
  begin
     PathName := RemoveBackSlash(ExtractFilePath(PathName));
     Log('Modified PathName: ' + PathName);
  end;

  FileName := AddBackSlash(PathName) + 'writetest.tmp';
  Result := SaveStringToFile(FileName, 'test', False);
  if Result then
  begin
    Log(Format(
      'Have write aElevationRequiredForSelectedFolderWarningccess to the last installation path [%s]', [WizardDirValue]));
    DeleteFile(FileName);
  end
    else
  begin
    Log(Format('Does not have write access to the last installation path [%s]', [
      WizardDirValue]));
  end;
end;

///<remarks>
///When user requests install for all users or into a protected directory,
///and the current installer isn't elevated, it will create a second instance
///of itself, passing in all the switches, and adding the `/ELEVATE` switch
///to complete the installation of the add-in under the elevated context.
///The original installer does not terminate but rather remains open for the
///elevated install to complete so that the user can then proceed with VBE
///addin registration on the <see cref="RegisterAddInOptionPage" /> page.
///<remarks>
function Elevate: Boolean;
var
  I: Integer;
  instance: HINSTANCE;
  Params: string;
  S: string;
begin
  { Collect current instance parameters }
  for I := 1 to ParamCount do
  begin
    S := ParamStr(I);
    Log('Parameter: ' + S);
    { Unique log file name for the elevated instance }
    if CompareText(Copy(S, 1, 5), '/LOG=') = 0 then
    begin
      S := S + 'elevated';
    end;
    { Do not pass our /SL5 switch }
    if CompareText(Copy(S, 1, 5), '/SL5=') <> 0 then
    begin
      Params := Params + AddQuotes(S) + ' ';
    end;
  end;

  if Pos('/DIR', Params) = 0 then
    Params := Params + ExpandConstant('/DIR="{app}" ');
  { ... and add selected language }
  if Pos('/LANG', Params) = 0 then
    Params := Params + '/LANG=' + ActiveLanguage + ' ';
  if Pos('/ELEVATE', Params) = 0 then
    Params := Params + '/ELEVATE';

  Log(Format('Elevating setup with parameters [%s]', [Params]));
  instance := ShellExecute(0, 'runas', ExpandConstant('{srcexe}'), Params, '', SW_SHOW);
  Log(Format('Running elevated setup returned [%d]', [instance]));
  Result := (instance > 32);
  { if elevated executing of this setup succeeded, then... }
  if Result then
  begin
    Log('Elevation succeeded');
  end
    else
  begin
    Log(Format('Elevation failed [%s]', [SysErrorMessage(instance)]));
  end;
end;

///<remarks>
///Adapted from http://kynosarges.org/DotNetVersion.html; used during
///the <see cref="InitializeSetup"> event function to ensure that the
///.NET framework is present on the computer.
///</remarks>
function IsDotNetDetected(version: string; service: cardinal): boolean;
// Indicates whether the specified version and service pack of the .NET Framework is installed.
//
// version -- Specify one of these strings for the required .NET Framework version:
//    'v1.1'          .NET Framework 1.1
//    'v2.0'          .NET Framework 2.0
//    'v3.0'          .NET Framework 3.0
//    'v3.5'          .NET Framework 3.5
//    'v4\Client'     .NET Framework 4.0 Client Profile
//    'v4\Full'       .NET Framework 4.0 Full Installation
//    'v4.5'          .NET Framework 4.5
//    'v4.5.1'        .NET Framework 4.5.1
//    'v4.5.2'        .NET Framework 4.5.2
//    'v4.6'          .NET Framework 4.6
//    'v4.6.1'        .NET Framework 4.6.1
//    'v4.6.2'        .NET Framework 4.6.2
//    'v4.7'          .NET Framework 4.7
//
// service -- Specify any non-negative integer for the required service pack level:
//    0               No service packs required
//    1, 2, etc.      Service pack 1, 2, etc. required
var
    key, versionKey: string;
    install, release, serviceCount, versionRelease: cardinal;
    success: boolean;
begin
    versionKey := version;
    versionRelease := 0;

    // .NET 1.1 and 2.0 embed release number in version key
    if version = 'v1.1' then begin
        versionKey := 'v1.1.4322';
    end else if version = 'v2.0' then begin
        versionKey := 'v2.0.50727';
    end

    // .NET 4.5 and newer install as update to .NET 4.0 Full
    else if Pos('v4.', version) = 1 then begin
        versionKey := 'v4\Full';
        case version of
          'v4.5':   versionRelease := 378389;
          'v4.5.1': versionRelease := 378675; // 378758 on Windows 8 and older
          'v4.5.2': versionRelease := 379893;
          'v4.6':   versionRelease := 393295; // 393297 on Windows 8.1 and older
          'v4.6.1': versionRelease := 394254; // 394271 before Win10 November Update
          'v4.6.2': versionRelease := 394802; // 394806 before Win10 Anniversary Update
          'v4.7':   versionRelease := 460798; // 460805 before Win10 Creators Update
        end;
    end;

    // installation key group for all .NET versions
    key := 'SOFTWARE\Microsoft\NET Framework Setup\NDP\' + versionKey;

    // .NET 3.0 uses value InstallSuccess in subkey Setup
    if Pos('v3.0', version) = 1 then begin
        success := RegQueryDWordValue(HKLM, key + '\Setup', 'InstallSuccess', install);
    end else begin
        success := RegQueryDWordValue(HKLM, key, 'Install', install);
    end;

    // .NET 4.0 and newer use value Servicing instead of SP
    if Pos('v4', version) = 1 then begin
        success := success and RegQueryDWordValue(HKLM, key, 'Servicing', serviceCount);
    end else begin
        success := success and RegQueryDWordValue(HKLM, key, 'SP', serviceCount);
    end;

    // .NET 4.5 and newer use additional value Release
    if versionRelease > 0 then begin
        success := success and RegQueryDWordValue(HKLM, key, 'Release', release);
        success := success and (release >= versionRelease);
    end;

    result := success and (install = 1) and (serviceCount >= service);
end;

///<remarks>
///Helper function used in non-code sections to set the default directory
///to be installed, based on the elevation context.
///</remarks>
function GetDefaultDirName(Param: string): string;
begin
  if IsElevated() then
  begin
    Result := ExpandConstant('{commonappdata}{\}{#AppName}');
  end
    else
  begin
    Result := ExpandConstant('{localappdata}{\}{#AppName}')
  end;
end;

///<remarks>
///Helper function used in non-code sections to provide the path that was
///either set by installer by default or customized by the user.
///</remarks>
function GetInstallPath(Unused: string): string;
begin
  result := ExpandConstant('{app}');
end;

///<remarks>
///Helper function used in non-code sections to provide the full path to
///the Rubberduck main DLL, based on the same path as <see cref="GetInstallPath" />
///</remarks>
function GetDllPath(Unused: string): string;
begin
  result := ExpandConstant('{app}') + '\Rubberduck.dll';
end;

///<remarks>
///Helper function used in non-code sections to provide the full path to
///the 32-bit type library for Rubberduck main DLL, based on teh same path
///as <cref="GetInstallPath" />
///</remarks>
function GetTlbPath32(Unused: string): string;
begin
  result := ExpandConstant('{app}') + '\Rubberduck.x32.tlb';
end;

///<remarks>
///Same as <see cref="GetTlbPath32" /> but for 64-bit.
///</remarks>
function GetTlbPath64(Unused: string): string;
begin
  result := ExpandConstant('{app}') + '\Rubberduck.x64.tlb';
end;

///<remarks>
///Helper function used in the Registry section to indicate whether
///the item should be installed. For example, HKLM registry requires
///elevated context, so the function must return true, while HKCU
///counterpart will expect the opposite to be installed. This prevents
///comparable item being installed into both places.
///</remarks>
function InstallAllUsers():boolean;
begin
  result := ShouldInstallAllUsers and IsElevated();
end;

///<remarks>
///Helper function used in the File section to assess whether an
///file(s) should be installed based on whether there's privilege
///to do so. This guards against the case of both the non-elevated
///installer and the elevated installer installing same files into
///same place, which will cause problems. This function ensures only
///one or other mode will actually install the files.
///</remarks>
function CheckShouldInstallFiles():boolean;
begin
  if ShouldInstallAllUsers then
    result := IsElevated()
  else
    result := true;
end;

///<remarks>
///Used by <see cref="RegisterAddIn" />, passing in parameters to actually create
///the per-user registry entries to enable VBE addin for that user.
///<remarks>
procedure RegisterAddinForIDE(const iRootKey: Integer; const sAddinSubKey: String; const sProgIDConnect: String; const bIncludeCommandLine: boolean);
begin
   RegWriteStringValue(iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'FriendlyName', '{#AppName}');
   RegWriteStringValue(iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'Description' , '{#AppName}');
   RegWriteDWordValue (iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'LoadBehavior', 3);
   if bIncludeCommandLine then
   begin
    RegWriteDWordValue (iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'CommandLineSafe', 0);
   end;
 end;

///<remarks>
///Unregisters the same keys, similar to <see cref="RegisterAddinForIDE" />
///</remarks>
procedure UnregisterAddinForIDE(const iRootKey: Integer; const sAddinSubKey: String; const sProgIDConnect: String);
begin
   if RegKeyExists(iRootKey, sAddinSubKey + '\' + sProgIDConnect) then
      RegDeleteKeyIncludingSubkeys(iRootKey, sAddinSubKey + '\' + sProgIDConnect);
end;

///<remarks>
///Called after successfully installing, including via the elevated installer
///to register the VBE addin.
///</remarks>
procedure RegisterAddin();
begin
    if IsWin64() then
    begin
      RegisterAddinForIDE(HKCU32, 'Software\Microsoft\VBA\VBE\6.0\Addins', '{#AddinProgId}', false);
      RegisterAddinForIDE(HKCU64, 'Software\Microsoft\VBA\VBE\6.0\Addins64', '{#AddinProgId}', false);
      RegisterAddinForIDE(HKCU32, 'Software\Microsoft\Visual Basic\6.0\Addins', '{#AddinProgId}', true);
    end
      else
    begin
      RegisterAddinForIDE(HKCU, 'Software\Microsoft\VBA\VBE\6.0\Addins', '{#AddinProgId}', false);
      RegisterAddinForIDE(HKCU, 'Software\Microsoft\Visual Basic\6.0\Addins', '{#AddinProgId}', true);
    end;
end;

///<remarks>
///Delete registry keys created by <see cref="RegisterAddin" />
///</remarks>
procedure UnregisterAddin();
begin
  if IsWin64() then
  begin
    UnregisterAddinForIDE(HKCU64, 'Software\Microsoft\VBA\VBE\6.0\Addins64', '{#AddinProgId}');
    UnregisterAddinForIDE(HKCU32, 'Software\Microsoft\VBA\VBE\6.0\Addins', '{#AddinProgId}');
    UnRegisterAddinForIDE(HKCU32, 'Software\Microsoft\Visual Basic\6.0\Addins', '{#AddinProgId}');
  end
    else
  begin
    UnregisterAddinForIDE(HKCU, 'Software\Microsoft\VBA\VBE\6.0\Addins', '{#AddinProgId}');
    UnRegisterAddinForIDE(HKCU, 'Software\Microsoft\Visual Basic\6.0\Addins', '{#AddinProgId}');
  end;
end;

///<remarks>
///Generate AppId based on whether it's going to be
///per-user or for all users. This enable separate
///install/uninstall of each mode. Used by AppId
///directive in the [Setup] section.
///</remarks>
///<param name="AppMode">
///If value is 'peruser', then returns per-user AppId
///If value is 'everyone', then returns everyone AppId
///Otherwise if left blank or contains invalid value, returns
///the AppId based on ShouldInstallAllUsers
///(e.g. based on user's selection.)
///</param>
function GetAppId(AppMode: string): string;
begin
  if AppMode = EveryoneAppMode then
    result := EveryoneAppId
  else if AppMode = PerUserAppMode then
    result := PerUserAppId
  else
    if ShouldInstallAllUsers then
      result := EveryoneAppId
    else
      result := PerUserAppId
end;

///<remarks>
///Used to help disambiguate multiple installs of Rubberduck
///by providing a suffix to indicate which mode it is
///</remarks>
function GetAppSuffix(): string;
begin
  if ShouldInstallAllUsers then
    result := ExpandConstant('{cm:Everyone}')
  else
    result := ExpandConstant('{cm:PerUser}');
end;

///<remakrs>
///Provide a suffixed name of uninstaller to help identify what mode
///the version is installed in.
///</remarks>
function GetUninstallDisplayName(Unused: string): string;
begin
  result := ExpandConstant('{#AppName} (') + GetAppSuffix() + ') {#AppVersion}';
end;

///<remarks>
///Prevent creating an uninstaller from the non-elevated installer
///The elevated installer will do it instead. Otherwise, we get
///weird behavior & errors when uninstalling mixed mode.
///</remarks>
function ShouldCreateUninstaller(): boolean;
begin
  if not IsElevated() and ShouldInstallAllUsers then
    result:= false
  else
    result:= true;
end;

///<remarks>
///Deterimine if there is a previous version of Rubberduck installed
///All uninstaller will store a registry key with the AppId, so we can
///use AppId to detect previous versions.
///</remarks>
///<param name="AppId">
///If contains a valid AppId, attempt to get uninstaller for that.
///If left blank, attempt to get uninstaller for current mode.
///</param>
function GetUninstallString(AppId: string): String;
var
  sAppId: string;
  sUnInstPath: String;
  sUnInstallString: String;
begin
  if AppId = '' then
    sAppId := GetAppId('')
  else
    sAppId := AppId;

  sUnInstPath := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\' + sAppId + '_is1';
  Log('Looking in registry: ' + sUnInstPath);
  sUnInstallString := '';
  if not RegQueryStringValue(HKLM, sUnInstPath, 'UninstallString', sUnInstallString) then
    RegQueryStringValue(HKCU, sUnInstPath, 'UninstallString', sUnInstallString);
  Log('Result of registry query: ' + sUnInstallString);
  Result := sUnInstallString;
end;

///<remarks>
///Encapuslates the check for previous versions
///from <see cref="GetUninstallString" /> as a
///boolean result
///</remarks>
function IsUpgrade(): boolean;
begin
  result := (GetUninstallString('') <> '');
end;

///<reamrks>
///An expanded version of IsUpgrade function, usuable only
///within [Code] section that additional provides information
///for all modes of installation.
///</remarks>
///<returns>
///An integer representing the fake ***IsInstalled enum
///</returns>
function IsUpgradeApp(): integer;
var
  PerUserExists: boolean;
  EveryoneExists: boolean;
begin
    PerUserExists := (GetUninstallString(PerUserAppId) <> '');
    EveryoneExists := (GetUninstallString(EveryoneAppId) <> '')

    if PerUserExists and EveryoneExists then
      result := BothAreInstalled
    else if PerUserExists and not EveryoneExists then
      result := PerUserIsInstalled
    else if not PerUserExists and EveryoneExists then
      result := EveryoneIsInstalled
    else
      result := NoneIsInstalled;
end;

///<remarks>
///Perform uninstall of old versions. Called in the
///<see cref="CurStepChanged" /> event function and only when
///a previous version was detected.
///</remarks>
///<param name="AppId">
///If non-blank, uninstall the specific AppId.
///Otherwise, use the selected mode to uninstall.
///</param>
///<returns>
/// Return Values:
///   1 - uninstall string is empty
///   2 - error executing the UnInstallString
///   3 - successfully executed the UnInstallString
///</returns>
function UnInstallOldVersion(AppId: string): integer;
var
  sUnInstallString: string;
  iResultCode: integer;
begin
  // default return value
  Result := 0;

  // get the uninstall string of the old app
  sUnInstallString := GetUninstallString(AppId);
  if sUnInstallString <> '' then begin
    sUnInstallString := RemoveQuotes(sUnInstallString);
    if Exec(sUnInstallString, '/SILENT /NORESTART /SUPPRESSMSGBOXES','', SW_HIDE, ewWaitUntilTerminated, iResultCode) then
      Result := 3
    else
      Result := 2;
  end else
    Result := 1;
end;

///<remarks>
///Encapuslates the UI interaction with user to see whether to uninstall
///the previous verison of a given mode.
///<remakrs>
///<param name = "AppMode">
///Indicate which app mode to uninstall
///</param>
///<returns>
///Boolean indicating whether it was uninstalled succesffully.
///</returns>
function PromptToUninstall(AppMode: string): boolean;
var
  ErrorCode: integer;
  LocalizedMode: string;
begin
  Log('A previous version of ' + AppMode + ' was detected; prompting the user whether to uninstall');
  if AppMode = EveryoneAppMode then
    LocalizedMode := ExpandConstant('{cm:Everyone}')
  else if AppMode = PerUserAppMode then
    LocalizedMode := ExpandConstant('{cm:PerUser}');
  if IDYES = MsgBox(Format(ExpandConstant('{cm:UninstallOldVersionPrompt}'), [LocalizedMode]), mbConfirmation, MB_YESNO) then
  begin
    ErrorCode := -1;
    ErrorCode := UnInstallOldVersion(GetAppId(AppMode));
    Log(Format('The result of UninstallOldVersion for %s was %d.', [AppMode, ErrorCode]));

    if ErrorCode <> 3 then
      MsgBox(ExpandConstant('{cm:UninstallOldVersionFail}'), mbError, MB_OK);
    result := (ErrorCode = 3);
  end
    else
  begin
    Log('Uninstall of previous version (%s) was declined by the user.');
    result := false;
  end;
end;

// Event functions called by Inno Setup
//
// NOTE: the ordering should be preserved to indicate the general sequence
//       that the Inno Setup will undergo for each event it calls. Note that
//       for certain events, it may be called more than once.

///<remarks>
///This is the first event of the installer, fires prior to the wizard
///being initialized. This is primarily used to validate that the
///pre-requisites are met, in this case, pre-existence of .NET framework.
///</remarks>
function InitializeSetup(): Boolean;
var
   ErrorCode: Integer;
begin
  // MS .NET Framework 4.6 must be installed for this application to work.
  if not IsDotNetDetected('v4.6', 0) then
  begin
    Log('User does not have the prerequisite .NET framework installed');
    MsgBox(ExpandConstant('{cm:NETFramework46NotInstalled}'), mbCriticalError, mb_Ok);
    ShellExec('open', 'https://www.microsoft.com/net/download/dotnet-framework-runtime/net46', '', '', SW_SHOW, ewNoWait, ErrorCode);
    Result := False;
  end
  else
  begin
    Log('.Net v4.5 Framework was found on the system');
    Result := True;
  end;
end;

///<remarks>
///The second event of installer allow us to customize the wizard by
///assessing whether we were launched in elevated context from an
///non-elevated installer; <see cref="HasElevateSwitch" />. We then
///set up the <see cref="InstallForWhoOptionPage" /> and
///<see cref="RegisterAddInOptionPage" /> pages. In both cases, their
///behavior differs depending on whether we are elevated, and need to be
///configured accordingly.
///</remarks>
procedure InitializeWizard();
begin
  HasElevateSwitch := CmdLineParamExists('/ELEVATE');

  Log(Format('HasElevateSwitch: %d', [HasElevateSwitch]));
  Log(Format('IsElevated: %d', [IsElevated()]));

  //Assume we are installing for all users if we were elevated
  ShouldInstallAllUsers := HasElevateSwitch;

  InstallForWhoOptionPage :=
    CreateInputOptionPage(
      wpWelcome,
      ExpandConstant('{cm:InstallPerUserOrAllUsersCaption}'),
      ExpandConstant('{cm:InstallPerUserOrAllUsersMessage}'),
      ExpandConstant('{cm:InstallPerUserOrAllUsersAdminDescription}'),
      True, False);

  InstallForWhoOptionPage.Add(ExpandConstant('{cm:InstallPerUserOrAllUsersAdminButtonCaption}'));
  InstallForWhoOptionPage.Add(ExpandConstant('{cm:InstallPerUserOrAllUsersUserButtonCaption}'));

  if PerUserOnly then
  begin
    InstallForWhoOptionPage.Values[1] := true;
    InstallForWhoOptionPage.CheckListBox.ItemEnabled[0] := false;
  end
    else if IsElevated() then
  begin
    InstallForWhoOptionPage.Values[0] := true;
    InstallForWhoOptionPage.CheckListBox.ItemEnabled[1] := false;
  end
    else
  begin
    InstallForWhoOptionPage.Values[1] := true;
  end;

  RegisterAddInOptionPage :=
    CreateInputOptionPage(
      wpInstalling,
      ExpandConstant('{cm:RegisterAddInCaption}'),
      ExpandConstant('{cm:RegisterAddInMessage}'),
      ExpandConstant('{cm:RegisterAddInDescription}'),
      false, false);

  RegisterAddInOptionPage.Add(ExpandConstant('{cm:RegisterAddInButtonCaption}'));
  RegisterAddInOptionPage.Values[0] := true;
end;

///<remarks>
///This is called prior to load of the next page in question
///and is fired for each page we are about to visit.
///Normally we don't skip unless we have the elevate switch in
///which case the elevated installer already has all the user's
///input so it needs not ask users for those again, making it quick
///to run the elevated installer when elevated from the non-elevated
///installer. Otherwise, we verify whether we need to show the
///<see cref="RegisterAddInOptionPage" /> which won't be if the installer
///is elevated (irrespective whether by the switch or directly by user)
///as installing addin registry keys will not work under an elevated context.
///</remarks>
function ShouldSkipPage(PageID: Integer): Boolean;
begin
  // if we've executed this instance as elevated, skip pages unless we're
  // on the directory selection page
  Result := not PagesSkipped and HasElevateSwitch and IsElevated() and (PageID <> wpReady);
  // if we've reached the Ready page, set our flag variable to avoid skipping further pages.
  if not Result then
  begin
    Log('PageSkipped set to true now');
    PagesSkipped := True;
  end;

  // We don't need to show the users finished panel from the elevated installer
  // if they already have the non-elevated installer running.
  if (PageID = wpFinished) and HasElevateSwitch then
  begin
    Log('Skipping Finished page because we are running with /ELEVATE switch');
    Result := true;
  end;
end;

///<remarks>
///This is called once the user has clicked on the Next button on the wizard
///and is called for each page. Thus, we have basically a switch for different
///page, then we assess what we need to do.
///<remarks>
function NextButtonClick(CurPageID: Integer): Boolean;
var
  RetVal: HINSTANCE;
  UpgradeResult: integer;
begin
  // Prevent accidental extra clicks
  Wizardform.NextButton.Enabled := False;

  // We should assume true because a false value will cause the
  // installer to stay on the same page, which may not be desirable
  // due to several branching in this prcocedure.
  Result := true;

  // We need to assess whether user might have selected some other directory
  // and whether we need to elevate in order to write to it. If elevation is
  // required, we need to confirm with the users. The actual elevation comes later.
  if CurPageID = wpSelectDir then
  begin
    if not ShouldInstallAllUsers then
    begin
      if not HaveWriteAccessToApp() then
      begin
        if IDYES = MsgBox(ExpandConstant('{cm:ElevationRequiredForSelectedFolderWarning}'), mbConfirmation, MB_YESNO) then
        begin
          Log('Setting ShouldIntallAllUsers to true because we need elevation to write to selected directory');
          ShouldInstallAllUsers := true;
        end
          else
        begin
          Log('User declined to elevate the permission so we will remain on wpSelectDir page to allow user to change the selection');
          Result := false
        end;
      end;
    end;
  end
    // If we need elevation, we will invoke the <see cref="Elevate" />
    // and verify we are able to do so. Failure should keep user on the
    // same page as the user can retry or cancel out from the non-elevated
    // installer.
    //
    // We also need to verify there are no previous versions and if there are,
    // to uninstall them.
    else if CurPageID = wpReady then
  begin
    // Log all output of functions called by non-code sections
    Log('GetInstallPath: ' + GetInstallPath(''));
    Log('GetDllPath: ' + GetDllPath(''));
    Log('GetTlbPath32: ' + GetTlbPath32(''));
    Log('GetTlbPath64: ' + GetTlbPath64(''));
    Log(Format('InstallAllUsers: %d', [InstallAllUsers()]));
    Log(Format('CheckShouldInstallFiles: %d', [CheckShouldInstallFiles()]));
    Log(Format('GetAppId: %s', [GetAppId('')]));
    Log(Format('AppSuffix: %s', [GetAppSuffix()]));
    Log(Format('ShouldCreateUninstaller: %d', [ShouldCreateUninstaller()]));
    Log(Format('ShouldInstallAllUsers variable: %d', [ShouldInstallAllUsers]));

    if not IsElevated() and ShouldInstallAllUsers then
    begin
      Log('All-users install is required but we don''t have privilege; requesting elevation...');
      if not Elevate() then
      begin
        Log('Elevation failed or was cancelled; we cannot continue.');
        Result := False;
        MsgBox(Format(ExpandConstant('{cm:ElevationRequestFailMessage}'), [RetVal]), mbError, MB_OK);
      end;
    end;
  end
    // We should set the selected directory (WizardForm.DirEdit) with
    // appropriate default directory depending on whether the user is
    // running the installer as elevated or not.
    else if CurPageID = InstallForWhoOptionPage.ID then
  begin
    if InstallForWhoOptionPage.Values[1] then
    begin
      ShouldInstallAllUsers := False;
      WizardForm.DirEdit.Text := ExpandConstant('{localappdata}{\}{#AppName}')
      Log('ShouldInstallAllUsers set to false because we chose You Only option');
    end
      else
    begin
      ShouldInstallAllUsers := True;
      WizardForm.DirEdit.Text := ExpandConstant('{commonappdata}{\}{#AppName}');
      Log('ShouldInstallAllUsers set to true because we chose All Users options');
    end;
    Log(Format('Selected default directory: %s', [WizardForm.DirEdit.Text]));

    //We need to check whether there's previous version to be uninstalled
    //We have to do it early enough or the installer will try to use the
    //previous version's directory, which will basically ignore the directory
    //being selected.
    UpgradeResult := IsUpgradeApp();

    if (UpgradeResult > NoneIsInstalled) then
    begin
      if IsElevated() then
      begin
        // Uninstall per user; continue regardless of result
        if (UpgradeResult = PerUserIsInstalled) or (UpgradeResult = BothAreInstalled) then
          PromptToUninstall(PerUserAppMode);

        // Uninstall all users; must succeed to continue
        if (UpgradeResult = EveryoneIsInstalled) or (UpgradeResult = BothAreInstalled) then
          result := PromptToUninstall(EveryoneAppMode);
      end
        else
      begin
        if ShouldInstallAllUsers then
        begin
          // We're asking to install for all users; so we must uninstall old version to continue
          if (UpgradeResult = EveryoneIsInstalled) or (UpgradeResult = BothAreInstalled) then
            result := PromptToUninstall(EveryoneAppMode);
        end
          else
        begin
          // Warn RE: multiple install (if both is installed, they already were warned)
          if (UpgradeResult = EveryoneIsInstalled) then
            result := (IDYES = MsgBox(ExpandConstant('{cm:WarnInstallPerUserOverEveryone}'), mbConfirmation, MB_YESNO));
        end;

        // Uninstall per user; must succeed to continue
        if result and ((UpgradeResult = PerUserIsInstalled) or (UpgradeResult = BothAreInstalled)) then
          result := PromptToUninstall(PerUserAppMode);
      end;
    end;
  end
    // if the user has allowed registration of the IDE (default) from our
    // custom page we should run RegisterAdd()
    else if CurPageID = RegisterAddInOptionPage.ID then
  begin
    if RegisterAddInOptionPage.Values[0] then
    begin
      Log('Addin registration was requested and will be performed');
      RegisterAddIn();
    end
      else
    begin
      Log('Addin registration was declined because the user unchecked the checkbox');
    end;
  end;

  // Re-enable the button disabled at start of procedure
  Wizardform.NextButton.Enabled := True;
end;

///<remarks>
///The event function is called when wizard reaches the ready to install page.
///Because we may or may not launch an elevated installer which will show similar
///page, we need to help to make clear to the user what the installer(s) will be
///doing by adding the extra custom messages accordingly to the page.
///</remarks>
function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo, MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
var
  output: String;
begin
  if IsElevated() then
  begin
    output := output + ExpandConstant('{cm:WillExecuteAdminInstall}') + NewLine + NewLine;
  end
    else
  begin
    if ShouldInstallAllUsers then
    begin
      output := output + ExpandConstant('{cm:WillLaunchAdminInstall}') + NewLine + NewLine;
    end
      else
    begin
      output := output + ExpandConstant('{cm:WillInstallForCurrentUser}') + NewLine + NewLine;
    end;
  end;

  output := output + MemoDirInfo + NewLine;

  result := output;
end;

///<remarks>
///Allow customization of the uninstall form
///Specificallly, show the version in the form
///</remarks>
procedure InitializeUninstallProgressForm();
var
  TempString: string;
begin
  TempString := UninstallProgressForm.Caption;
  Log('Original Uninstall caption: ' + TempString);
  StringChange(TempString, '{#AppName}', '{#AppName} {#AppVersion}');
  Log('Modified Uninstall caption: ' + TempString);
  UninstallProgressForm.Caption := TempString;
end;

///<remarks>
///Called during uninstall, once for each step but for our purpose, we are
///interested in only one step doing the actual uninstall.
///
///As a rule, the addin registration should be always uninstalled; there is no
///purpose in having the addin registered when the DLL gets uninstalled. Note
///this is also unconditional - it will uninstall the related registry keys.
///</remarks>
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    UnregisterAddin();
  end;
end;
