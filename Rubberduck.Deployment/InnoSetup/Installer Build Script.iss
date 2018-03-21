#pragma include __INCLUDE__ + ";" + SourcePath + "\Includes\"

#define protected
#ifndef Config
#define Config "Debug"
#endif
#define BuildDir ExtractFileDir(ExtractFileDir(SourcePath)) + "\bin\" + Config + "\"
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
#define License SourcePath + "\License.rtf"
#define OutputDirectory SourcePath + "Installers\"
#define AddinProgId "Rubberduck.Extension"
#define AddinCLSID "8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66"

; Output the defined constants to aid in verification
#pragma message "Include: " + __INCLUDE__
#pragma message "Config: " + Config
#pragma message "SourcePath: " + SourcePath
#pragma message "BuildDir: " + BuildDir
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
; TODO this CLSID should match the one used by the current installer.
AppId={{979AFF96-DD9E-4FC2-802D-9E0C36A60D09}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={code:GetDefaultDirName}
DefaultGroupName=Rubberduck
DisableProgramGroupPage=yes
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

; should be a 164 x 314 bitmap
WizardImageFile=InstallerBitMap.bmp

[Languages]
Name: "English"; MessagesFile: "compiler:Default.isl"
Name: "French"; MessagesFile: "compiler:Languages\French.isl"
Name: "German"; MessagesFile: "compiler:Languages\German.isl"

[Files]
; Install the correct bitness binaries.
; Source: "{#BuildDir}*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs replacesameversion; Permissions: users-readexec;
Source: "{#BuildDir}*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs replacesameversion; Excludes: "Rubberduck.Deployment.*,Rubberduck.dll.xml,Rubberduck.x32.tlb.xml,{#AddinDLL},\NativeBinaries"; Check: CheckShouldInstallFiles
Source: "{#BuildDir}{#AddinDLL}"; DestDir: "{app}"; Flags: ignoreversion replacesameversion; Check: CheckShouldInstallFiles

[Registry]
; DO NOT attempt to register VBE Add-In with this section. It doesn't work
; Use [Code] section (RegisterAddIn procedure) to register the entries instead.
#include <Rubberduck.reg.iss>

[UninstallDelete]
Type: filesandordirs; Name: "{localappdata}\{#AppName}"

[CustomMessages]
; TODO add additional languages here.
English.ProgramOnTheWeb=Rubberduck VBA website
English.UninstallProgram=Uninstall Rubberduck
English.NETFramework40NotInstalled=Microsoft .NET Framework 4.0 installation was not detected.
English.InstallPerUserOrAllUsersCaption=Choose installation options
English.InstallPerUserOrAllUsersMessage=Who should this application be installed for? 
English.InstallPerUserOrAllUsersAdminDescription=Please select whether you wish to make this software available for all users or just yourself.
English.InstallPerUserOrAllUsersAdminButtonCaption=&Anyone who use this computer
English.InstallPerUserOrAllUsersUserButtonCaption=&You only
English.ElevationRequiredForSelectedFolderWarning=The selected folder requires administrative privilege to write. Do you want to install for all users?
English.ElevationRequestFailMessage=Elevating of this setup failed. Code: %d
English.RegisterAddInCaption=Register VBE Add-in
English.RegisterAddInMessage=Perform user registration of VBE addin.
English.RegisterAddInDescription=VBE Add-ins are registered on a per-user basis, even if the VBE add-in was installed for all users. Therefore, this must be executed on a per-user basis.
English.RegisterAddInButtonCaption=Register the Rubberduck VBE Add-in
English.WillExecuteAdminInstall=Rubberduck Add-In will be available to all users.%n%nNOTE: each user individually must register the Rubberduck Add-In as%nthis is a per-user setting and cannot be deployed to all users.
English.WillLaunchAdminInstall=The installer will request for admin privilege to install for all users and will%nresume afterward to perform the add-in registration.
English.WillInstallForCurrentUser=Rubberduck Add-In will be made available to the current user only and will not require admin privileges.

[Icons]
Name: "{group}\{cm:ProgramOnTheWeb,{#AppName}}"; Filename: "{#AppURL}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"

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
  Result := IsAdminLoggedOn or IsPowerUserLoggedOn;
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
  FileName: string;
begin
{
  FileName := AddBackslash(WizardDirValue) + 'writetest.tmp';
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
}
Result := true;
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
  waitSignal: DWORD;
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

///<remarks>
///Helper function used in non-code sections to set the default directory
///to be installed, based on the elevation context.
///</remarks>
function GetDefaultDirName(Param: string): string;
begin
  if IsElevated() then
  begin
    Result := ExpandConstant('{pf}{#AppName}');
  end
    else
  begin
    Result := ExpandConstant('{userappdata}{#AppName}');
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
  result := ShouldInstallAllUsers and HasElevateSwitch;
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
    result := HasElevateSwitch and IsElevated()
  else
    result := true;
end;

///<remarks>
///Used by <see cref="RegisterAddIn" />, passing in parameters to actually create
///the per-user registry entries to enable VBE addin for that user.
///<remarks>
procedure RegisterAddinForIDE(const iRootKey: Integer; const sAddinSubKey: String; const sProgIDConnect: String);
begin
   RegWriteStringValue(iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'FriendlyName', '{#AppName}');
   RegWriteStringValue(iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'Description' , '{#AppName}');
   RegWriteDWordValue (iRootKey, sAddinSubKey + '\' + sProgIDConnect, 'LoadBehavior', 3);
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
///to register the VBE addin. Should be never run under elevated context
///or the registration may not work as expected.
///</remarks>
procedure RegisterAddin();
begin
  if not IsElevated() then
  begin
    RegisterAddinForIDE(HKCU32, 'Software\Microsoft\VBA\VBE\6.0\Addins', '{#AddinProgId}');

    if IsWin64() then 
      RegisterAddinForIDE(HKCU64, 'Software\Microsoft\VBA\VBE\6.0\Addins64', '{#AddinProgId}');
  end;
end;

///<remarks>
///Delete registry keys created by <see cref="RegisterAddin" />
///</remarks>
procedure UnregisterAddin();
begin
  UnregisterAddinForIDE(HKCU32, 'Software\Microsoft\VBA\VBE\6.0\Addins', '{#AddinProgId}');
  if IsWin64() then 
    UnregisterAddinForIDE(HKCU64, 'Software\Microsoft\VBA\VBE\6.0\Addins64', '{#AddinProgId}');
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

  InstallForWhoOptionPage :=
    CreateInputOptionPage(
      wpWelcome,
      ExpandConstant('{cm:InstallPerUserOrAllUsersCaption}'), 
      ExpandConstant('{cm:InstallPerUserOrAllUsersMessage}'),
      ExpandConstant('{cm:InstallPerUserOrAllUsersAdminDescription}'),
      True, False);

  InstallForWhoOptionPage.Add(ExpandConstant('{cm:InstallPerUserOrAllUsersAdminButtonCaption}'));
  InstallForWhoOptionPage.Add(ExpandConstant('{cm:InstallPerUserOrAllUsersUserButtonCaption}'));

  if IsAdminLoggedOn then
  begin
    InstallForWhoOptionPage.Values[0] := True;
    InstallForWhoOptionPage.CheckListBox.ItemEnabled[1] := False;
  end
    else
  begin
    InstallForWhoOptionPage.Values[1] := True;
  end;

  RegisterAddInOptionPage :=
    CreateInputOptionPage(
      wpInstalling,
      ExpandConstant('{cm:RegisterAddInCaption}'),
      ExpandConstant('{cm:RegisterAddInMessage}'),
      ExpandConstant('{cm:RegisterAddInDescription}'),
      False, False);

  RegisterAddInOptionPage.Add(ExpandConstant('{cm:RegisterAddInButtonCaption}'));
  RegisterAddInOptionPage.CheckListBox.ItemEnabled[0] := not IsElevated();
  RegisterAddInOptionPage.Values[0] := not IsElevated();
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
    PagesSkipped := True;

  // If the installer is elevated, we cannot register the addin so we must skip the
  // custom page.
  if (PageID = RegisterAddInOptionPage.ID) and IsElevated() then
    Result := true;

  // We don't need to show the users finished panel from the elevated installer
  // if they already have the non-elevated installer running.
  if (PageID = wpFinished) and HasElevateSwitch then
    Result := true;
end;

///<remarks>
///This is called once the user has clicked on the Next button on the wizard
///and is called for each page. Thus, we have basically a switch for different
///page, then we assess what we need to do.
///<remarks>
function NextButtonClick(CurPageID: Integer): Boolean;
var
  Params: string; 
  RetVal: HINSTANCE;
  Respone: integer;
begin
  // We should assume true because a false value will cause the 
  // installer to stay on the same apge, which may not be desirable
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
          ShouldInstallAllUsers := true;
        end
          else
        begin
          Result := false
        end;
      end;
    end;
  end
    // If we need elevation, we will invoke the <see cref="Elevate" />
    // and verify we are able to do so. Failure should keep user on the
    // same page as the user can retry or cancel out from the non-elevated
    // installer.
    else if CurPageID = wpReady then
  begin
    if not IsElevated() and ShouldInstallAllUsers then
    begin
      if not Elevate() then
      begin
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
    end
      else
    begin
      ShouldInstallAllUsers := True;
      WizardForm.DirEdit.Text := ExpandConstant('{commonappdata}{\}{#AppName}');
    end;
  end
    // if the user has allowed registration of the IDE (default) from our
    // custom page we should run RegisterAdd()
    else if CurPageID = RegisterAddInOptionPage.ID then
  begin
    if not IsElevated() and RegisterAddInOptionPage.Values[0] then
      RegisterAddIn();
  end;
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
///As a rule, the addin registration should be always uninstalled; there is no
///purpose in having the addin registered when the DLL gets uninstalled. Note
///this is also unconditional - it will uninstall the related registry keys.
///</remarks>
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then UnregisterAddin();
end;
