;
; Script generated by the ASCOM Driver Installer Script Generator 6.0.0.0
; Generated by Pavle Gartner on 26.11.2015 (UTC)
;
[Setup]
AppID={{8231a81f-6547-43ad-b4b0-9b3206debc31}
AppName=ASCOM Pulsar2 Telescope Driver
AppVerName=ASCOM Pulsar2 Telescope Driver 1.1.0
AppVersion=1.1.0
AppPublisher=Pavle Gartner <pavle.gartner@gmail.com>
AppPublisherURL=mailto:pavle.gartner@gmail.com
AppSupportURL=http://tech.groups.yahoo.com/group/ASCOM-Talk/
AppUpdatesURL=http://ascom-standards.org/
VersionInfoVersion=1.1.0
MinVersion=0,5.0.2195sp4
DefaultDirName="{cf}\ASCOM\Telescope"
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir="."
OutputBaseFilename="Pulsar2 Setup"
Compression=lzma
SolidCompression=yes
; Put there by Platform if Driver Installer Support selected
WizardImageFile="C:\Program Files (x86)\ASCOM\Platform 6 Developer Components\Installer Generator\Resources\WizardImage.bmp"
LicenseFile="C:\Program Files (x86)\ASCOM\Platform 6 Developer Components\Installer Generator\Resources\CreativeCommons.txt"
; {cf}\ASCOM\Uninstall\Telescope folder created by Platform, always
UninstallFilesDir="{cf}\ASCOM\Uninstall\Telescope\Pulsar2"

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Dirs]
Name: "{cf}\ASCOM\Uninstall\Telescope\Pulsar2"
; TODO: Add subfolders below {app} as needed (e.g. Name: "{app}\MyFolder")

[Files]
Source: "D:\Projects\Pulsar2.ASCOM\Pulsar2Driver\bin\Release\ASCOM.Pulsar2.Telescope.dll"; DestDir: "{app}"
; Require a read-me HTML to appear after installation, maybe driver's Help doc
Source: "D:\Projects\Pulsar2.ASCOM\Pulsar2Driver\ReadMe.htm"; DestDir: "{app}"; Flags: isreadme
; TODO: Add other files needed by your driver here (add subfolders above)
Source: "D:\Projects\Pulsar2.ASCOM\Pulsar2Driver\readme.jpg"; DestDir: "{app}";


; Only if driver is .NET
[Run]
; Only for .NET assembly/in-proc drivers
Filename: "{dotnet4032}\regasm.exe"; Parameters: "/codebase ""{app}\ASCOM.Pulsar2.Telescope.dll"""; Flags: runhidden 32bit
Filename: "{dotnet4064}\regasm.exe"; Parameters: "/codebase ""{app}\ASCOM.Pulsar2.Telescope.dll"""; Flags: runhidden 64bit; Check: IsWin64




; Only if driver is .NET
[UninstallRun]
; Only for .NET assembly/in-proc drivers
Filename: "{dotnet4032}\regasm.exe"; Parameters: "-u ""{app}\ASCOM.Pulsar2.Telescope.dll"""; Flags: runhidden 32bit
; This helps to give a clean uninstall
Filename: "{dotnet4064}\regasm.exe"; Parameters: "/codebase ""{app}\ASCOM.Pulsar2.Telescope.dll"""; Flags: runhidden 64bit; Check: IsWin64
Filename: "{dotnet4064}\regasm.exe"; Parameters: "-u ""{app}\ASCOM.Pulsar2.Telescope.dll"""; Flags: runhidden 64bit; Check: IsWin64




[CODE]
//
// Before the installer UI appears, verify that the (prerequisite)
// ASCOM Platform 6.0 or greater is installed, including both Helper
// components. Utility is required for all types (COM and .NET)!
//
function InitializeSetup(): Boolean;
var
   U : Variant;
   H : Variant;
begin
   Result := FALSE;  // Assume failure
   // check that the DriverHelper and Utilities objects exist, report errors if they don't
   try
      H := CreateOLEObject('DriverHelper.Util');
   except
      MsgBox('The ASCOM DriverHelper object has failed to load, this indicates a serious problem with the ASCOM installation', mbInformation, MB_OK);
   end;
   try
      U := CreateOLEObject('ASCOM.Utilities.Util');
   except
      MsgBox('The ASCOM Utilities object has failed to load, this indicates that the ASCOM Platform has not been installed correctly', mbInformation, MB_OK);
   end;
   try
      if (U.IsMinimumRequiredVersion(6,0)) then	// this will work in all locales
         Result := TRUE;
   except
   end;
   if(not Result) then
      MsgBox('The ASCOM Platform 6.0 or greater is required for this driver.', mbInformation, MB_OK);
end;

// Code to enable the installer to uninstall previous versions of itself when a new version is installed
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  UninstallExe: String;
  UninstallRegistry: String;
begin
  if (CurStep = ssInstall) then // Install step has started
	begin
      // Create the correct registry location name, which is based on the AppId
      UninstallRegistry := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}' + '_is1');
      // Check whether an extry exists
      if RegQueryStringValue(HKLM, UninstallRegistry, 'UninstallString', UninstallExe) then
        begin // Entry exists and previous version is installed so run its uninstaller quietly after informing the user
          MsgBox('Setup will now remove the previous version.', mbInformation, MB_OK);
          Exec(RemoveQuotes(UninstallExe), ' /SILENT', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode);
          sleep(1000);    //Give enough time for the install screen to be repainted before continuing
        end
  end;
end;

