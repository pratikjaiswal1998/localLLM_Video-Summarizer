
[Setup]
AppName=AI Video Summarizer
AppVersion=1.0
AppPublisher=Pratik
DefaultDirName={autopf}\AI Video Summarizer
DefaultGroupName=AI Video Summarizer
UninstallDisplayIcon={app}\summarizer_gui.exe
Compression=lzma2
SolidCompression=yes
OutputDir=C:\Users\Pratik\Documents\AntiAutomation2\AI_Summarizer\Output
OutputBaseFilename=AISummarizerSetup
SetupIconFile=C:\Users\Pratik\Documents\AntiAutomation2\AI_Summarizer\icon.ico
LicenseFile=C:\Users\Pratik\Documents\AntiAutomation2\AI_Summarizer\LICENSE.txt
ArchitecturesInstallIn64BitMode=x64

[Files]
; Main Python Executable and Env
Source: "C:\Users\Pratik\Documents\AntiAutomation2\AI_Summarizer\dist\summarizer_gui\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Transcription Backend
Source: "C:\Users\Pratik\Documents\AntiAutomation2\AI_Summarizer\dist\transcriber\transcriber_backend\*"; DestDir: "{app}\transcriber_backend_data"; Flags: ignoreversion recursesubdirs createallsubdirs
; Ollama Offline Setup
Source: "C:\Users\Pratik\Documents\AntiAutomation2\AI_Summarizer\OllamaSetup.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Icons]
Name: "{autoprograms}\AI Video Summarizer"; Filename: "{app}\summarizer_gui.exe"
Name: "{autodesktop}\AI Video Summarizer"; Filename: "{app}\summarizer_gui.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"

[Run]
; Silent Install Ollama (Skipped if already installed)
Filename: "{tmp}\OllamaSetup.exe"; Parameters: "/SILENT"; StatusMsg: "Installing Local AI Engine (Ollama)..."; Flags: waituntilterminated; Check: not IsOllamaInstalled
; Run main program
Filename: "{app}\summarizer_gui.exe"; Description: "Launch AI Video Summarizer"; Flags: nowait postinstall skipifsilent

[Code]
function IsOllamaInstalled(): Boolean;
var
  OllamaPath: String;
begin
  if RegQueryStringValue(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Ollama', 'InstallLocation', OllamaPath) then
  begin
    Result := True;
  end
  else if FileExists(ExpandConstant('{localappdata}\Programs\Ollama\ollama app.exe')) then
  begin
    Result := True;
  end
  else
  begin
    Result := False;
  end;
end;
