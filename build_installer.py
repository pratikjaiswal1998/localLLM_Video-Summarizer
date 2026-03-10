import os
import sys
import urllib.request
import zipfile
import subprocess
import shutil

DIR = os.path.dirname(os.path.abspath(__file__))

# 1. Download FFmpeg
FFMPEG_URL = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
ZIP_PATH = os.path.join(DIR, "ffmpeg.zip")
FFMPEG_EXE = os.path.join(DIR, "ffmpeg.exe")
FFPROBE_EXE = os.path.join(DIR, "ffprobe.exe")

if not os.path.exists(FFMPEG_EXE) or not os.path.exists(FFPROBE_EXE):
    print("[*] Downloading FFmpeg statically...")
    urllib.request.urlretrieve(FFMPEG_URL, ZIP_PATH)
    print("[*] Extracting FFmpeg...")
    with zipfile.ZipFile(ZIP_PATH, 'r') as zf:
        # The zip structure is typically: ffmpeg-master-latest-win64-gpl/bin/ffmpeg.exe
        extracted = 0
        for member in zf.namelist():
            if member.endswith("ffmpeg.exe"):
                # Extract specifically ffmpeg.exe
                with zf.open(member) as source:
                    with open(FFMPEG_EXE, "wb") as f:
                        shutil.copyfileobj(source, f)
                extracted += 1
            elif member.endswith("ffprobe.exe"):
                # Extract specifically ffprobe.exe
                with zf.open(member) as source:
                    with open(FFPROBE_EXE, "wb") as f:
                        shutil.copyfileobj(source, f)
                extracted += 1
            if extracted >= 2:
                break
    try:
        os.remove(ZIP_PATH)
    except Exception as e:
        print(f"[*] Warning: Could not remove zip file: {e}")

# 2. PyInstaller Build
print("[*] Compiling Python code via PyInstaller...")
# We use one-dir for faster extraction and bundle both the GUI and backend
# The transcriber_backend.py is separate but we can just use the exact same Pyinstaller folder
gui_exe = os.path.join(DIR, "dist", "summarizer_gui", "summarizer_gui.exe")

# Force rebuild to pick up new ffprobe
if os.path.exists(gui_exe):
    os.remove(gui_exe)

if not os.path.exists(gui_exe):
    gui_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--onedir", "--windowed",
        "--add-data", f"{FFMPEG_EXE};.",
        "--add-data", f"{FFPROBE_EXE};.",
        "summarizer_gui.py"
    ]
    subprocess.run(gui_cmd, check=True)

    backend_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--onedir", "--console",
        "--distpath", os.path.join(DIR, "dist", "transcriber"),
        "transcriber_backend.py"
    ]
    subprocess.run(backend_cmd, check=True)
else:
    print("[*] PyInstaller build exists, skipping...")

# 3. Create Inno Setup Script
print("[*] Generating Inno Setup Script...")
iss_content = f"""
[Setup]
AppName=AI Video Summarizer
AppVersion=1.0
AppPublisher=Pratik
DefaultDirName={{autopf}}\\AI Video Summarizer
DefaultGroupName=AI Video Summarizer
UninstallDisplayIcon={{app}}\\summarizer_gui.exe
Compression=lzma2
SolidCompression=yes
OutputDir={DIR}\\Output
OutputBaseFilename=AISummarizerSetup
SetupIconFile={DIR}\\icon.ico
LicenseFile={DIR}\\LICENSE.txt
ArchitecturesInstallIn64BitMode=x64

[Files]
; Main Python Executable and Env
Source: "{DIR}\\dist\\summarizer_gui\\*"; DestDir: "{{app}}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Transcription Backend
Source: "{DIR}\\dist\\transcriber\\transcriber_backend\\*"; DestDir: "{{app}}\\transcriber_backend_data"; Flags: ignoreversion recursesubdirs createallsubdirs
; Ollama Offline Setup
Source: "{DIR}\\OllamaSetup.exe"; DestDir: "{{tmp}}"; Flags: deleteafterinstall

[Icons]
Name: "{{autoprograms}}\\AI Video Summarizer"; Filename: "{{app}}\\summarizer_gui.exe"
Name: "{{autodesktop}}\\AI Video Summarizer"; Filename: "{{app}}\\summarizer_gui.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"

[Run]
; Silent Install Ollama (Skipped if already installed)
Filename: "{{tmp}}\\OllamaSetup.exe"; Parameters: "/SILENT"; StatusMsg: "Installing Local AI Engine (Ollama)..."; Flags: waituntilterminated; Check: not IsOllamaInstalled
; Run main program
Filename: "{{app}}\\summarizer_gui.exe"; Description: "Launch AI Video Summarizer"; Flags: nowait postinstall skipifsilent

[Code]
function IsOllamaInstalled(): Boolean;
var
  OllamaPath: String;
begin
  if RegQueryStringValue(HKEY_CURRENT_USER, 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\Ollama', 'InstallLocation', OllamaPath) then
  begin
    Result := True;
  end
  else if FileExists(ExpandConstant('{{localappdata}}\\Programs\\Ollama\\ollama app.exe')) then
  begin
    Result := True;
  end
  else
  begin
    Result := False;
  end;
end;
"""
iss_path = os.path.join(DIR, "setup.iss")
with open(iss_path, "w", encoding="utf-8") as f:
    f.write(iss_content)

print("[*] Compiling Installer (Inno Setup)...")
inno_path = r"C:\\Users\\Pratik\\AppData\\Local\\Programs\\Inno Setup 6\\ISCC.exe"
if not os.path.exists(inno_path):
    print(f"Inno Setup not found at {{inno_path}}!")
    sys.exit(1)

subprocess.run([inno_path, iss_path], check=True)
print("[*] SUCCESSFULLY BUILT INSTALLER!")
