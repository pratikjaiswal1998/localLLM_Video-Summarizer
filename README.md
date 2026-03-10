# AI Video Summarizer 🎓

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Python](https://img.shields.io/badge/Python-3.12-blue.svg)
![Offline Status](https://img.shields.io/badge/Status-100%25_Offline-brightgreen)

Welcome to your own private content engine! **AI Video Summarizer** is a completely free, 100% offline desktop application designed to ingest massive YouTube videos and distill them into visually stunning, actionable PowerPoint presentations.

---

## 🚀 Core Features

- **100% Local & Private:** No cloud subscriptions. No internet required for processing. Complete privacy.
- **Blistering Fast Transcription:** Powered by hardware-accelerated `Faster-Whisper` for incredibly fast speech-to-text.
- **Smart Generation & Structuring:** Uses [Ollama](https://ollama.com) (Llama 3 local LLM) to create intelligent, concise bullet points from raw audio transcripts.
- **One-Click Presentations:** Automatically formats and outputs a premium 16:9 `.pptx` PowerPoint presentation perfectly spaced and centered.
- **Native Windows Installer:** Comes bundled as a sleek `.exe` installer that automatically detects and manages all heavy AI dependencies (like `FFmpeg` and `Ollama`) without the hassle of virtual environments.

## 🛠️ Installation

### Option 1: The Smart Windows Installer (Recommended)
1. Navigate to the `Output` directory and run `AISummarizerSetup.exe`.
2. The smart installer will check your PC for existing dependencies:
   - If you don't have **Ollama** installed, it will install it for you silently.
   - It seamlessly bundles **FFmpeg**, `yt-dlp`, and the **PyTorch** engine.
3. Launch the app from your Desktop or Start Menu. Note: On the very first run, it will natively cache the 1.5GB Whisper model directly to your system (so you never have to download it again).

### Option 2: Running from Source
If you are a developer and wish to run the app directly from code:
```bash
git clone https://github.com/pratikjaiswal1998/localLLM_Video-Summarizer.git
cd localLLM_Video-Summarizer

# Create a virtual environment
python -m venv .venv
source .venv/Scripts/activate

# Install dependencies
pip install -r requirements.txt

# Run the UI
python summarizer_gui.py
```
*(Ensure `ffmpeg.exe` is in your PATH and Ollama is running (`ollama serve`)).*

## 📚 Marketing Material
Check out our gorgeous PDF brochure in the `assets/marketing/` folder for a quick overview of the application's capabilities, perfect for sharing with colleagues!

## 🤝 Contributing
Open source thrives on community help! We welcome any pull requests, issue reports, or feature requests. 
Please read through the [CONTRIBUTING.md](CONTRIBUTING.md) and [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md) before pushing code.

## 📜 License
This project is licensed under the MIT License - see the `LICENSE.txt` file for details.
