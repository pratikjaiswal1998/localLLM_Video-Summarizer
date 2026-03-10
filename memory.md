# Local Video Summarizer - Project Memory

**Project Path:** `C:\Users\Pratik\Documents\AntiAutomation2\AI_Summarizer`

This document serves as an operational memory state for cross-agent collaboration. It details the architecture, critical hardware constraints, and specific hacks required to make the application run smoothly on the user's hardware.

## Architecture & Files
- **`summarizer_gui.py`**: The main `tkinter`-based GUI. Handles downloading (via `yt-dlp`), UI threading, Ollama process management, transcript chunking, and final summarization.
- **`transcriber_backend.py`**: An isolated subprocess script that handles `faster-whisper` inference.
- **`Run_AI_Summarizer.bat`**: A simple launcher that executes the GUI using the standard `python` binary so the console stays open for debugging if it crashes.
- **`transcript.txt` & `summary.txt`**: Temporary/output files generated during the pipeline.

## Core Technologies
- **Transcription**: `faster-whisper` using the `large-v3-turbo` model optimized with `float16` compute type.
- **Summarization**: `ollama` running Meta's `llama3.1` (8B parameters).
- **Downloading**: `yt-dlp` for robust media extraction.

## Critical Hardware Constraints & Applied Fixes
The user is running on an **NVIDIA RTX 3070 with 8GB VRAM**. Because both Whisper (~3GB) and Llama 3.1 (~5GB + Context window) consume significant memory, running both in the same continuous Python session causes fatal **CUDA Out Of Memory (OOM) 500 Errors**.

The following hacks and architecture requirements have been hardcoded to bypass these limits:

### 1. The "C++ Memory Hoard" Fix (Subprocess Isolation)
`CTranslate2` (the C++ engine behind `faster-whisper`) pools VRAM and refuses to pass it back to Python's garbage collector. 
**The Fix:** Transcription was ripped out of the main GUI thread and moved to `transcriber_backend.py`. The GUI runs this as a complete `subprocess.Popen`. When the subprocess finishes, the Windows OS itself steps in to forcefully annihilate the process and reclaim 100% of the 3GB VRAM pool so Ollama has a clean slate.

### 2. CTranslate2 Windows Destructor Crash (0xC0000005)
On Windows, when `transcriber_backend.py` reaches the EOF and gracefully shuts down Python, `CTranslate2`'s C++ destructor often triggers an Access Violation Segfault. This causes the subprocess to return a crash code, making the GUI think transcription failed even though `transcript.txt` was fully written.
**The Fix:** `transcriber_backend.py` is hardcoded to instantly terminate using `os._exit(0)` the exact millisecond the file is saved to disk, completely bypassing Python's graceful teardown and the buggy C++ destructors.

### 3. Massive Context Window OOM (1-Hour Videos)
A 1-hour video contains ~12,000 words. When Llama 3.1 processes this, its KV Cache expands and demands ~3GB of extra VRAM, instantly crashing the 8GB card.
**The Fix:**
- The GUI explicitly caps the LLM's context size: `ollama.chat(..., options={'num_ctx': 8192})`.
- The GUI dynamically sections the massive text file into **overlapping 4,500-word chunks**, sending them to Ollama sequentially and appending the summaries (`Part 1 of 3...`). 

### 4. Ollama Background Daemon Conflicts
Running `subprocess.Popen("ollama serve")` on Windows often fails silently if the Ollama Tray App is asleep or installed in a specific user context. 
**The Fix:** The GUI directly targets the actual executable via `%LOCALAPPDATA%\Programs\Ollama\ollama app.exe` to forcefully wake the official system tray daemon, guaranteeing the API is reachable.

### 5. Auto-Downloading Llama 3.1
If the user hasn't downloaded the 4.7GB model yet, the Python API throws an explicit `ollama.ResponseError` (Status 404). 
**The Fix:** The GUI catches this 404 explicitly, initiates an `ollama.pull("llama3.1", stream=True)` download loop, parses the bytes safely (handling `NoneType` values the server occasionally sends during init), draws the UI progress bar, and automatically restarts the summary once the 4.7GB download hits 100%.

### 6. Aggressive VRAM Purge (keep_alive=0)
By default, Ollama holds the Llama 3.1 model in VRAM for 5 minutes after a prompt. Because the user might immediately summarize a new video and need that 8GB for Faster-Whisper, this cache is dangerous.
**The Fix:** The exact millisecond the summary finishes, the GUI sends an empty chat request with the `keep_alive=0` flag: `ollama.chat(model='llama3.1', messages=[{'role': 'user', 'content': ''}], keep_alive=0)`. This instantly purges the 5GB model from memory, returning the GPU to 0% usage.

### 7. Modernized UI (CustomTkinter)
Replaced standard `tkinter` with `customtkinter` for a 3D dark-mode aesthetic. 
**The Fix:** Implemented a **3-Stage Tracker** (`Extract` -> `Listen` -> `Write`) linked to the progress bar to provide visual feedback on which part of the VRAM-intensive pipeline is currently active.

### 8. Hindi & Unicode Support (Windows Encoding)
Windows defaults to `cp1252` for subprocess pipes, causing crashes (`UnicodeEncodeError`) when processing Hindi or other non-Latin transcripts.
**The Fix:** Forced `encoding="utf-8"` on `subprocess.Popen` in the GUI and added `sys.stdout.reconfigure(encoding='utf-8')` to the transcriber backend.

### 9. Automated PowerPoint Generation (v2 — Design Review Overhaul, March 2026)
Integration of `python-pptx` to build slide decks with automated visuals.
**The Logic (v2.3):** Prompt dynamically calculates `min_slides` from chunk word count (~1 slide per 600 words, min 3). Includes a 4-slide JSON example so Llama 3.1 copies the structure. States "YOU MUST PRODUCE AT LEAST N SLIDES" explicitly. `format='json'` removed — was forcing flat string arrays. If LLM produces fewer slides than `min_slides`, Python auto-supplements from the uncovered portion of the transcript.

**v2 Design Fixes Applied:**
- **Title Slide + Closing Slide**: Auto-generated programmatically (not LLM). Title slide shows video topic centered. Closing slide shows "KEY TAKEAWAYS" with first bullet from each content slide + a verification disclaimer.
- **Entity Auto-Correction**: `ENTITY_CORRECTIONS` dict fixes common transcription errors (Anthropik→Anthropic, Amadei→Amodei) in both titles and bullets.
- **Dynamic Title Font Sizing**: ≤45 chars = 40pt, ≤60 chars = 34pt, >60 chars = 28pt. Word wrap enabled (`tf_title.word_wrap = True`).
- **Colored Bullet Markers**: Blue `\u2022` character as a separate run before body text (white). No more invisible flat paragraphs.
- **Empty First Paragraph Fix**: First bullet uses `tf.paragraphs[0]` instead of `tf.add_paragraph()`.
- **Image Aspect Ratio Clamping**: Uses PIL to check aspect ratio; clamps height if image would exceed 4.8" vertical space.
- **Image Border**: 1pt gray (#505050) outline around screenshots to separate from dark background.
- **Timestamp Captions**: Italic muted-gray caption below each image showing "Timestamp: M:SS".
- **Full-Width Fallback Layout**: If screenshot fails, bullets span full 11.7" width instead of cramped right column.
- **Dynamic Bullet Font + Spacing**: ≤3 = 22pt/14pt, ≤5 = 18pt/10pt, ≤7 = 16pt/8pt, >7 = 14pt/6pt. Both font size AND paragraph spacing scale together.
- **Dynamic Closing Slide Recap**: Font scales by recap count: ≤5 = 20pt, ≤8 = 16pt, ≤12 = 13pt, >12 = 11pt. Box position and height also adjust. Wider recap box (11.3") for long-form summaries.
- **Slide Numbers**: Bottom-right corner, "N / Total" format, 11pt muted gray.
- **Title Color Hierarchy**: Titles are pure white (#FFFFFF), body text is #DCDCDC, captions/numbers are #828282.
- **Thicker Accent Bar**: 0.25" instead of 0.10" — visible when projected.
- **VRAM Purge After PPT**: `keep_alive=0` now fires after PPT generation path too (was missing before).

### 10. Robust JSON Extraction (LLM Conversational Filler)
Llama 3.1 8B occasionally ignores "JSON only" instructions and adds conversational text or malformed markdown blocks.
**The Fix:** Implemented a **Multi-Stage Extractor**:
1. Try Regex for ` ```json ` blocks.
2. Fallback: Locate the first `[` and last `]` characters to surgically slice the JSON array out of the raw string.
3. **Duplicate-Key Dict Fix (v2)**: If LLM returns `{"slide":{...}, "slide":{...}}` instead of an array, regex extracts all `{...}` objects and wraps them in `[...]`. Also searches for `"slide"` key in addition to `"slides"/"data"/"presentation"`.
4. **Removed `format='json'` (v2.3)**: Ollama's `format='json'` was forcing Llama 3.1 8B into the simplest valid JSON (flat string arrays). Removing it lets the model output naturally; JSON is then extracted from free-form text via `_extract_json_from_text()`.
5. **Python Transcript Structurer (v2.3)**: Pure Python fallback `_python_structure_transcript()` that requires zero LLM cooperation. Parses `[Xs -> Ys] text` lines from the timestamped transcript, groups every 5 lines into a slide, auto-generates titles from content words, and preserves real timestamps. Activates when LLM output is completely unusable.
6. **Flexible key detection (v2.3)**: Accepts `bullets`/`bullet_points`/`points` for bullet arrays, and `title`/`heading`/`header` for titles. Salvages dicts-without-bullets by extracting long string values.
7. Error logging includes exception type and raw content snippet.

### 12. Defensive PPT Pipeline (Fail-Safe Builder)
LLMs occasionally deviate from JSON schemas despite hard constraints.
**The Fix: Triple-Layer Architecture:**
1. **Engine Level:** `format='json'` used in Ollama API.
2. **Aggregation Level:** `isinstance(item, dict)` filtering added to the slide collector to prevent strings from entering the build queue.
3. **Builder Level:** Added type-normalization in `create_ppt` to auto-convert single strings into bullet lists if the AI forgets the array format.

### 13. Bare Exception Cleanup (v2)
- `clean_timestamp()`: Changed bare `except` to `except (ValueError, TypeError)`.
- JSON parsing: Changed bare `except Exception` to log actual error type + message + raw snippet.
