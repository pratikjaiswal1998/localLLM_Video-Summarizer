import sys
import os
from faster_whisper import WhisperModel

def main():
    # Force UTF-8 encoding for Windows so foreign languages (Hindi, etc) don't crash the print statements
    if sys.stdout.encoding.lower() != 'utf-8':
        sys.stdout.reconfigure(encoding='utf-8')
        
    if len(sys.argv) < 3:
        print("[!] Error: Not enough arguments provided to transcriber.", flush=True)
        sys.exit(1)
        
    audio_file = sys.argv[1]
    output_path = sys.argv[2]
    
    import shutil
    print("[*] Loading Faster-Whisper (large-v3-turbo on 8GB VRAM)...", flush=True)
    
    model = None
    for attempt in range(2):
        try:
            model = WhisperModel("large-v3-turbo", device="cuda", compute_type="float16")
            break # Success
        except Exception as e:
            error_str = str(e)
            if "model.bin" in error_str and "models--mobiuslabsgmbh--faster-whisper" in error_str and attempt == 0:
                print(f"[!] Detected corrupted HuggingFace cache: {e}", flush=True)
                print("[*] Attempting to auto-delete corrupted cache and redownload...", flush=True)
                
                cache_dir = os.path.expanduser("~/.cache/huggingface/hub/models--mobiuslabsgmbh--faster-whisper-large-v3-turbo")
                if os.path.exists(cache_dir):
                    try:
                        shutil.rmtree(cache_dir)
                        print("[*] Corrupted cache deleted successfully. Redownloading model (this may take a few minutes)...", flush=True)
                        continue # Try again
                    except Exception as rm_err:
                        print(f"[!] Failed to delete cache: {rm_err}. Please manually delete {cache_dir}", flush=True)
                        sys.exit(1)
                else:
                    print(f"[!] Cache directory not found at {cache_dir}. Cannot auto-repair.", flush=True)
                    sys.exit(1)
            else:
                print(f"[!] Error loading whisper model: {e}", flush=True)
                sys.exit(1)
                
    if model is None:
        print("[!] Failed to initialize Whisper model after retries.", flush=True)
        sys.exit(1)
        
    print("[*] Transcribing audio (this will be very fast)...", flush=True)
    segments, info = model.transcribe(audio_file, beam_size=5)
    
    print(f"[*] Detected language: {info.language} with probability {info.language_probability:.2f}", flush=True)
    
    text = ""
    timestamped_text = ""
    for segment in segments:
        seg_text = f"[{segment.start:.2f}s -> {segment.end:.2f}s] {segment.text}"
        print(seg_text, flush=True)
        text += segment.text + " "
        timestamped_text += seg_text + "\n"
        
    transcript = text.strip()
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(transcript)
        
    # Save a timestamped version for PowerPoint screenshot logic
    timestamped_path = output_path.replace(".txt", "_timestamped.txt")
    with open(timestamped_path, "w", encoding="utf-8") as f:
        f.write(timestamped_text)
        
    print(f"\n[*] Saved full transcript to {output_path}", flush=True)
    print(f"[*] Saved timestamped transcript to {timestamped_path}", flush=True)
    
    # Forcefully bypass Python's graceful shutdown and C++ destructors
    # CTranslate2 has a known bug on Windows where it Segfaults during garbage collection
    os._exit(0)

if __name__ == "__main__":
    main()
