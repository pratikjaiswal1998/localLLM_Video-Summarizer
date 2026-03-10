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
    
    print("[*] Loading Faster-Whisper (large-v3-turbo on 8GB VRAM)...", flush=True)
    try:
        model = WhisperModel("large-v3-turbo", device="cuda", compute_type="float16")
    except Exception as e:
        print(f"[!] Error loading whisper model: {e}", flush=True)
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
