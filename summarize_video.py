import os
import sys
import subprocess
import tempfile
import yt_dlp
from faster_whisper import WhisperModel
import ollama

def download_audio(url, output_dir):
    print(f"[*] Downloading audio from {url}...")
    
    ydl_opts = {
        'format': 'bestaudio/best',
        'outtmpl': os.path.join(output_dir, '%(title)s.%(ext)s'),
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '128',
        }],
        'quiet': False,
        'no_warnings': True
    }
    
    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(url, download=True)
        filename = ydl.prepare_filename(info)
        name, _ = os.path.splitext(filename)
        return name + ".mp3"

def transcribe_audio(audio_path):
    print(f"\n[*] Loading Faster-Whisper (large-v3-turbo)...")
    # compute_type="float16" fits nicely in 8GB VRAM
    model = WhisperModel("large-v3-turbo", device="cuda", compute_type="float16")
    
    print(f"[*] Transcribing audio (this will be very fast)...")
    segments, info = model.transcribe(audio_path, beam_size=5)
    
    print(f"[*] Detected language: {info.language} with probability {info.language_probability:.2f}")
    
    text = ""
    for segment in segments:
        print(f"[{segment.start:.2f}s -> {segment.end:.2f}s] {segment.text}")
        text += segment.text + " "
        
    return text.strip()

def summarize_text(text):
    print(f"\n[*] Sending transcript to Ollama (Llama 3.1 8B) for summarization...")
    print(f"[*] Make sure Ollama is installed and you ran: 'ollama run llama3.1'")
    
    prompt = f"""
    Please provide a detailed, well-structured summary of the key points discussed in this transcript. 
    Include bullet points for the main takeaways.
    
    TRANSCRIPT:
    {text}
    """
    
    try:
        response = ollama.chat(model='llama3.1', messages=[
            {
                'role': 'user',
                'content': prompt
            }
        ])
        return response['message']['content']
    except Exception as e:
        return f"[!] Error connecting to Ollama: {e}\nIs Ollama running?"

def main():
    print("="*50)
    print("Local AI Video Summarizer (Whisper + Llama 3.1)")
    print("="*50)
    
    if len(sys.argv) < 2:
        url = input("Enter video URL or local file path: ")
    else:
        url = sys.argv[1]
        
    if not url.strip():
        print("No input provided.")
        return

    # Use current directory to save final outputs
    current_dir = os.getcwd()

    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            # Step 1: Download
            if url.startswith("http"):
                audio_file = download_audio(url, temp_dir)
            else:
                audio_file = url # Treat as local file path
                
            if not os.path.exists(audio_file):
                print(f"[!] Error: Audio file {audio_file} not found.")
                return
                
            # Step 2: Transcribe
            transcript = transcribe_audio(audio_file)
            
            # Save transcript
            transcript_path = os.path.join(current_dir, "transcript.txt")
            with open(transcript_path, "w", encoding="utf-8") as f:
                f.write(transcript)
            print(f"\n[*] Saved full transcript to {transcript_path}")
            
            # Step 3: Summarize
            summary = summarize_text(transcript)
            
            # Save summary
            summary_path = os.path.join(current_dir, "summary.txt")
            with open(summary_path, "w", encoding="utf-8") as f:
                f.write(summary)
            
            print("\n" + "="*50)
            print("SUMMARY OF VIDEO:")
            print("="*50)
            print(summary)
            print("="*50)
            print(f"\n[*] Saved summary to {summary_path}")
            
        except Exception as e:
            print(f"\n[!] An error occurred: {e}")

if __name__ == "__main__":
    main()
