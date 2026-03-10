import os
from fpdf import FPDF

# Configure output path
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "marketing")
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "AI_Video_Summarizer_Brochure.pdf")

# Icon image for the PDF (using the funny professor)
ICON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
if not os.path.exists(ICON_PATH):
    ICON_PATH = None

class MarketingPDF(FPDF):
    def header(self):
        # Header banner
        self.set_fill_color(30, 30, 30)
        self.rect(0, 0, 210, 40, 'F')
        
        # Title
        self.set_font('helvetica', 'B', 24)
        self.set_text_color(255, 255, 255)
        self.set_y(15)
        self.cell(0, 10, 'AI Video Summarizer', border=False, align='C', new_x="LMARGIN", new_y="NEXT")
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, 'Open Source | 100% Local | Privacy First', align='C', new_x="LMARGIN", new_y="NEXT")

def create_brochure():
    pdf = MarketingPDF()
    pdf.add_page()
    
    # Hero Section
    pdf.set_font('helvetica', 'B', 20)
    pdf.set_text_color(0, 102, 204)
    pdf.ln(10)
    pdf.cell(0, 10, 'Your Own Private Content Engine', align='C', new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    # Body Text
    pdf.set_font('helvetica', '', 12)
    pdf.set_text_color(60, 60, 60)
    intro_text = (
        "Welcome to the future of content digestion. The AI Video Summarizer is a completely free, "
        "100% offline engine designed to ingest massive YouTube videos and distill them into visually "
        "stunning, actionable PowerPoints.\n\n"
        "No cloud subscriptions. No internet required for processing. Complete privacy."
    )
    pdf.multi_cell(0, 8, intro_text, align='C')
    pdf.ln(15)
    
    # Features Section
    pdf.set_font('helvetica', 'B', 16)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(0, 10, 'Core Features', new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    
    features = [
        ("Zero Cloud Dependencies", "Don't pay $20/month for cloud AI. Run everything entirely on your own hardware."),
        ("Blistering Fast Transcription", "Powered by Faster-Whisper, experience hardware-accelerated speech-to-text."),
        ("Smart Generation", "Ollama drives local LLMs (like Llama 3) to create intelligent, concise bullet points."),
        ("One-Click Presentations", "Automatically formats and generates a 16:9 premium PowerPoint presentation."),
        ("Native Windows Experience", "Bundled as a sleek .exe installer that cleanly manages all complex AI dependencies like FFmpeg.")
    ]
    
    for title, desc in features:
        pdf.set_font('helvetica', 'B', 12)
        pdf.set_text_color(0, 102, 204)
        pdf.cell(0, 8, f"- {title}", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(6)
        
        pdf.set_font('helvetica', '', 11)
        pdf.set_text_color(80, 80, 80)
        pdf.set_x(15)
        pdf.multi_cell(0, 6, desc)
        pdf.ln(4)
        
    pdf.ln(10)
    
    # Call to action
    pdf.set_fill_color(240, 240, 240)
    pdf.rect(10, pdf.get_y(), 190, 30, 'F')
    pdf.set_y(pdf.get_y() + 5)
    pdf.set_font('helvetica', 'B', 14)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(0, 10, 'Ready to Reclaim Your Time?', align='C', new_x="LMARGIN", new_y="NEXT")
    pdf.set_font('helvetica', '', 11)
    pdf.multi_cell(0, 6, "Download the Smart Installer today. It automatically detects and skips existing dependencies, keeping your PC clean and your workflow fast.", align='C')
    
    pdf.output(OUTPUT_FILE)
    print(f"[*] Generated Marketing Brochure: {OUTPUT_FILE}")

if __name__ == "__main__":
    create_brochure()
