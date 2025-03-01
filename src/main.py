from fastapi import FastAPI, BackgroundTasks
from pydantic import BaseModel
from typing import Optional
import os
from dotenv import load_dotenv
from pptx import Presentation
from pptx.dml.color import RGBColor  # Import for color styling
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import google.generativeai as genai  # Correct import
from pptx.util import Pt  # Import for font size control
from pptx.enum.text import PP_ALIGN  # Import for text alignment
from pptx.dml.color import RGBColor  # Import for color styling
from pptx.util import Inches

# Load environment variables
load_dotenv()

# Get API key for Google Gemini
api_key = os.getenv("GOOGLE_API_KEY")
if not api_key:
    raise ValueError("GOOGLE_API_KEY is not set in the environment or .env file!")

# Configure Google Gemini API
genai.configure(api_key=api_key)

app = FastAPI()

# Enable CORS for frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "https://intelli-slide-ai.vercel.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# Request model with language support
class PPTRequest(BaseModel):
    topic: str
    num_slides: Optional[int] = 5
    language: Optional[str] = "English"  # Default language is English

# Generate slide content using Google Gemini
def generate_slide_content(topic: str, num_slides: int, language: str):
    prompt = (
        f"Generate {num_slides} slides in {language} about {topic}. "
        f"Each slide should include:\n"
        f"- A title\n"
        f"- Bullet points with short descriptions\n"
        f"- A short content or wider explanation for each bullet\n"
        f"- A placeholder for images if relevant (indicate it with 'Insert Image: [description]')\n"
        f"Format each slide as: 'Title\\n- Bullet 1 (short explanation)\\n- Bullet 2'."
    )

    # Use the Gemini model correctly
    model = genai.GenerativeModel("gemini-2.0-flash")
    response = model.generate_content(prompt)

    if response and hasattr(response, "text") and response.text:
        slides = response.text.strip().split("\n\n")  # Split into slides
        return [slide.split("\n") for slide in slides]  # Split title & bullets
    return []

# Create PowerPoint file with Font Styling and Proper Fitting
def create_pptx(topic: str, slides_data, filename="presentation.pptx"):
    prs = Presentation()

    # Title Slide
    slide_layout = prs.slide_layouts[0]  
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = topic
    subtitle.text = "AI-Generated Presentation"

    # Apply styling to title slide
    title_text_frame = title.text_frame
    title_text_frame.paragraphs[0].font.size = Pt(44)  # Larger Font
    title_text_frame.paragraphs[0].font.color.rgb = RGBColor(34, 139, 34)  # Green
    title_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER  # Center alignment

    # Content Slides
    for slide_content in slides_data:
        if len(slide_content) < 2:
            continue
        title_text = slide_content[0]
        bullet_points = slide_content[1:]

        slide_layout = prs.slide_layouts[5]  # Use a layout with more space
        slide = prs.slides.add_slide(slide_layout)
        
        # Manually position title to leave more space for content
        title = slide.shapes.title
        title.text = title_text
        title.left = Inches(0.5)
        title.top = Inches(0.3)  # Move title slightly up
        title.width = Inches(9)
        title.height = Inches(1)

        # Apply styling to slide title
        title_text_frame = title.text_frame
        title_text_frame.paragraphs[0].font.size = Pt(32)  # Slightly smaller than before
        title_text_frame.paragraphs[0].font.color.rgb = RGBColor(34, 139, 34)  # Green
        title_text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER  # Center alignment

        # Add a textbox manually for content with extra spacing
        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(9), Inches(5))  # Lower position
        content_text_frame = content_box.text_frame
        content_text_frame.word_wrap = True

        # Apply styling to content and ensure it fits properly
        for bullet in bullet_points:
            p = content_text_frame.add_paragraph()
            p.text = bullet.strip("- ")  # Clean up bullet points
            p.font.size = Pt(24)  # Readable size
            p.font.color.rgb = RGBColor(128, 128, 128)  # Gray
            p.space_after = Pt(8)  # Add spacing between bullet points
            p.alignment = PP_ALIGN.LEFT  # Align left

    prs.save(filename)
    return filename

@app.get("/")
async def welcome():
    return {"message": "Welcome to IntelliSlide-AI! Use /generate_ppt to create a PowerPoint presentation."}

@app.post("/generate_ppt")
async def generate_ppt(request: PPTRequest, background_tasks: BackgroundTasks):
    slides_data = generate_slide_content(request.topic, request.num_slides, request.language)
    filename = f"{request.topic.replace(' ', '_')}.pptx"
    background_tasks.add_task(create_pptx, request.topic, slides_data, filename)
    return {"message": "Presentation is being generated.", "filename": filename}

@app.get("/download_ppt/{filename}")
async def download_ppt(filename: str):
    file_path = f"./{filename}"
    if os.path.exists(file_path):
        return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename=filename)
    return {"error": "File not found"}
