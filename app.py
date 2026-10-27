import io, json, os, re, subprocess, tempfile, time
from pathlib import Path

import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

APP_DIR = Path(__file__).parent
OUTPUT_DIR = APP_DIR / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

st.set_page_config(page_title="MotionDeck AI", page_icon="🎬", layout="wide")

# ---------- Styling ----------
st.markdown("""
<style>
.stApp { background: linear-gradient(135deg,#0b1020 0%,#111a31 55%,#17102a 100%); color:#f4f6ff; }
.block-container { padding-top: 1.7rem; max-width: 1400px; }
.hero { padding: 1.6rem 1.8rem; border:1px solid #303a5d; border-radius:24px;
        background:linear-gradient(120deg,rgba(42,63,115,.55),rgba(91,42,122,.38)); margin-bottom:1rem; }
.hero h1 { font-size:2.5rem; margin:0; }
.hero p { color:#c6cee8; font-size:1.05rem; margin:.5rem 0 0; }
div[data-testid="stMetric"] { background:#161f38; border:1px solid #303a5d; padding:15px; border-radius:16px; }
.stButton>button, .stDownloadButton>button { border-radius:12px; min-height:2.8rem; font-weight:650; }
.small-note { color:#aeb9da; font-size:.88rem; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <h1>🎬 MotionDeck AI</h1>
  <p>Turn notes, articles, and ideas into editable PowerPoint decks and narrated animated videos.</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("## ⚙️ Project settings")
    model_mode = st.selectbox("Content engine", ["Built-in (works offline)", "Ollama local LLM"])
    ollama_model = st.text_input("Ollama model name", "llama3.2", disabled=model_mode != "Ollama local LLM")
    slide_count = st.slider("Number of slides", 3, 12, 6)
    theme = st.selectbox("Presentation theme", ["Midnight", "Ocean", "Sunset", "Minimal"])
    video_style = st.selectbox("Animated video style", [
        "Motion graphics", "Animated presenter", "Cinematic storyboard", "All styles"
    ])
    fps = st.select_slider("Video frame rate", options=[6, 8, 10, 12], value=8)
    seconds_per_slide = st.slider("Seconds per slide", 3, 8, 4)
    make_voice = st.checkbox("Generate local voice-over (pyttsx3)", value=True)
    st.caption("Voice quality depends on voices installed on your computer.")

theme_colors = {
    "Midnight": ("10172A", "7C5CFC", "F4F6FF", "AEB9DA"),
    "Ocean": ("082F49", "22D3EE", "F0FDFF", "A5F3FC"),
    "Sunset": ("351515", "FB7185", "FFF7ED", "FED7AA"),
    "Minimal": ("F4F6FA", "2563EB", "111827", "4B5563"),
}
bg_hex, accent_hex, fg_hex, muted_hex = theme_colors[theme]

def rgb(hexstr):
    return tuple(int(hexstr[i:i+2], 16) for i in (0,2,4))

def extract_text(uploaded):
    if not uploaded:
        return ""
    name = uploaded.name.lower()
    raw = uploaded.getvalue()
    if name.endswith(".txt") or name.endswith(".md"):
        return raw.decode("utf-8", errors="ignore")
    if name.endswith(".pdf"):
        try:
            from pypdf import PdfReader
            reader = PdfReader(io.BytesIO(raw))
            return "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception as e:
            st.error(f"Could not read PDF: {e}. Install pypdf or paste text instead.")
            return ""
    return ""

def call_ollama(text, n, model):
    prompt = f"""Create a teaching presentation from the source below.
Return ONLY valid JSON with this schema:
{{"title":"short title","slides":[{{"title":"slide title","bullets":["bullet 1","bullet 2","bullet 3"],"visual":"visual idea","narration":"2-4 spoken sentences"}}]}}
Create exactly {n} slides. Keep bullets concise and faithful to the source. Do not add unsupported facts.
SOURCE:
{text[:14000]}"""
    result = subprocess.run(
        ["ollama", "run", model], input=prompt, text=True,
        capture_output=True, timeout=180
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr or "Ollama failed")
    out = result.stdout.strip()
    # Tolerate code fences from models
    out = re.sub(r"^```(?:json)?\s*|\s*```$", "", out, flags=re.I)
    data = json.loads(out)
    slides = data.get("slides", [])
    if not slides:
        raise ValueError("Ollama returned no slides")
    return {"title": data.get("title", "AI-generated presentation"), "slides": slides[:n]}

def built_in_plan(text, n):
    clean = re.sub(r'\s+', ' ', text).strip()
    if not clean:
        clean = "Artificial intelligence uses data and algorithms to identify patterns, make predictions, and support decisions."
    # Sentence split, with a simple fallback
    sentences = re.split(r'(?<=[.!?])\s+', clean)
    sentences = [s.strip() for s in sentences if len(s.strip()) > 20]
    if not sentences:
        sentences = [clean[i:i+180] for i in range(0, min(len(clean), 1800), 180)]
    if not sentences:
        sentences = [clean]
    chunks = [[] for _ in range(n)]
    for i, sentence in enumerate(sentences):
        chunks[i % n].append(sentence)
    title = (sentences[0][:70].strip(" .,:;") if sentences else "My presentation").title()
    slides = []
    for i in range(n):
        items = chunks[i] or [sentences[i % len(sentences)]]
        bullets = []
        for item in items[:3]:
            item = item.strip()
            if len(item) > 135:
                item = item[:132].rsplit(" ", 1)[0] + "…"
            bullets.append(item)
        if not bullets:
            bullets = ["Key idea", "Explain the concept", "Connect it to a practical example"]
        slide_title = " ".join(bullets[0].split()[:7]).strip(" .,:;").title()
        narration = " ".join(items[:2])
        if len(narration) > 420:
            narration = narration[:417].rsplit(" ", 1)[0] + "..."
        slides.append({
            "title": slide_title or f"Key idea {i+1}",
            "bullets": bullets,
            "visual": "Use a clear diagram or visual metaphor for this idea",
            "narration": narration or "This slide introduces an important idea from the source."
        })
    return {"title": title or "My presentation", "slides": slides}

def build_plan(text, n, mode, model):
    if mode == "Ollama local LLM":
        return call_ollama(text, n, model)
    return built_in_plan(text, n)

def font(size, bold=False):
    candidates = [
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf"
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()

def wrap_text(draw, text, fnt, max_width):
    words = str(text).split()
    lines, line = [], ""
    for word in words:
        test = f"{line} {word}".strip()
        if draw.textbbox((0,0), test, font=fnt)[2] <= max_width:
            line = test
        else:
            if line: lines.append(line)
            line = word
    if line: lines.append(line)
    return lines

def render_frame(plan, slide, idx, total, style, theme_name, progress, width=1280, height=720):
    bg, accent, fg, muted = [rgb(x) for x in theme_colors[theme_name]]
    im = Image.new("RGB", (width, height), bg)
    d = ImageDraw.Draw(im)
    # subtle top band and accent
    d.rounded_rectangle((48, 38, width-48, height-38), radius=28, outline=tuple(min(255,int(c*1.2)) for c in bg), width=2)
    d.rounded_rectangle((70, 72, 84, height-72), radius=7, fill=accent)
    d.text((112, 76), f"MOTIONDECK  /  {idx+1:02d}", font=font(20, True), fill=accent)
    title = slide.get("title", f"Slide {idx+1}")
    title_lines = wrap_text(d, title, font(44, True), 980)
    y = 128
    for line in title_lines[:2]:
        d.text((112, y), line, font=font(44, True), fill=fg)
        y += 55
    # body bullets
    bullets = slide.get("bullets", [])[:4]
    y += 20
    for j, bullet in enumerate(bullets):
        d.ellipse((120, y+8, 134, y+22), fill=accent)
        lines = wrap_text(d, bullet, font(25), 930)
        for line in lines[:2]:
            d.text((156, y), line, font=font(25), fill=fg)
            y += 34
        y += 15
    # Style-specific visual
    if style == "Animated presenter":
        # A friendly illustrated presenter with simple animated mouth
        cx, cy = 1090, 485
        d.ellipse((cx-72, cy-100, cx+72, cy+44), fill=(245, 197, 153), outline=accent, width=5)
        d.pieslice((cx-76, cy-110, cx+76, cy+50), 180, 360, fill=(54, 40, 48))
        d.ellipse((cx-42, cy-38, cx-25, cy-21), fill=(30, 30, 40))
        d.ellipse((cx+25, cy-38, cx+42, cy-21), fill=(30, 30, 40))
        mouth_h = 7 + int(12 * abs(__import__('math').sin(progress * 18)))
        d.ellipse((cx-22, cy+2, cx+22, cy+2+mouth_h), fill=(120, 35, 55))
        d.rounded_rectangle((cx-96, cy+42, cx+96, cy+225), radius=42, fill=accent)
        d.text((cx-95, cy+245), "AI PRESENTER", font=font(16, True), fill=muted)
    elif style == "Cinematic storyboard":
        # cinematic scene card, layered hills, sun, and a camera-like frame
        d.rounded_rectangle((850, 310, 1150, 570), radius=22, outline=accent, width=4)
        d.ellipse((1000, 340, 1060, 400), fill=(255, 203, 107))
        d.polygon([(870,540),(960,420),(1040,535)], fill=accent)
        d.polygon([(950,540),(1070,400),(1130,540)], fill=tuple(min(255,c+35) for c in accent))
        d.rectangle((870, 540, 1130, 555), fill=tuple(max(0,c-20) for c in accent))
        d.text((875, 580), "SCENE PREVIEW", font=font(16, True), fill=muted)
    else:
        # animated infographic bars
        base_y = 555
        bar_x = 880
        for j, h in enumerate([75, 120, 165, 205]):
            animated_h = int(h * (0.65 + 0.35 * progress))
            x = bar_x + j*58
            d.rounded_rectangle((x, base_y-animated_h, x+34, base_y), radius=8, fill=accent if j == 3 else tuple(max(0,int(c*(0.55+j*.12))) for c in accent))
        d.line((865, base_y+2, 1145, base_y+2), fill=muted, width=2)
        d.text((875, 580), "KEY CONCEPT", font=font(16, True), fill=muted)
    # progress and footer
    d.rounded_rectangle((112, 650, width-112, 660), radius=5, fill=tuple(max(0,int(c*.35)) for c in accent))
    d.rounded_rectangle((112, 650, 112+int((width-224)*(idx+progress)/total), 660), radius=5, fill=accent)
    d.text((112, 674), plan.get("title", "AI Presentation")[:80], font=font(14), fill=muted)
    d.text((width-180, 674), f"{idx+1} / {total}", font=font(14, True), fill=fg)
    return im

def create_pptx(plan, out_path, theme_name):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    bg, accent, fg, muted = [rgb(x) for x in theme_colors[theme_name]]
    for i, item in enumerate(plan["slides"]):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        # background
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*bg)
        # accent stripe
        shape = slide.shapes.add_shape(1, Inches(.35), Inches(.45), Inches(.12), Inches(6.35))
        shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(*accent); shape.line.fill.background()
        title_box = slide.shapes.add_textbox(Inches(.8), Inches(.55), Inches(11.7), Inches(1.1))
        tf = title_box.text_frame; tf.clear()
        p = tf.paragraphs[0]; p.text = item.get("title", f"Slide {i+1}")
        p.font.name = "Aptos Display"; p.font.size = Pt(30); p.font.bold = True
        p.font.color.rgb = RGBColor(*fg)
        body = slide.shapes.add_textbox(Inches(.95), Inches(1.9), Inches(7.8), Inches(4.7))
        tf = body.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
        for j, bullet in enumerate(item.get("bullets", [])[:5]):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.text = "•  " + str(bullet)
            p.font.name = "Aptos"; p.font.size = Pt(21); p.font.color.rgb = RGBColor(*fg)
            p.space_after = Pt(18)
        # right-side visual card
        card = slide.shapes.add_shape(1, Inches(9.1), Inches(2.1), Inches(3.25), Inches(3.8))
        card.fill.solid(); card.fill.fore_color.rgb = RGBColor(*accent); card.fill.transparency = 80
        card.line.color.rgb = RGBColor(*accent)
        vbox = slide.shapes.add_textbox(Inches(9.35), Inches(2.55), Inches(2.75), Inches(2.9))
        vtf = vbox.text_frame; vtf.word_wrap = True
        p = vtf.paragraphs[0]; p.text = "VISUAL IDEA"; p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = RGBColor(*accent)
        p = vtf.add_paragraph(); p.text = item.get("visual", "Add a clear diagram"); p.font.size = Pt(17); p.font.color.rgb = RGBColor(*fg)
        footer = slide.shapes.add_textbox(Inches(.8), Inches(7.0), Inches(11.7), Inches(.25))
        p = footer.text_frame.paragraphs[0]; p.text = f"{plan.get('title','AI Presentation')}   •   {i+1}/{len(plan['slides'])}"
        p.font.size = Pt(10); p.font.color.rgb = RGBColor(*muted)
    prs.save(out_path)

def make_tts(text, out_path):
    try:
        import pyttsx3
        engine = pyttsx3.init()
        engine.setProperty("rate", 165)
        engine.save_to_file(text, str(out_path))
        engine.runAndWait()
        return Path(out_path).exists() and Path(out_path).stat().st_size > 1000
    except Exception:
        return False

def create_video(plan, out_path, style, theme_name, fps, seconds_per_slide, voice_enabled):
    try:
        import imageio.v2 as imageio
        import numpy as np
    except Exception as e:
        raise RuntimeError("Install imageio and numpy: pip install imageio imageio-ffmpeg numpy") from e
    frames_dir = Path(tempfile.mkdtemp(prefix="motiondeck_"))
    slide_audio = []
    all_frames = []
    total = len(plan["slides"])
    for idx, slide in enumerate(plan["slides"]):
        selected_style = style
        if style == "All styles":
            selected_style = ["Motion graphics", "Animated presenter", "Cinematic storyboard"][idx % 3]
        audio_path = frames_dir / f"voice_{idx:03d}.wav"
        has_audio = voice_enabled and make_tts(slide.get("narration", ""), audio_path)
        if has_audio:
            slide_audio.append(str(audio_path))
        else:
            slide_audio.append(None)
        for f in range(fps * seconds_per_slide):
            progress = f / max(1, fps*seconds_per_slide-1)
            frame = render_frame(plan, slide, idx, total, selected_style, theme_name, progress)
            # gentle zoom/pan effect simulated with a small scale animation
            all_frames.append(np.array(frame))
    imageio.mimsave(str(out_path), all_frames, fps=fps, codec="libx264", quality=7)
    # If audio exists, combine each slide's narration with video via ffmpeg concat.
    # This implementation keeps the base MP4 portable even when local TTS is unavailable.
    if any(slide_audio):
        try:
            # Create a single narration track by concatenating audio files and silence gaps.
            # Use ffmpeg; silence is inserted for slides without audio.
            concat_file = frames_dir / "audio_list.txt"
            entries = []
            for i, audio in enumerate(slide_audio):
                if audio:
                    entries.append(f"file '{audio.replace(chr(39), chr(39)+chr(92)+chr(39)+chr(39))}'")
                else:
                    silence = frames_dir / f"silence_{i}.wav"
                    subprocess.run(["ffmpeg","-y","-f","lavfi","-i","anullsrc=r=22050:cl=mono","-t",str(seconds_per_slide),str(silence)],
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
                    entries.append(f"file '{silence}'")
            concat_file.write_text("\n".join(entries), encoding="utf-8")
            merged_audio = frames_dir / "narration.wav"
            subprocess.run(["ffmpeg","-y","-f","concat","-safe","0","-i",str(concat_file),"-c:a","pcm_s16le",str(merged_audio)],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
            muxed = frames_dir / "muxed.mp4"
            subprocess.run(["ffmpeg","-y","-i",str(out_path),"-i",str(merged_audio),"-shortest","-c:v","copy","-c:a","aac",str(muxed)],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
            if muxed.exists() and muxed.stat().st_size > 10000:
                import shutil
                shutil.copy2(muxed, out_path)
        except Exception:
            pass
    return out_path

# ---------- Main UI ----------
left, right = st.columns([1.5, 1])
with left:
    st.markdown("### ✍️ Your source content")
    uploaded = st.file_uploader("Upload a PDF, TXT, or Markdown file", type=["pdf","txt","md"])
    typed_text = st.text_area(
        "Or paste your text here",
        height=230,
        placeholder="Paste a lesson, article, project brief, or topic description..."
    )
    source_text = typed_text.strip() or extract_text(uploaded)
    st.caption(f"Source length: {len(source_text):,} characters")
with right:
    st.markdown("### 🎨 Output preview")
    st.markdown(f"""
    <div style="border:1px solid #303a5d;border-radius:18px;padding:18px;background:rgba(17,26,49,.8)">
      <div style="font-size:.8rem;color:#aeb9da">SELECTED STYLE</div>
      <div style="font-size:1.3rem;font-weight:700;margin:5px 0 12px">{video_style}</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <span style="background:#26365e;padding:6px 10px;border-radius:20px">📊 Editable PPTX</span>
        <span style="background:#26365e;padding:6px 10px;border-radius:20px">🎞️ MP4 video</span>
        <span style="background:#26365e;padding:6px 10px;border-radius:20px">🔊 Optional narration</span>
      </div>
      <p style="color:#aeb9da;margin-top:16px">Your text is transformed into slide outlines, visual ideas, narration, and rendered scenes.</p>
    </div>
    """, unsafe_allow_html=True)
    st.info("For stronger slide writing, install Ollama and pull a local instruction model. Built-in mode needs no API key.")

if st.button("✨ Generate PowerPoint + Animated Video", type="primary", use_container_width=True):
    if len(source_text.strip()) < 30:
        st.error("Please paste at least a few sentences or upload a text/PDF file.")
    else:
        try:
            with st.spinner("Planning your presentation..."):
                plan = build_plan(source_text, slide_count, model_mode, ollama_model)
            stamp = time.strftime("%Y%m%d_%H%M%S")
            ppt_path = OUTPUT_DIR / f"motiondeck_{stamp}.pptx"
            video_path = OUTPUT_DIR / f"motiondeck_{stamp}.mp4"
            json_path = OUTPUT_DIR / f"motiondeck_{stamp}.json"
            with st.spinner("Creating editable PowerPoint..."):
                create_pptx(plan, ppt_path, theme)
            st.success("PowerPoint created.")
            with st.spinner("Rendering animated video. This may take a few minutes..."):
                create_video(plan, video_path, video_style, theme, fps, seconds_per_slide, make_voice)
            json_path.write_text(json.dumps(plan, indent=2, ensure_ascii=False), encoding="utf-8")
            st.success("Your files are ready!")
            st.markdown("### 📦 Download your outputs")
            c1, c2, c3 = st.columns(3)
            c1.download_button("⬇️ PowerPoint (.pptx)", ppt_path.read_bytes(), file_name=ppt_path.name,
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
            c2.download_button("⬇️ Animated video (.mp4)", video_path.read_bytes(), file_name=video_path.name,
                               mime="video/mp4", use_container_width=True)
            c3.download_button("⬇️ Slide plan (.json)", json_path.read_bytes(), file_name=json_path.name,
                               mime="application/json", use_container_width=True)
            st.markdown("### 🧠 Generated slide plan")
            st.write(f"**Title:** {plan.get('title','Presentation')}")
            for i, slide in enumerate(plan["slides"], 1):
                with st.expander(f"Slide {i}: {slide.get('title','')}"):
                    st.write("**Bullets**")
                    for bullet in slide.get("bullets", []):
                        st.write(f"- {bullet}")
                    st.write("**Visual idea:**", slide.get("visual", ""))
                    st.write("**Narration:**", slide.get("narration", ""))
            st.video(video_path.read_bytes())
        except Exception as e:
            st.error(f"Generation failed: {e}")
            st.exception(e)

st.markdown("---")
st.markdown('<p class="small-note">Prototype note: the local presenter and cinematic scenes are animated illustrations/storyboards, not photorealistic AI video. For cinematic generative footage or realistic talking avatars, integrate a separately hosted/open-weight video or avatar model and verify its hardware requirements and license.</p>', unsafe_allow_html=True)
