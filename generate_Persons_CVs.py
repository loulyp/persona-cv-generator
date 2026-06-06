"""
Persona-Driven CV Generator    
Reads personas from a JSON file and generates CVs 
"""

import os
import json
import re
import random
import time
import requests
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable

from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

MALE_FIRST_NAMES = [
    "Mohammed", "Ahmad", "Mahmoud", "Abdulrahman", "Abdulaziz", "Abdullah",
    "Fahad", "Saud", "Khalid", "Turki", "Faisal", "Nawaf", "Sultan", "Bandar",
    "Majed", "Rayan", "Ziyad", "Mishaal", "Talal", "Hamad", "Yasser", "Omar",
    "Ali", "Hassan", "Hussein", "Ibrahim", "Yousef", "Samer", "Tamer", "Wael",
    "Karim", "Amjad", "Nasser", "Badr", "Fahd", "Adnan", "Firas", "Haitham",
    "Qasim", "Jamal", "Marwan", "Bilal", "Anas", "Sami", "Mazen", "Raed",
    "Shadi", "Osama", "Ayman", "Luay"
]

FEMALE_FIRST_NAMES = [
    "Abeer", "Aljawhara", "Reem", "Layan", "Nouf", "Hala", "Dalia", "Maha",
    "Raghad", "Sarah", "Huda", "Amal", "Dana", "Shahad", "Basma", "Lama",
    "Arwa", "Noura", "Jawaher", "Hind", "Fatima", "Aisha", "Khadija", "Maryam",
    "Zainab", "Yasmin", "Lina", "Rania", "Dina", "Salma", "Nour", "Farah",
    "Iman", "Samar", "Ghada", "Wafa", "Najwa", "Sahar", "Rabab", "Suha",
    "Manal", "Hanan", "Bushra", "Ilham", "Dalal", "Latifa", "Afaf", "Ruqayyah",
    "Sumaya", "Tamara"
]

FAMILY_NAMES = [
    "Al-Qahtani", "Al-Harbi", "Al-Otaibi", "Al-Mutairi", "Al-Dossari",
    "Al-Shammari", "Al-Zahrani", "Al-Ghamdi", "Al-Anazi", "Al-Rashidi",
    "Al-Balawi", "Al-Subaie", "Al-Faraj", "Al-Shehri", "Al-Malki",
    "Al-Harthi", "Al-Amri", "Al-Sabah", "Al-Ghanim", "Al-Kharafi",
    "Al-Roumi", "Al-Bahar", "Al-Nassar", "Al-Enezi", "Al-Nahyan",
    "Al-Maktoum", "Al-Thani", "Al-Khalifa", "Al-Marri", "Al-Mansoori",
    "Al-Hammadi", "Al-Masri", "Al-Sayed", "Al-Haddad", "Al-Najjar",
    "Al-Khatib", "Al-Salem", "Al-Ahmad", "Al-Yousef", "Al-Hussein",
    "Al-Qasim", "Al-Sharif", "Al-Badawi", "Al-Tamimi", "Al-Obeid",
    "Al-Jabri", "Al-Hakim"
]

# =============================================================
# SETTINGS
# =============================================================
OUTPUT_DIR  = "output_cvs5"
MODEL_NAME  = "llama3:8b"
OLLAMA_URL  = "http://localhost:11434/api/generate"
TIMEOUT_SEC = 240
INPUT_FILE  = "career_personas_100.json"  #  personas file

os.makedirs(OUTPUT_DIR, exist_ok=True)

# =============================================================
# LOAD PERSONAS FROM FILE
# =============================================================
def load_personas_from_file(filepath):
    """Load personas from JSON file """
    with open(filepath, 'r', encoding='utf-8') as f:
        data = json.load(f)
        
    if isinstance(data, list):
        if len(data) > 0 and isinstance(data[0], dict):
            if 'description' in data[0]:
                return [item['description'] for item in data]
                
    raise ValueError(f"Unknown file format in {filepath}")

# =============================================================
# SECTION SYNONYM HINT
# =============================================================
SECTION_SYNONYM_HINT = """
Pick ONLY the sections relevant to this person.
Use varied names — do not default to the most generic option:

  Summary     → "Profile" | "About Me" | "Professional Summary" | "Career Objective"
  Experience  → "Work Experience" | "Employment History" | "Professional Experience"
  Education   → "Education" | "Academic Background" | "Qualifications"
  Skills      → "Technical Skills" | "Core Competencies" | "Skills & Tools" | "Key Skills"
  Projects    → "Key Projects" | "Selected Projects" | "Personal Projects"
  Certs       → "Certifications" | "Professional Development"
  Languages   → "Languages" | "Language Skills"
  Publications → only for researchers / academics
  Volunteering → "Volunteer Work" | "Community Involvement"
  References  → a single line like "References available upon request"
"""

# =============================================================
# RANDOM CV STYLE GENERATOR
# =============================================================

# Color pools
ACCENT_COLORS = [
    "#1a3c5e",  # Navy (finance, corporate)
    "#2d5a3f",  # Forest green (healthcare, environment)
    "#6b2e3e",  # Burgundy (law, executive)
    "#2c6e6e",  # Teal (tech, creative)
    "#4a2c6e",  # Purple (education, arts)
    "#8b4513",  # Saddle brown (construction, traditional)
    "#b22222",  # Firebrick (sales, energetic)
    "#3a6b4b",  # Sage (consulting)
    "#7a4c2c",  # Coffee (hospitality)
    "#1e5f6e",  # Deep cyan (engineering)
]

# Font pools
FONT_PAIRS = [
    ("Helvetica", "Helvetica-Bold"),           # Modern sans-serif
    ("Times-Roman", "Times-Bold"),             # Classic serif
    ("Courier", "Courier-Bold"),               # Technical/monospace
    ("Helvetica", "Times-Bold"),               # Mixed modern/classic
    ("Times-Roman", "Helvetica-Bold"),         # Mixed classic/modern
]

# Layout variations
LAYOUT_STYLES = [
    {"dividing_line": "thick", "spacing": "loose", "bullet": "•"},
    {"dividing_line": "thin", "spacing": "tight", "bullet": "-"},
    {"dividing_line": "double", "spacing": "medium", "bullet": "▪"},
    {"dividing_line": "none", "spacing": "loose", "bullet": "→"},
    {"dividing_line": "dashed", "spacing": "tight", "bullet": "★"},
]

# Section name capitalization styles
CAP_STYLES = [
    "upper",      # WORK EXPERIENCE
    "title",      # Work Experience
    "small",      # WORK EXPERIENCE (small caps)
    "normal",     # Work experience
]

def get_random_style():
    """Generate random styling for a CV"""
    return {
        "accent_color": random.choice(ACCENT_COLORS),
        "font_body": random.choice(FONT_PAIRS)[0],
        "font_bold": random.choice(FONT_PAIRS)[1],
        "layout": random.choice(LAYOUT_STYLES),
        "cap_style": random.choice(CAP_STYLES),
        "has_photo": random.random() < 0.2,  # 20% have photo placeholder
        "has_border": random.random() < 0.3,  # 30% have page border
        "alignment": random.choice(["left", "justified"]),
    }
# =============================================================
# Snames
# =============================================================
def generate_gulf_name(gender="male"):
    """Generate name with 50% 2-part, 50% 3-part"""
    if gender == "male":
        first = random.choice(MALE_FIRST_NAMES)
    else:
        first = random.choice(FEMALE_FIRST_NAMES)
    
    last = random.choice(FAMILY_NAMES)
    
    # % chance for 3-part name
    if random.random() < 0.5:
        if gender == "male":
            middle = random.choice(MALE_FIRST_NAMES)
        else:
            middle = random.choice(FEMALE_FIRST_NAMES)
        # Avoid first == middle
        while middle == first:
            middle = random.choice(MALE_FIRST_NAMES if gender == "male" else FEMALE_FIRST_NAMES)
        full_name = f"{first} {middle} {last}"
        return full_name, first, middle, last
    else:
        return f"{first} {last}", first, None, last


# =============================================================
# PROMPT
# =============================================================
def build_prompt(description: str, full_name: str, gender: str) -> str:
    return f"""You are an expert CV writer. Given only the persona description below,
create one complete, realistic CV as a JSON object.

PERSONA DESCRIPTION:
"{description}"

REQUIRED NAME (MUST USE EXACTLY): {full_name}
GENDER: {gender}

CRITICAL INSTRUCTION - USE WHAT'S IN THE PERSONA:
- If the persona mentions a country (e.g., "Botswana", "Colombia", "Austria"), use that country.
- If the persona mentions a city or region, use that location.
- If the persona mentions a specific company, university, or organization, USE IT EXACTLY.
- If the persona mentions a nationality (e.g., "Bahamian", "Seychellois", "Colombian"), reflect that in location and context.
- DO NOT override persona details with generic Saudi names.
- Invent only what's missing, never contradict what's given.

YOUR TASKS:
1. USE THE EXACT NAME PROVIDED: {full_name}
2. Invent plausible contact details (email, phone, city).
3. Decide the experience leve, and appropriate length based on what the persona realstically.
4. Choose which sections to include, what to name them, and in what order.
5. Fill every section with rich, specific, believable content — no placeholders.

LENGTH & DENSITY — treat these as hard requirements:

- Junior persona (0–3 years experience) → at least 1 full page of dense content.
- Mid-level persona (3–7 years) → at least 1.5 pages.
- Senior persona (7+ years) → at least 2 full pages of dense content.
- Every job entry → 5 bullet points minimum.
- Education entries → include GPA or grade, relevant coursework or thesis title.
- Summary / Profile / Objective → 3 to 5 full sentences.
- Skills list → at least 5 specific items.
- Academic personas → Publications section with at least 1 entry.

REALISM RULES (no synthetic artifacts):
- No repeated phrases across different jobs or sections.
- Each bullet point must be unique and specific (include metric or outcome when possible).
- Avoid listing 15+ unrelated technical skills for non-technical personas.
- Use realistic company names.

=== COMPANY & UNIVERSITY NAMES (USE THESE OR INVENT SIMILAR) ===

SAUDI COMPANIES (realistic):
- Saudi Telecom Company (STC)
- SABIC (Saudi Basic Industries Corporation)
- Aramco (Saudi Arabian Oil Company)
- Al Rajhi Bank
- Riyad Bank
- NCB (National Commercial Bank)
- Mobily (Etihad Etisalat)
- Zain Saudi Arabia
- Ma'aden (Saudi Arabian Mining Co.)
- SEC (Saudi Electricity Company)
- King Faisal Specialist Hospital & Research Centre
- Dr. Sulaiman Al Habib Medical Group
- Almarai Company
- Savola Group
- Jarir Bookstore
- Abdullah Al Othaim Markets
- Tawakkalna (government digital platform)
- Elm (government digital solutions)
- NEOM (giga-project)
- Red Sea Global
- ROSHN (real estate)
- ACWA Power

SAUDI UNIVERSITIES (realistic):
- King Saud University (KSU) - Riyadh
- King Abdulaziz University (KAU) - Jeddah
- King Fahd University of Petroleum and Minerals (KFUPM) - Dhahran
- Princess Nourah bint Abdulrahman University - Riyadh
- Imam Muhammad bin Saud Islamic University - Riyadh
- Umm Al-Qura University - Makkah
- Taibah University - Medina
- King Khalid University - Abha
- Qassim University - Buraidah
- Islamic University of Madinah
- Alfaisal University - Riyadh (private)
- Prince Sultan University - Riyadh (private)
- Effat University - Jeddah (private)
- Dar Al Uloom University - Riyadh (private)

INTERNATIONAL COMPANIES WITH SAUDI OFFICES (realistic):
- Google Saudi Arabia
- Microsoft Arabia
- Deloitte Middle East
- PwC Middle East
- EY Saudi Arabia
- KPMG Saudi Arabia
- McKinsey & Company Saudi Arabia
- Boston Consulting Group (BCG) Riyadh
- Bain & Company Saudi Arabia
- Accenture Saudi Arabia
- IBM Saudi Arabia
- Oracle Saudi Arabia
- SAP Saudi Arabia
- Cisco Saudi Arabia
- Amazon KSA (Souq.com)
- Uber Saudi Arabia
- Careem KSA

JOB TITLES (realistic Arabic-English mix, use English):
- Junior / Senior / Lead / Principal
- Analyst, Associate, Specialist
- Coordinator, Administrator, Officer
- Engineer (Software, Network, Civil, Mechanical)
- Developer (Frontend, Backend, Full Stack, Mobile)
- Project Manager, Program Manager, Product Manager
- IT Manager, CTO, CIO
- Director (Regional, Operations, Sales, Marketing)
- Consultant (Management, Technical, Strategy)
- Surgeon, Physician, Nurse, Pharmacist
- Teacher, Lecturer, Assistant Professor, Professor
- Accountant, Financial Analyst, Auditor
- HR Specialist, Recruiter, Talent Acquisition
- Marketing Specialist, Digital Marketer, SEO Specialist
- Sales Executive, Account Manager, Business Development

=== END OF REFERENCE LISTS === 
SECTION SYNONYMS (use these or your own, but be consistent):
- Summary can be called "Professional Summary", "Profile", "About Me"
- Experience can be "Work Experience", "Employment History"
- Education can be "Education", "Academic Background"
- Skills can be "Core Competencies", "Technical Skills"
- Languages can be "Language Proficiency"

SECTION GUIDANCE:
{SECTION_SYNONYM_HINT}

IMPORTANT: Return ONLY valid JSON. Start with {{ and end with }}. No text before or after.

EXAMPLE SHAPE:
{{
  "name": "...",
  "email": "...",
  "phone": "...",
  "location": "...",
  "Professional Summary": "...",
  "Work Experience": [
    {{
      "title": "...",
      "company": "...",
      "dates": "...",
      "bullets": ["bullet1", "bullet2", "bullet3", "bullet4", "bullet5"]
    }}
  ],
  "Education": [
    {{
      "degree": "...",
      "institution": "...",
      "year": "...",
      "notes": "..."
    }}
  ],
  "Skills": ["skill1", "skill2", "skill3", "skill4", "skill5", "skill6", "skill7", "skill8", "skill9", "skill10", "skill11", "skill12"],
  "Languages": ["Language1 (fluent)", "Language2 (native)"]
}}
"""

# =============================================================
# LLM CALL WITH RETRY
# =============================================================
def call_llm_with_retry(prompt: str, max_retries: int = 2) -> str:
    """Call LLM with retry on failure"""
    for attempt in range(max_retries):
        try:
            r = requests.post(
                OLLAMA_URL,
                json={
                    "model": MODEL_NAME,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "num_predict": 4096,
                        "temperature": 0.7,
                    },
                },
                timeout=TIMEOUT_SEC,
            )
            r.raise_for_status()
            return r.json()["response"].strip()
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            print(f"    Retry {attempt + 1}/{max_retries} after error: {e}")
            time.sleep(2)
    return ""

# =============================================================
# IMPROVED JSON EXTRACTION
# =============================================================
def extract_json(raw: str) -> dict:
    """Robust JSON extraction with multiple strategies"""
    
    # Strategy 1: Remove markdown and try direct parse
    cleaned = re.sub(r"```(?:json)?|```", "", raw).strip()
    
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass
    
    # Strategy 2: Find JSON object boundaries
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    
    if start != -1 and end != -1 and end > start:
        json_str = cleaned[start:end+1]
        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            pass
    
    # Strategy 3: Try to fix common JSON issues
    if start != -1:
        json_str = cleaned[start:]
        
        # Fix trailing commas
        json_str = re.sub(r',\s*}', '}', json_str)
        json_str = re.sub(r',\s*]', ']', json_str)
        
        # Fix missing quotes around keys
        json_str = re.sub(r'([{,])\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*:', r'\1"\2":', json_str)
        
        # Fix unquoted strings
        json_str = re.sub(r':\s*([a-zA-Z][a-zA-Z0-9\s]*[a-zA-Z])([,}])', r':"\1"\2', json_str)
        
        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            pass
    
    # Strategy 4: Try to complete truncated JSON
    if start != -1:
        json_str = cleaned[start:]
        
        # Count braces
        brace_count = json_str.count('{') - json_str.count('}')
        bracket_count = json_str.count('[') - json_str.count(']')
        
        if brace_count > 0:
            json_str += '}' * brace_count
        if bracket_count > 0:
            json_str += ']' * bracket_count
        
        # Ensure it ends with }
        if not json_str.endswith('}'):
            json_str += '}'
        
        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            pass
    
    # All strategies failed - save for debugging
    debug_path = os.path.join(OUTPUT_DIR, f"_debug_{abs(hash(raw)) % 9999:04d}.txt")
    with open(debug_path, "w", encoding="utf-8") as f:
        f.write(raw)
    print(f"    ⚠️  Could not parse JSON - raw saved to {debug_path}")
    
    return {"_parse_error": True, "_raw": raw[:500]}

# =============================================================
# SECTION HELPERS
# =============================================================
CONTACT_KEYS = {"name", "email", "phone", "location", "linkedin",
                "github", "website", "address", "_meta", "_description", "_parse_error", "_raw"}

def get_sections(cv: dict):
    return [(k, v) for k, v in cv.items()
            if k.lower() not in CONTACT_KEYS and not k.startswith("_")]

def safe_str(val) -> str:
    if val is None:
        return ""
    return str(val)

def normalise_section_value(val):
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        return str(val)
    if isinstance(val, dict):
        return [val]
    return val

# =============================================================
# PDF WRITER
# =============================================================
def write_pdf(cv: dict, filepath: str):
    def clean_text(text):
        """Remove problematic characters and fix spacing"""
        if not text:
            return ""
        # Replace common problematic chars
        text = text.replace('', '•')
        text = text.replace('', '•')
        # Fix missing spaces after section headers
        text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
        # Remove duplicate spaces
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
        
    # Get random style for this CV
    style = get_random_style()
    
    doc = SimpleDocTemplate(
        filepath, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm,
    )
    styles = getSampleStyleSheet()
    
    # Apply random styles
    name_s = ParagraphStyle("N", parent=styles["Normal"],
                           fontSize=random.choice([16, 18, 20, 22]),
                           fontName=style["font_bold"],
                           spaceAfter=random.choice([2, 4, 6]),
                           textColor=colors.HexColor(style["accent_color"]),
                           alignment=0)  # Center aligned
    
    contact_s = ParagraphStyle("C", parent=styles["Normal"],
                               fontSize=random.choice([8, 9, 10]),
                               textColor=colors.HexColor("#555555"),
                               spaceAfter=random.choice([6, 8, 10]),
                               alignment=0)
    
    # Section style based on cap_style
    section_font_size = random.choice([10, 11, 12])
    section_name = style["cap_style"]
    
    section_s = ParagraphStyle("S", parent=styles["Normal"],
                               fontSize=section_font_size,
                               fontName=style["font_bold"],
                               spaceBefore=random.choice([8, 10, 12]),
                               spaceAfter=random.choice([2, 3, 4]),
                               textColor=colors.HexColor(style["accent_color"]),
                               alignment=0)
    
    body_s = ParagraphStyle("B", parent=styles["Normal"],
                           fontSize=random.choice([9, 9.5, 10]),
                           leading=random.choice([13, 14, 15]),
                           spaceAfter=random.choice([2, 3, 4]),
                           alignment=1 if style["alignment"] == "justified" else 0)
    
    role_s = ParagraphStyle("R", parent=styles["Normal"],
                           fontSize=random.choice([9.5, 10, 10.5]),
                           fontName=style["font_bold"],
                           spaceAfter=1)
    
    sub_s = ParagraphStyle("Sub", parent=styles["Normal"],
                          fontSize=random.choice([8, 8.5, 9]),
                          textColor=colors.HexColor("#666666"),
                          spaceAfter=1)
    
    bullet_char = style["layout"]["bullet"]
    bullet_s = ParagraphStyle("Bl", parent=styles["Normal"],
                             fontSize=random.choice([9, 9.5, 10]),
                             leading=random.choice([12, 13, 14]),
                             leftIndent=random.choice([12, 14, 16]),
                             spaceAfter=random.choice([1, 2, 3]))
    
    story = []
    
    # Random name formatting
    name_text = cv.get("name", "")
    if random.random() < 0.3:
        name_text = name_text.upper()
    story.append(Paragraph(name_text, name_s))
    
    # Contact info
    parts = [cv.get(k, "") for k in ("email", "phone", "location") if cv.get(k)]
    if parts:
        contact_text = "  |  ".join(parts)
        if random.random() < 0.2:
            contact_text = contact_text.replace(" | ", " • ")
        story.append(Paragraph(contact_text, contact_s))
    
    # Divider based on layout
    if style["layout"]["dividing_line"] == "thick":
        story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor(style["accent_color"]), spaceAfter=6))
    elif style["layout"]["dividing_line"] == "thin":
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#aaaaaa"), spaceAfter=4))
    elif style["layout"]["dividing_line"] == "double":
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor(style["accent_color"]), spaceAfter=2))
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor(style["accent_color"]), spaceAfter=6))
    elif style["layout"]["dividing_line"] == "dashed":
        # Dashed line approximation
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#999999"), spaceAfter=6, dash=(3,3)))
    # else "none" - skip divider
    
    # Process sections with random capitalization
    for sec, raw_val in get_sections(cv):
        val = normalise_section_value(raw_val)
        
        # Apply cap_style to section name
        if style["cap_style"] == "upper":
            section_title = sec.upper()
        elif style["cap_style"] == "title":
            section_title = sec.title()
        else:
            section_title = sec
        
        story.append(Paragraph(section_title, section_s))
        
        # Add small divider after section (varied)
        if style["layout"]["dividing_line"] != "none" and random.random() < 0.7:
            story.append(HRFlowable(width=str(random.randint(30, 70)) + "%", thickness=0.3,
                                   color=colors.HexColor(style["accent_color"]), spaceAfter=4))
        
        if isinstance(val, str):
            story.append(Paragraph(clean_text(val), body_s))
        elif isinstance(val, list):
            if all(isinstance(x, str) for x in val):
                for skill in val:
                    story.append(Paragraph(f"{bullet_char} {safe_str(skill)}", bullet_s))
            else:
                for item in val:
                    if not isinstance(item, dict):
                        story.append(Paragraph(safe_str(item), body_s))
                        continue
                    title = safe_str(item.get("title") or item.get("degree") or item.get("name",""))
                    org = safe_str(item.get("company") or item.get("institution") or "")
                    dates = safe_str(item.get("dates") or item.get("year") or "")
                    parts = [p for p in (title, org) if p]
                    if parts:
                        # Random date position (left or right aligned)
                        if random.random() < 0.5:
                            story.append(Paragraph(" — ".join(parts), role_s))
                            if dates:
                                story.append(Paragraph(dates, sub_s))
                        else:
                            # Date on same line as title
                            title_with_date = f"{' — '.join(parts)} | {dates}" if dates else " — ".join(parts)
                            story.append(Paragraph(title_with_date, role_s))
                    
                    for key in ("notes", "description", "summary"):
                        if item.get(key):
                            story.append(Paragraph(safe_str(item[key]), body_s))
                    for b in item.get("bullets", []):
                        clean_b = re.sub(r'^[\s•\-★▪→*]+', '', safe_str(b))
                        p = doc.add_paragraph(style="List Bullet")
                        r = p.add_run(clean_b)  # ← Remove {bullet_char}
                    story.append(Paragraph(f"{bullet_char} {clean_b}", bullet_s))
           
    
    doc.build(story)

# =============================================================
# DOCX WRITER
# =============================================================
def write_docx(cv: dict, filepath: str):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    # Get random style for this CV
    style = get_random_style()
    
    doc = DocxDocument()
    
    # Random page margins
    margin = random.choice([0.7, 0.8, 0.9, 1.0])
    for sec in doc.sections:
        sec.left_margin = sec.right_margin = Inches(margin)
        sec.top_margin = sec.bottom_margin = Inches(margin - 0.1)
    
    # Random color for headings
    accent_rgb = RGBColor(
        int(style["accent_color"][1:3], 16),
        int(style["accent_color"][3:5], 16),
        int(style["accent_color"][5:7], 16)
    )
    
    # Random font size for name
    name_size = random.choice([16, 18, 20, 22])
    name_bold = random.choice([True, False])
    
    # Grey color for contact info
    GREY = RGBColor(0x55, 0x55, 0x55)
    
    def divider(p, thick=False):
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bot = OxmlElement("w:bottom")
        
        if style["layout"]["dividing_line"] == "thick":
            bot.set(qn("w:sz"), "12")
            bot.set(qn("w:color"), style["accent_color"][1:])
        elif style["layout"]["dividing_line"] == "thin":
            bot.set(qn("w:sz"), "4")
            bot.set(qn("w:color"), "aaaaaa")
        elif style["layout"]["dividing_line"] == "double":
            bot.set(qn("w:sz"), "6")
            bot.set(qn("w:color"), style["accent_color"][1:])
        else:  # none or dashed
            return
        
        bot.set(qn("w:val"), "single")
        pBdr.append(bot)
        pPr.append(pBdr)
    
    # Name
    p = doc.add_paragraph()
    r = p.add_run(cv.get("name", ""))
    r.bold = name_bold
    r.font.size = Pt(name_size)
    r.font.color.rgb = accent_rgb
    p.paragraph_format.space_after = Pt(random.choice([2, 4, 6]))
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Contact info
    parts = [cv.get(k, "") for k in ("email", "phone", "location") if cv.get(k)]
    if parts:
        separator = random.choice([" | ", " • ", "  |  ", " — "])
        contact_text = separator.join(parts)
        p = doc.add_paragraph()
        r = p.add_run(contact_text)
        r.font.size = Pt(random.choice([8, 9, 10]))
        r.font.color.rgb = GREY
        p.paragraph_format.space_after = Pt(random.choice([6, 8, 10]))
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        divider(p, thick=True)
    
    # Sections
    bullet_char = style["layout"]["bullet"]
    
    for sec, raw_val in get_sections(cv):
        val = normalise_section_value(raw_val)
        
        # Apply cap_style
        if style["cap_style"] == "upper":
            section_title = sec.upper()
        elif style["cap_style"] == "title":
            section_title = sec.title()
        else:
            section_title = sec
        
        p = doc.add_paragraph()
        r = p.add_run(section_title)
        r.bold = True
        r.font.size = Pt(random.choice([10, 11, 12]))
        r.font.color.rgb = accent_rgb
        p.paragraph_format.space_before = Pt(random.choice([8, 10, 12]))
        p.paragraph_format.space_after = Pt(random.choice([2, 3, 4]))
        divider(p)
        
        if isinstance(val, str):
            p = doc.add_paragraph(val)
            for run in p.runs:
                run.font.size = Pt(random.choice([9, 9.5, 10]))
        elif isinstance(val, list):
            if all(isinstance(x, str) for x in val):
                for skill in val:
                    p = doc.add_paragraph(style="List Bullet")
                    r = p.add_run(f"{bullet_char} {safe_str(skill)}")
                    r.font.size = Pt(random.choice([9, 9.5, 10]))
            else:
                for item in val:
                    if not isinstance(item, dict):
                        doc.add_paragraph(safe_str(item))
                        continue
                    
                    title = safe_str(item.get("title") or item.get("degree") or item.get("name", ""))
                    org = safe_str(item.get("company") or item.get("institution") or "")
                    dates = safe_str(item.get("dates") or item.get("year") or "")
                    parts_list = [x for x in (title, org) if x]
                    
                    if parts_list:
                        # Random date position
                        if random.random() < 0.5:
                            p = doc.add_paragraph()
                            r = p.add_run(" — ".join(parts_list))
                            r.bold = True
                            r.font.size = Pt(random.choice([9.5, 10, 10.5]))
                            if dates:
                                p = doc.add_paragraph(dates)
                                for run in p.runs:
                                    run.font.size = Pt(random.choice([8, 8.5, 9]))
                                    run.font.color.rgb = GREY
                        else:
                            title_with_date = f"{' — '.join(parts_list)} | {dates}" if dates else " — ".join(parts_list)
                            p = doc.add_paragraph()
                            r = p.add_run(title_with_date)
                            r.bold = True
                            r.font.size = Pt(random.choice([9.5, 10, 10.5]))
                    
                    for key in ("notes", "description", "summary"):
                        if item.get(key):
                            p = doc.add_paragraph(safe_str(item[key]))
                            for run in p.runs:
                                run.font.size = Pt(random.choice([9, 9.5, 10]))
                    
                    for b in item.get("bullets", []):
                        clean_b = re.sub(r'^[\s•\-★▪→*]+', '', safe_str(b))
                        p = doc.add_paragraph(style="List Bullet")
                        r = p.add_run(f"{bullet_char} {clean_b}")
                        r.font.size = Pt(9.5)
                   
                    
                    doc.add_paragraph().paragraph_format.space_after = Pt(random.choice([2, 3, 4]))
    
    doc.save(filepath)
# =============================================================
# MAIN
# =============================================================
if __name__ == "__main__":
    # Load personas from file
    if not os.path.exists(INPUT_FILE):
        print(f"❌ File not found: {INPUT_FILE}")
        print(f"Please create {INPUT_FILE} with your personas first.")
        exit(1)
    
    print(f"📖 Loading personas from {INPUT_FILE}...")
    descriptions = load_personas_from_file(INPUT_FILE)
    print(f"✅ Loaded {len(descriptions)} personas\n")
    
    # Generate CVs
    formats = ["pdf", "docx"] * (len(descriptions) // 2 + 1)
    random.shuffle(formats)
    
    success_count = 0
    fail_count = 0
    
    print(f"🎯 Generating {len(descriptions)} CVs  →  '{OUTPUT_DIR}/'\n")
    
    for i, description in enumerate(descriptions, start=1):
        fmt = formats[i - 1]
        short_desc = description[:60] + "..." if len(description) > 60 else description
        
        # Determine gender (50/50 split based on index)
        if i <= 50:
            gender = "male"
        else:
            gender = "female"
        
        # Generate name for this persona
        full_name, first_name, middle_name, last_name = generate_gulf_name(gender)
        
        print(f"[{i}/{len(descriptions)}] {full_name} ({gender}) - \"{short_desc}\"  [{fmt}]")
        
        try:
            # Call LLM with the prompt (pass name and gender)
            raw_response = call_llm_with_retry(build_prompt(description, full_name, gender))
            
            # Extract JSON
            cv = extract_json(raw_response)
            
            if cv.get("_parse_error"):
                print(f"    ⚠️  Parse failed - skipping this persona")
                fail_count += 1
                continue
            
            # Override the name to ensure it matches what we generated
            cv["name"] = full_name
            cv["_description"] = description
            
            # Generate filename using our name
            safe_name = re.sub(r"[^\w]", "_", full_name)[:50]
            filepath = os.path.join(OUTPUT_DIR, f"{i:03d}_{safe_name}.{fmt}")
            
            # Write CV
            if fmt == "pdf":
                write_pdf(cv, filepath)
            else:
                write_docx(cv, filepath)
            
            sections = [k for k, _ in get_sections(cv)]
            print(f"    ✅ {full_name}  |  {', '.join(sections[:3])}")
            print(f"    💾 {filepath}\n")
            success_count += 1
            # 🔥 COOLING BREAK - every 3 CVs
            if success_count % 5 == 0:
                print(f"    🌡️  Cooling break for 30 seconds...")
                time.sleep(30)
            
        except requests.exceptions.ConnectionError:
            print("    ❌ Cannot reach Ollama — run: ollama serve")
            break
        except Exception as e:
            print(f"    ❌ Error: {e}\n")
            fail_count += 1
    
    print("=" * 60)
    print(f"📊 SUMMARY:")
    print(f"   ✅ Success: {success_count}")
    print(f"   ❌ Failed: {fail_count}")
    print(f"   📁 Output: {OUTPUT_DIR}/")
    print("=" * 60)
