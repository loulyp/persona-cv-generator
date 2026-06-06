# persona-cv-generator
Perfect! Now I can see your complete pipeline. You have:

## How This Dataset Was Generated

### Overview

The dataset was created in two main phases:

1. **Filtering** – Selecting career-relevant personas from PersonaHub
2. **Generation** – Converting personas into realistic CVs using LLaMA-3 and rendering them as PDF/DOCX

---

### Phase 1: Filtering PersonaHub

The original dataset `proj-persona/PersonaHub` contains 200,000 personas. To create a career-focused CV dataset, I filtered for personas with job-related keywords.

| Step | Description |
|------|-------------|
| **Input** | `proj-persona/PersonaHub` – 200,000 personas |
| **Filter method** | Keyword matching (no LLM used for filtering) |
| **Keywords** | 100+ career terms: `software engineer`, `manager`, `professor`, `analyst`, `consultant`, `nurse`, `accountant`, etc. |
| **Result** | 100 career-ready personas |
| **Output file** | `career_personas_100.json` |

**Notebook used:** `filtteringpersonas.ipynb`

The filtering script:
- Loads PersonaHub dataset from Hugging Face
- Scans each persona description for career keywords
- Collects the first 100 matching personas
- Saves to JSON for CV generation

**Example filtered personas:**
```
1. an IT project manager who adopted extreme programming (XP) methodologies
2. A software engineer specializing in document management systems
3. an orthopedic surgeon relatively new to AOSSM
4. A senior journalist with years of experience in investigative reporting
5. A university professor of gender studies
```

---

### Phase 2: Generating CVs from Personas

The filtered personas are fed into a pipeline that generates complete, realistic CVs.

```
Persona Description → LLaMA-3 → JSON CV → PDF/DOCX
```

#### Step 1: LLM Generation (Persona → JSON)

Each persona description is sent to **LLaMA-3 (8B)** via Ollama with a structured prompt. The prompt instructs the model to create a complete CV as JSON.

**LLM Settings:**
| Parameter | Value |
|-----------|-------|
| Model | `llama3:8b` (via Ollama) |
| Temperature | `0.7` |
| Max tokens | `4096` |

**The prompt requires:**
- Use the exact name provided (generated from Gulf name lists)
- Respect all persona details (country, city, company, university)
- Minimum content density: 1–2 pages depending on experience level
- 5+ bullet points per job entry
- Realistic company and university names (Saudi and international)

**Example of generated JSON structure:**
```json
{
  "name": "Mohammed Al-Qahtani",
  "email": "m.alqahtani@example.com",
  "phone": "+966 5X XXX XXXX",
  "location": "Riyadh, Saudi Arabia",
  "Professional Summary": "...",
  "Work Experience": [
    {
      "title": "Senior Software Engineer",
      "company": "SABIC",
      "dates": "2020 - Present",
      "bullets": ["Led migration of legacy systems...", "..."]
    }
  ],
  "Education": [...],
  "Skills": [...],
  "Languages": ["Arabic (native)", "English (fluent)"]
}
```

#### Step 2: Name Generation

Before calling the LLM, each persona is assigned a realistic Gulf name:

| Gender | Name pools |
|--------|------------|
| Male | 50 first names (Mohammed, Ahmad, Abdullah, Fahad, Khalid, etc.) |
| Female | 50 first names (Abeer, Reem, Nouf, Hala, Fatima, etc.) |
| Family | 46 family names (Al-Qahtani, Al-Harbi, Al-Otaibi, Al-Mutairi, etc.) |

Each CV gets a unique name with either 2-part or 3-part format (e.g., `Mohammed Al-Qahtani` or `Mohammed Abdullah Al-Qahtani`).

#### Step 3: Random CV Styling

To make CVs visually diverse, each document receives random styling:

| Style Element | Options |
|---------------|---------|
| Accent color | Navy, forest green, burgundy, teal, purple, brown, etc. (10 colors) |
| Font pair | Helvetica/Times/Courier combinations |
| Section header case | UPPER, Title Case, normal, Small Caps |
| Layout | Thick/thin/double/dashed dividers, different spacing |
| Bullet symbol | `•` `-` `▪` `→` `★` |
| Date position | Left-aligned or right-aligned with title |

#### Step 4: Output Formats

Each CV is generated in **both PDF and DOCX formats** (alternating):

| Format | Library | Use case |
|--------|---------|----------|
| PDF | `reportlab` | Print-ready, fixed layout |
| DOCX | `python-docx` | Editable, recruiter-ready |

**Script used:** `generate_cv5.py`

---

### Summary Diagram

```
PersonaHub (200,000 personas)
        ↓
[Keyword Filtering] filteringpersonas.ipynb
        ↓
career_personas_100.json (100 personas)
        ↓
[LLM Generation + Name Generation] generate_cv5.py
        ↓
    ┌───────────────┴───────────────┐
    ↓                               ↓
PDF files (50)              DOCX files (50)
```

---

### Requirements to Reproduce

| Dependency | Version/Purpose |
|------------|-----------------|
| `datasets` | Loading PersonaHub |
| `ollama` | Running LLaMA-3 locally |
| `reportlab` | PDF generation |
| `python-docx` | DOCX generation |
| `requests` | API calls to Ollama |

**Note:** Ollama must be running locally with `llama3:8b` pulled:
```bash
ollama pull llama3:8b
ollama serve
```
