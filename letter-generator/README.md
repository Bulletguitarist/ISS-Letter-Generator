# 📄 ISS Letter Generator
**Innovative Staffing Solutions — Bulk Document Automation**

Generate Offer Letters, Internship Certificates, and LORs for 1–1000+ candidates in one click.

---

## 🚀 Quick Start

### 1. Install Python dependencies
```bash
pip install streamlit pandas python-docx docxtpl groq openpyxl
```

### 2. Generate sample templates (first-time setup)
```bash
python generate_templates.py
```
This creates three `.docx` templates in `/templates`:
- `Offer_Letter_Template.docx`
- `Internship_Certificate_Template.docx`
- `LOR_Template.docx`

### 3. Run the web app
```bash
streamlit run app.py
```
Open `http://localhost:8501` in your browser.

---

## 📁 Folder Structure
```
letter-generator/
├── app.py                         ← Main Streamlit web app
├── batch_generate.py              ← CLI batch processor
├── generate_templates.py          ← Template generator script
├── requirements.txt
├── templates/
│   ├── Offer_Letter_Template.docx
│   ├── Internship_Certificate_Template.docx
│   └── LOR_Template.docx
├── sample_data/
│   └── sample_candidates.csv
└── output/                        ← Generated letters saved here
```

---

## 📋 CSV Format

Your CSV must have a `name` column. All other columns are optional (defaults will be used if missing).

```csv
name,designation,domain,department,joining_date,last_working_date,email,duration,performance,skills,basic_salary,hra,special_allowance,gross_salary,net_salary,monthly_ctc,annual_ctc,probation_period
Rahul Sharma,Web Dev Intern,Web Development,Technology,01 Feb 2026,01 May 2026,rahul@gmail.com,3 months,Excellent,"HTML CSS React",10000,5000,3000,18000,18000,20000,240000,3 months
```

### All Supported Columns

| Column | Used In | Description |
|--------|---------|-------------|
| `name` | All | Full candidate name *(required)* |
| `designation` | All | Job title / role |
| `domain` | All | Skill domain (e.g. Web Development) |
| `department` | Offer Letter | Department name |
| `joining_date` | All | Start date |
| `last_working_date` | Cert / LOR | End date |
| `email` | All | Email address |
| `duration` | Certificate | Duration string (e.g. "3 months") |
| `performance` | Cert / LOR | Rating (Excellent, Good, etc.) |
| `skills` | Cert / LOR | Comma-separated skills list |
| `basic_salary` | Offer Letter | Basic salary amount |
| `hra` | Offer Letter | House Rent Allowance |
| `special_allowance` | Offer Letter | Special Allowance |
| `gross_salary` | Offer Letter | Gross monthly salary |
| `net_salary` | Offer Letter | Take-home salary |
| `monthly_ctc` | Offer Letter | Monthly CTC |
| `annual_ctc` | Offer Letter | Annual CTC |
| `probation_period` | Offer Letter | Probation duration |

---

## 🏷️ Template Placeholders

Use `{{placeholder}}` syntax in your DOCX templates:

### Standard Placeholders
```
{{name}}              {{designation}}        {{domain}}
{{department}}        {{joining_date}}        {{last_working_date}}
{{email}}             {{duration}}            {{performance}}
{{skills}}            {{today_date}}          {{current_year}}
{{basic_salary}}      {{hra}}                 {{special_allowance}}
{{gross_salary}}      {{net_salary}}          {{monthly_ctc}}
{{annual_ctc}}        {{probation_period}}
{{company_name}}      {{hr_name}}             {{hr_designation}}
```

### AI-Generated Placeholders (require Groq API key)
```
{{ai_generated_content}}       ← Smart, context-aware content for any letter
{{ai_internship_description}}  ← Full internship certificate body paragraph
{{ai_lor_body}}                ← Full 3-paragraph LOR body
{{ai_performance_summary}}     ← One-sentence performance summary
{{ai_skills_summary}}          ← One-sentence skills summary
```

---

## 🤖 Groq API Setup

1. Go to https://console.groq.com/keys
2. Create a new API key
3. Either:
   - Paste it in the **Groq API Key** field in the sidebar, OR
   - Set environment variable: `export GROQ_API_KEY=gsk_...`

The app uses `llama3-70b-8192` by default (configurable in the sidebar).

---

## 💻 CLI Usage (No UI)

For server-side / headless batch processing:

```bash
# Offer Letters
python batch_generate.py \
  --csv sample_data/sample_candidates.csv \
  --template templates/Offer_Letter_Template.docx \
  --type "Offer Letter" \
  --groq-key gsk_xxx \
  --output output/

# Internship Certificates
python batch_generate.py \
  --csv sample_data/sample_candidates.csv \
  --template templates/Internship_Certificate_Template.docx \
  --type "Internship Certificate" \
  --output output/

# LOR
python batch_generate.py \
  --csv sample_data/sample_candidates.csv \
  --template templates/LOR_Template.docx \
  --type "Letter of Recommendation (LOR)" \
  --groq-key gsk_xxx \
  --output output/
```

---

## 📤 Output

- **UI mode**: Download as ZIP or save to local `/output` folder
- **CLI mode**: Files saved directly to `--output` directory
- **Naming format**: `Name - Letter Type.docx`
  - `Rahul Sharma - Offer Letter.docx`
  - `Priya Mehta - Internship Certificate.docx`
  - `Amit Joshi - LOR.docx`

---

## ✅ Design Principles

- **Zero format modification**: `docxtpl` renders placeholders while preserving every font, table, image, header, and footer exactly as in your template
- **Bulk-ready**: Processes 1000+ candidates efficiently
- **AI-optional**: Works fully without Groq; AI fields simply render as empty strings
- **Error resilience**: Failed rows don't crash the batch; errors are logged separately

---

## 🛠️ Tech Stack

| Package | Purpose |
|---------|---------|
| `streamlit` | Web UI |
| `pandas` | CSV processing |
| `python-docx` | Template creation |
| `docxtpl` | Placeholder replacement (format-safe) |
| `groq` | AI content generation |

---

*Innovative Staffing Solutions · Pune, Maharashtra · hr@innovativestaffingsolutions.online*
