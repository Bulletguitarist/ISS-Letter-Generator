# 📄 ISS Letter Generator
**Innovative Staffing Solutions — Bulk Document Automation**

Generate Offer Letters, Internship Certificates, and LORs for 1–1000+ candidates in one click.

🔗 **Live App:** [iss-letter-generator.streamlit.app](https://iss-letter-generator-4ytwxqtfzhkvckqshnfqgr.streamlit.app/)

---

## ✨ Features

- ✅ Bulk generate Offer Letters, Internship Certificates, and LORs from a single CSV
- ✅ Domain-specific content auto-generated (no API key needed)
- ✅ ISS letterhead with logo across all documents
- ✅ Optional Groq AI for smarter, personalized content
- ✅ Download all letters as a ZIP in one click
- ✅ Supports 7 domains: Web Development, Business Development, UI/UX Design, Digital Marketing, Data Analytics, Content Writing, Social Media Management

---

## 🚀 Quick Start (Local)

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/ISS-Letter-Generator.git
cd ISS-Letter-Generator
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Add logo
Place your `logo.jpeg` file in the project root folder.

### 4. Generate templates
```bash
python generate_templates.py
```

### 5. Run the app
```bash
streamlit run app.py
```

Open `http://localhost:8501` in your browser.

---

## 📁 Folder Structure

```
ISS-Letter-Generator/
├── app.py                          ← Main Streamlit web app
├── generate_templates.py           ← Generates DOCX templates
├── batch_generate.py               ← CLI batch processor
├── requirements.txt
├── logo.jpeg                       ← ISS logo (add manually, not in repo)
├── templates/
│   ├── Offer_Letter_Template.docx
│   ├── Internship_Certificate_Template.docx
│   └── LOR_Template.docx
├── sample_data/
│   └── sample_candidates.csv
└── output/                         ← Generated letters saved here
```

---

## 📋 CSV Format

```csv
name,designation,domain,department,joining_date,last_working_date,email,phone,address,duration,performance,skills
Priyanka Sharma,Web Developer Intern,Web Development,Technology,19-11-2025,19-05-2026,priyanka@gmail.com,7483529164,Pune Maharashtra,6 months,Excellent,HTML CSS JavaScript React
```

### Supported Columns

| Column | Used In | Description |
|--------|---------|-------------|
| `name` | All | Full candidate name *(required)* |
| `designation` | All | Job title / role |
| `domain` | All | Skill domain — see supported domains below |
| `department` | Offer Letter | Department name |
| `joining_date` | All | Start date |
| `last_working_date` | Cert / LOR | End date |
| `email` | All | Email address |
| `phone` | Offer Letter | Phone number |
| `address` | Offer Letter | Candidate address |
| `duration` | Certificate | Duration string (e.g. "3 months") |
| `performance` | Cert / LOR | Excellent / Good / Satisfactory |
| `skills` | Cert / LOR | Comma-separated skills list |

---

## 🏢 Supported Domains

Content (responsibilities, certificate body, LOR paragraphs) is automatically generated based on the `domain` column:

| Domain | Auto Content |
|--------|-------------|
| `Web Development` | HTML, CSS, JS, React, API work |
| `Business Development` | Lead gen, CRM, client outreach |
| `UI/UX Design` | Figma, wireframes, prototyping |
| `Digital Marketing` | SEO, social media, campaigns |
| `Data Analytics` | Python, SQL, Power BI, dashboards |
| `Content Writing` | Blogs, SEO writing, copywriting |
| `Social Media Management` | Instagram, LinkedIn, scheduling |

---

## 🤖 Groq AI (Optional)

AI generates smarter, personalized content for each candidate.

1. Get a free API key at [console.groq.com](https://console.groq.com/keys)
2. Paste it in the **Groq API Key** field in the sidebar
3. Check **Enable AI Content Generation**

> Without Groq, domain-specific content is still generated automatically from hardcoded templates.

---

## 💻 CLI Usage

```bash
# Offer Letters
python batch_generate.py \
  --csv sample_data/sample_candidates.csv \
  --template templates/Offer_Letter_Template.docx \
  --type "Offer Letter" \
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
  --output output/
```

---

## ⚠️ Setup Notes

- `logo.jpeg` is **not included** in the repo — add it manually to the project root
- Run `python generate_templates.py` after adding logo to regenerate templates with logo
- `output/` folder is gitignored — generated letters are saved here locally

---

## 🛠️ Tech Stack

| Package | Purpose |
|---------|---------|
| `streamlit` | Web UI |
| `pandas` | CSV processing |
| `python-docx` | Template creation |
| `docxtpl` | Placeholder replacement |
| `groq` | AI content generation (optional) |
| `openpyxl` | Excel support |

---

## 📤 Output

- **UI mode**: Download as ZIP or save to local `/output` folder
- **CLI mode**: Files saved to `--output` directory
- **Naming**: `Candidate Name - Letter Type.docx`
  - `Priyanka Sharma - Offer Letter.docx`
  - `Rahul Mehta - Internship Certificate.docx`
  - `Anjali Singh - LOR.docx`

---
THE BRATS (Ashvini Goswami & Jyotirmoy Mahapatra)
