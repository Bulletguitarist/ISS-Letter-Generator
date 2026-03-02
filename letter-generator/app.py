"""
Innovative Staffing Solutions — Letter Generator
Bulk generate Offer Letters, Internship Certificates, and LORs
from CSV + DOCX templates with optional Groq AI content generation.
"""

import streamlit as st
import pandas as pd
import io
import os
import zipfile
import tempfile
from pathlib import Path
from datetime import datetime

# ─── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Letter Generator — Innovative Staffing Solutions",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a3a5c 0%, #2e6da4 100%);
        padding: 2rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 2rem;
        text-align: center;
    }
    .main-header h1 { margin: 0; font-size: 2rem; }
    .main-header p  { margin: 0.5rem 0 0; opacity: 0.85; }
    .stat-card {
        background: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
    }
    .stat-card .number { font-size: 2rem; font-weight: bold; color: #1a3a5c; }
    .stat-card .label  { font-size: 0.85rem; color: #6c757d; }
    .success-box {
        background: #d4edda; border: 1px solid #c3e6cb;
        border-radius: 8px; padding: 1rem; color: #155724;
    }
    .warning-box {
        background: #fff3cd; border: 1px solid #ffeeba;
        border-radius: 8px; padding: 1rem; color: #856404;
    }
    .info-box {
        background: #cce5ff; border: 1px solid #b8daff;
        border-radius: 8px; padding: 1rem; color: #004085;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #1a3a5c, #2e6da4);
        color: white;
        border: none;
        padding: 0.75rem;
        border-radius: 8px;
        font-size: 1rem;
        font-weight: 600;
        cursor: pointer;
    }
    .stButton>button:hover { opacity: 0.9; }
</style>
""", unsafe_allow_html=True)

# ─── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📄 Letter Generator</h1>
    <p>Innovative Staffing Solutions — Bulk Document Automation</p>
</div>
""", unsafe_allow_html=True)

# ─── Imports ──────────────────────────────────────────────────────────────────
def import_docxtpl():
    try:
        from docxtpl import DocxTemplate
        return DocxTemplate
    except ImportError:
        st.error("❌ `docxtpl` not installed. Run: `pip install docxtpl`")
        st.stop()

def import_groq():
    try:
        from groq import Groq
        return Groq
    except ImportError:
        return None

# ─── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://via.placeholder.com/200x60/1a3a5c/ffffff?text=ISS", use_container_width=True)
    st.markdown("---")
    st.subheader("⚙️ Configuration")

    letter_type = st.selectbox(
        "Letter Type",
        ["Offer Letter", "Internship Certificate", "Letter of Recommendation (LOR)"],
    )

    st.markdown("---")
    st.subheader("🤖 Groq AI (Optional)")
    groq_key = st.text_input(
        "Groq API Key",
        type="password",
        value=os.environ.get("GROQ_API_KEY", ""),
        placeholder="gsk_..."
    )
    groq_model = st.selectbox(
        "Groq Model",
        ["llama3-70b-8192", "llama3-8b-8192", "llama-3.1-8b-instant", "gemma2-9b-it"],
    )
    use_ai = st.checkbox("Enable AI Content Generation", value=bool(groq_key))

    st.markdown("---")
    st.subheader("📁 Output")
    output_mode = st.radio("Save to:", ["Download ZIP", "Local /output folder"])

    st.markdown("---")
    with st.expander("ℹ️ CSV Column Reference"):
        st.markdown("""
| Column | Description |
|--------|-------------|
| `name` | Full candidate name |
| `designation` | Job title / role |
| `domain` | Web Development / Business Development / UI UX Design / Digital Marketing / Data Analytics / Content Writing / Social Media Management |
| `department` | Department name |
| `joining_date` | Start date |
| `last_working_date` | End date |
| `email` | Email address |
| `phone` | Phone number |
| `address` | Candidate address |
| `duration` | Duration (e.g. "3 months") |
| `performance` | Excellent / Good / Satisfactory |
| `skills` | Comma-sep skills list |
        """)

# ─── Domain Content ────────────────────────────────────────────────────────────
DOMAIN_CONTENT = {
    "Web Development": {
        "responsibilities": [
            "Assisting in designing and developing responsive web pages",
            "Working with HTML, CSS, JavaScript, and modern frameworks",
            "Supporting backend integrations, APIs, and database operations",
            "Testing, debugging, and improving application functionality",
            "Contributing to documentation, research, and team discussions",
            "Supporting the development of AI automation and agent workflows",
        ],
        "skills_used": "HTML, CSS, JavaScript, React, API Integration",
        "certificate_body": (
            "demonstrated strong technical skills in web development, contributing to real-time projects "
            "involving frontend and backend development, UI/UX implementation, and AI-based workflow automation."
        ),
        "lor_para1": (
            "During their tenure as a Web Developer Intern, {name} worked on live projects involving "
            "responsive web design, JavaScript frameworks, and backend API integrations."
        ),
        "lor_para2": (
            "{name} demonstrated exceptional problem-solving skills, attention to detail, and the ability "
            "to deliver quality work within deadlines. Their contributions significantly improved our "
            "product's performance and usability."
        ),
        "lor_para3": (
            "We strongly recommend {name} for any web development role. They have the technical foundation, "
            "professional attitude, and growth mindset required to excel in the industry."
        ),
    },
    "Business Development": {
        "responsibilities": [
            "Data scraping and lead database creation",
            "Lead qualification and verification",
            "Outreach via Email, LinkedIn, WhatsApp and Calls",
            "Converting hot leads into confirmed clients",
            "Daily reporting and CRM updates",
            "Supporting client onboarding and coordination",
        ],
        "skills_used": "Lead Generation, CRM, Client Outreach, Communication, Reporting",
        "certificate_body": (
            "demonstrated outstanding business acumen and communication skills, contributing to lead "
            "generation, client outreach, and business growth initiatives at Innovative Staffing Solutions."
        ),
        "lor_para1": (
            "During their tenure as a Business Development Intern, {name} actively contributed to "
            "lead generation, client outreach campaigns, and CRM management, playing a key role in "
            "expanding our client base."
        ),
        "lor_para2": (
            "{name} showed exceptional communication skills, persistence, and a results-driven attitude. "
            "Their ability to identify and convert leads into clients was commendable and added real "
            "business value to our operations."
        ),
        "lor_para3": (
            "We wholeheartedly recommend {name} for any business development or sales role. "
            "Their drive, professionalism, and interpersonal skills make them a valuable asset "
            "to any organization."
        ),
    },
    "UI/UX Design": {
        "responsibilities": [
            "Creating wireframes, mockups, and prototypes using Figma and Adobe XD",
            "Conducting user research and usability testing",
            "Designing intuitive and visually appealing user interfaces",
            "Collaborating with developers for seamless design implementation",
            "Maintaining design consistency across all platforms",
            "Contributing to brand identity and visual communication",
        ],
        "skills_used": "Figma, Adobe XD, Canva, Wireframing, Prototyping, User Research",
        "certificate_body": (
            "showcased exceptional creativity and design thinking, contributing to UI/UX projects "
            "involving wireframing, prototyping, and user interface design for web and mobile platforms."
        ),
        "lor_para1": (
            "During their tenure as a UI/UX Design Intern, {name} worked on designing user interfaces "
            "for web and mobile applications, creating wireframes, and conducting usability testing."
        ),
        "lor_para2": (
            "{name} demonstrated a strong eye for detail, creative thinking, and the ability to "
            "translate complex requirements into clean, intuitive designs. Their work consistently "
            "received positive feedback from stakeholders."
        ),
        "lor_para3": (
            "We highly recommend {name} for any UI/UX or design role. Their creativity, technical "
            "proficiency, and user-centric approach make them an excellent candidate."
        ),
    },
    "Digital Marketing": {
        "responsibilities": [
            "Planning and executing digital marketing campaigns across platforms",
            "Managing social media accounts and content calendars",
            "SEO optimization and keyword research",
            "Creating engaging content for blogs, emails, and social media",
            "Analyzing campaign performance using Google Analytics and Meta Insights",
            "Supporting paid advertising campaigns on Google and Meta platforms",
        ],
        "skills_used": "SEO, Social Media Marketing, Google Analytics, Content Creation, Email Marketing",
        "certificate_body": (
            "demonstrated strong marketing skills and digital acumen, contributing to campaigns "
            "across social media, SEO, and content marketing that improved our online presence."
        ),
        "lor_para1": (
            "During their tenure as a Digital Marketing Intern, {name} actively contributed to "
            "social media management, SEO initiatives, and content creation."
        ),
        "lor_para2": (
            "{name} showed a data-driven approach to marketing, analytical thinking, and creative "
            "content development. Their campaigns delivered measurable improvements in reach and engagement."
        ),
        "lor_para3": (
            "We confidently recommend {name} for any digital marketing role. Their skills in "
            "SEO, social media, and campaign analytics make them a well-rounded marketing professional."
        ),
    },
    "Data Analytics": {
        "responsibilities": [
            "Collecting, cleaning, and processing large datasets",
            "Performing exploratory data analysis and generating insights",
            "Creating dashboards and visual reports using Excel, Power BI, or Tableau",
            "Supporting data-driven decision making across departments",
            "Writing Python/SQL scripts for data extraction and transformation",
            "Documenting analytical processes and findings",
        ],
        "skills_used": "Python, SQL, Excel, Power BI, Tableau, Data Visualization",
        "certificate_body": (
            "demonstrated strong analytical and technical skills, contributing to data collection, "
            "analysis, and visualization projects that supported data-driven decision-making."
        ),
        "lor_para1": (
            "During their tenure as a Data Analytics Intern, {name} worked on data collection, "
            "cleaning, and analysis projects, creating insightful dashboards and reports."
        ),
        "lor_para2": (
            "{name} showed exceptional analytical thinking, attention to detail, and proficiency "
            "in tools like Python, SQL, and Power BI. Their ability to translate raw data into "
            "actionable insights was highly valuable."
        ),
        "lor_para3": (
            "We strongly recommend {name} for any data analytics or business intelligence role. "
            "Their technical skills and analytical mindset make them a strong candidate."
        ),
    },
    "Content Writing": {
        "responsibilities": [
            "Researching and writing high-quality blog posts and articles",
            "Creating SEO-optimized content for websites and landing pages",
            "Writing scripts for video and social media content",
            "Editing and proofreading content for grammar and clarity",
            "Collaborating with design and marketing teams for content campaigns",
            "Maintaining consistent brand voice across all written materials",
        ],
        "skills_used": "Content Writing, SEO Writing, Editing, Research, WordPress, Copywriting",
        "certificate_body": (
            "demonstrated excellent writing and communication skills, contributing to content "
            "creation across blogs, social media, and marketing materials."
        ),
        "lor_para1": (
            "During their tenure as a Content Writing Intern, {name} produced high-quality "
            "written content including blog posts, social media copies, and SEO articles."
        ),
        "lor_para2": (
            "{name} demonstrated a strong command of language, research ability, and the "
            "capacity to adapt their writing style to different audiences and platforms."
        ),
        "lor_para3": (
            "We gladly recommend {name} for any content writing or communications role. "
            "Their writing skills, creativity, and professionalism make them an excellent addition."
        ),
    },
    "Social Media Management": {
        "responsibilities": [
            "Managing and scheduling content across Instagram, LinkedIn, Facebook, and Twitter",
            "Creating engaging posts, reels, and stories aligned with brand guidelines",
            "Monitoring and responding to audience engagement and comments",
            "Analyzing social media metrics and preparing performance reports",
            "Collaborating with design team for visual content creation",
            "Researching trends and competitor strategies for content improvement",
        ],
        "skills_used": "Instagram, LinkedIn, Facebook, Canva, Social Media Analytics, Content Scheduling",
        "certificate_body": (
            "demonstrated strong social media skills and creative thinking, managing content "
            "across multiple platforms and contributing to audience growth and engagement."
        ),
        "lor_para1": (
            "During their tenure as a Social Media Management Intern, {name} managed our "
            "social media presence across platforms, creating engaging content and implementing "
            "growth strategies."
        ),
        "lor_para2": (
            "{name} showed creativity, consistency, and a deep understanding of social media "
            "algorithms and audience behavior. Their content strategies delivered measurable growth."
        ),
        "lor_para3": (
            "We enthusiastically recommend {name} for any social media or digital content role. "
            "Their creativity, platform knowledge, and strategic thinking make them outstanding."
        ),
    },
}

def get_domain_content(domain: str) -> dict:
    for key in DOMAIN_CONTENT:
        if key.lower() in domain.lower() or domain.lower() in key.lower():
            return DOMAIN_CONTENT[key]
    return DOMAIN_CONTENT["Web Development"]

# ─── Helper Functions ──────────────────────────────────────────────────────────
def call_groq(client, model: str, prompt: str) -> str:
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are an expert HR writer for Innovative Staffing Solutions, "
                        "a professional staffing company based in Pune, Maharashtra, India. "
                        "Write formal, professional, concise content. "
                        "Do NOT include greetings or sign-offs."
                    )
                },
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=600,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"[AI generation failed: {e}]"


def build_context(row: dict, letter_type: str, groq_client=None, model: str = "") -> dict:
    name        = str(row.get("name", "")).strip()
    designation = str(row.get("designation", "")).strip()
    domain      = str(row.get("domain", "")).strip()
    department  = str(row.get("department", "Business Development")).strip()
    joining     = str(row.get("joining_date", "")).strip()
    leaving     = str(row.get("last_working_date", "")).strip()
    email       = str(row.get("email", "")).strip()
    phone       = str(row.get("phone", "")).strip()
    address     = str(row.get("address", "")).strip()
    duration    = str(row.get("duration", "")).strip()
    performance = str(row.get("performance", "Excellent")).strip()
    skills      = str(row.get("skills", "")).strip()

    today = datetime.today().strftime("%d %B %Y")
    year  = datetime.today().strftime("%Y")

    dc = get_domain_content(domain)

    ctx = {
        "name":              name,
        "full_name":         name,
        "designation":       designation,
        "domain":            domain,
        "department":        department,
        "joining_date":      joining,
        "last_working_date": leaving,
        "email":             email,
        "phone":             phone,
        "address":           address,
        "duration":          duration,
        "performance":       performance,
        "skills":            skills if skills else dc["skills_used"],
        "today_date":        today,
        "current_year":      year,
        "company_name":      "Innovative Staffing Solutions",
        "company_address":   "Pune, Maharashtra, India",
        "hr_name":           "Kaif Khan",
        "hr_designation":    "HR Manager",
        "founder_name":      "Founder",
        "basic_salary":      str(row.get("basic_salary", "16,000")),
        "hra":               str(row.get("hra", "8,000")),
        "special_allowance": str(row.get("special_allowance", "6,000")),
        "gross_salary":      str(row.get("gross_salary", "30,000")),
        "net_salary":        str(row.get("net_salary", "30,000")),
        "monthly_ctc":       str(row.get("monthly_ctc", "32,000")),
        "annual_ctc":        str(row.get("annual_ctc", "3,84,000")),
        "probation_period":  str(row.get("probation_period", "3 months")),
        # Domain specific
        "domain_skills":           dc["skills_used"],
        "domain_certificate_body": dc["certificate_body"],
        "domain_lor_para1":        dc["lor_para1"].format(name=name),
        "domain_lor_para2":        dc["lor_para2"].format(name=name),
        "domain_lor_para3":        dc["lor_para3"].format(name=name),
        # AI placeholders — filled with domain content by default
        "ai_generated_content":       dc["lor_para1"].format(name=name),
        "ai_internship_description":  dc["certificate_body"],
        "ai_lor_body": (
            dc["lor_para1"].format(name=name) + "\n\n" +
            dc["lor_para2"].format(name=name) + "\n\n" +
            dc["lor_para3"].format(name=name)
        ),
        "ai_performance_summary": f"{name} performed {performance.lower()} throughout the internship.",
        "ai_skills_summary":      f"{name} demonstrated proficiency in {dc['skills_used']}.",
    }

    # Override with Groq if available
    if groq_client:
        if letter_type == "Offer Letter":
            ctx["ai_generated_content"] = call_groq(
                groq_client, model,
                f"Write a warm, professional welcome paragraph (2-3 sentences) for an offer letter "
                f"addressed to {name} joining as {designation} in {domain} at "
                f"Innovative Staffing Solutions, Pune."
            )
        elif letter_type == "Internship Certificate":
            ctx["ai_internship_description"] = call_groq(
                groq_client, model,
                f"Write 2-3 sentence internship certificate body for {name} who completed "
                f"{domain} internship at Innovative Staffing Solutions from {joining} to {leaving}. "
                f"Performance: {performance}. Skills: {dc['skills_used']}. Formal, third person."
            )
            ctx["ai_performance_summary"] = call_groq(
                groq_client, model,
                f"One sentence about {name}'s {performance} performance in {domain} internship."
            )
            ctx["ai_skills_summary"] = call_groq(
                groq_client, model,
                f"One sentence about skills demonstrated by {name} in {domain}. Skills: {dc['skills_used']}."
            )
            ctx["ai_generated_content"] = ctx["ai_internship_description"]
        elif letter_type == "Letter of Recommendation (LOR)":
            ctx["ai_lor_body"] = call_groq(
                groq_client, model,
                f"Write 3-paragraph LOR for {name} as {designation} in {domain} "
                f"at Innovative Staffing Solutions from {joining} to {leaving}. "
                f"Performance: {performance}. Skills: {dc['skills_used']}. "
                f"Para 1: Role. Para 2: Achievements. Para 3: Recommendation. Formal, third person."
            )
            ctx["ai_generated_content"] = ctx["ai_lor_body"]

    return ctx


def replace_placeholders_in_docx(template_bytes: bytes, context: dict) -> bytes:
    try:
        from docxtpl import DocxTemplate
    except ImportError:
        raise ImportError("`docxtpl` not installed. Run: pip install docxtpl")

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_in:
        tmp_in.write(template_bytes)
        tmp_in_path = tmp_in.name

    try:
        tpl = DocxTemplate(tmp_in_path)
        tpl.render(context)
        buf = io.BytesIO()
        tpl.save(buf)
        buf.seek(0)
        return buf.read()
    finally:
        os.unlink(tmp_in_path)


def get_output_filename(name: str, letter_type: str) -> str:
    safe_name  = name.replace("/", "-").replace("\\", "-").strip()
    short_type = letter_type.replace(" (LOR)", "").replace("Letter of Recommendation", "LOR")
    return f"{safe_name} - {short_type}.docx"


def process_all(df, template_bytes, letter_type, groq_client=None, model="", progress_bar=None, status_text=None):
    results = {}
    total   = len(df)
    for i, (_, row) in enumerate(df.iterrows()):
        name = str(row.get("name", f"Candidate_{i+1}")).strip()
        if status_text:
            status_text.text(f"Generating: {name} ({i+1}/{total})")
        try:
            ctx       = build_context(row.to_dict(), letter_type, groq_client, model)
            doc_bytes = replace_placeholders_in_docx(template_bytes, ctx)
            filename  = get_output_filename(name, letter_type)
            results[filename] = doc_bytes
        except Exception as e:
            results[f"ERROR_{name}.txt"] = f"Failed: {e}".encode()
        if progress_bar:
            progress_bar.progress((i + 1) / total)
    return results


def save_to_local(results: dict) -> int:
    out_dir = Path("output")
    out_dir.mkdir(exist_ok=True)
    count = 0
    for filename, data in results.items():
        try:
            (out_dir / filename).write_bytes(data)
            count += 1
        except Exception:
            pass
    return count


def create_zip(results: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, data in results.items():
            zf.writestr(filename, data)
    buf.seek(0)
    return buf.read()


# ─── Main UI ───────────────────────────────────────────────────────────────────
col_upload, col_preview = st.columns([1, 1], gap="large")

with col_upload:
    st.subheader("📂 Upload Files")
    csv_file  = st.file_uploader("1️⃣  CSV File (candidates)", type=["csv"])
    tmpl_file = st.file_uploader("2️⃣  DOCX Template", type=["docx"])
    st.markdown("""
    <div class="info-box">
    <strong>💡 Placeholder syntax:</strong> Use <code>{{name}}</code>, <code>{{designation}}</code>,
    <code>{{joining_date}}</code>, <code>{{ai_lor_body}}</code>, etc. in your DOCX template.
    </div>
    """, unsafe_allow_html=True)

with col_preview:
    st.subheader("👁️ Preview")
    if csv_file:
        try:
            df = pd.read_csv(csv_file)
            csv_file.seek(0)
            st.success(f"✅ CSV loaded: **{len(df)} candidates**, **{len(df.columns)} columns**")
            c1, c2, c3 = st.columns(3)
            c1.markdown(f'<div class="stat-card"><div class="number">{len(df)}</div><div class="label">Candidates</div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="stat-card"><div class="number">{len(df.columns)}</div><div class="label">Columns</div></div>', unsafe_allow_html=True)
            c3.markdown(f'<div class="stat-card"><div class="number">1</div><div class="label">Template</div></div>', unsafe_allow_html=True)
            st.dataframe(df.head(5), use_container_width=True, height=200)
            recommended = ["name", "designation", "domain", "joining_date", "email"]
            missing = [c for c in recommended if c not in df.columns]
            if missing:
                st.warning(f"⚠️ Recommended columns missing: `{'`, `'.join(missing)}`")
        except Exception as e:
            st.error(f"Could not read CSV: {e}")
    else:
        st.markdown('<div class="warning-box">📋 Upload your CSV file to see a preview here.</div>', unsafe_allow_html=True)

# ─── Generate Button ───────────────────────────────────────────────────────────
st.markdown("---")
generate_col, _ = st.columns([1, 2])
with generate_col:
    generate_btn = st.button("🚀 Generate All Letters", use_container_width=True)

if generate_btn:
    errors = []
    if not csv_file:  errors.append("No CSV file uploaded.")
    if not tmpl_file: errors.append("No DOCX template uploaded.")
    if use_ai and not groq_key: errors.append("Groq API key required when AI is enabled.")

    if errors:
        for e in errors:
            st.error(f"❌ {e}")
    else:
        csv_file.seek(0)
        df = pd.read_csv(csv_file)
        template_bytes = tmpl_file.read()

        groq_client = None
        if use_ai and groq_key:
            Groq = import_groq()
            if Groq is None:
                st.error("❌ `groq` not installed. Run: `pip install groq`")
                st.stop()
            try:
                groq_client = Groq(api_key=groq_key)
                groq_client.models.list()
                st.success("✅ Groq API connected.")
            except Exception as e:
                st.error(f"❌ Groq API error: {e}")
                st.stop()

        st.markdown("### ⏳ Generating Documents…")
        progress_bar = st.progress(0)
        status_text  = st.empty()
        start_time   = datetime.now()

        results = process_all(
            df, template_bytes, letter_type,
            groq_client=groq_client, model=groq_model,
            progress_bar=progress_bar, status_text=status_text,
        )

        elapsed   = (datetime.now() - start_time).total_seconds()
        status_text.empty()
        progress_bar.progress(1.0)

        successes = {k: v for k, v in results.items() if not k.startswith("ERROR_")}
        failures  = {k: v for k, v in results.items() if k.startswith("ERROR_")}

        st.markdown(f"""
        <div class="success-box">
        ✅ <strong>Done!</strong> — {len(successes)} letters in {elapsed:.1f}s
        {f"<br>⚠️ {len(failures)} failed" if failures else ""}
        </div>
        """, unsafe_allow_html=True)

        if failures:
            with st.expander(f"⚠️ {len(failures)} errors"):
                for fname, err in failures.items():
                    st.error(f"**{fname}**: {err.decode()}")

        if output_mode == "Download ZIP" and successes:
            zip_bytes = create_zip(successes)
            zip_name  = f"ISS_{letter_type.replace(' ', '_')}_{datetime.today().strftime('%Y%m%d_%H%M')}.zip"
            st.download_button(
                label=f"⬇️ Download ZIP ({len(successes)} files)",
                data=zip_bytes,
                file_name=zip_name,
                mime="application/zip",
                use_container_width=True,
            )
        elif output_mode == "Local /output folder" and successes:
            saved = save_to_local(successes)
            st.success(f"💾 {saved} files saved to `./output/`")

        if len(successes) <= 20:
            with st.expander("📄 Individual Downloads"):
                for filename, doc_bytes in successes.items():
                    st.download_button(
                        label=f"⬇️ {filename}",
                        data=doc_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=filename,
                    )

# ─── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<center style='color:#6c757d; font-size:0.85rem;'>"
    "Innovative Staffing Solutions · Pune, Maharashtra · "
    "hr@innovativestaffingsolutions.online · +91 7447802076"
    "</center>",
    unsafe_allow_html=True,
)