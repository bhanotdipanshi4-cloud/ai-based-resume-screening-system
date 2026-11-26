from __future__ import annotations

import base64
import html
import hashlib
import io
import re
from datetime import datetime
from typing import Any

import pandas as pd
import plotly.express as px
import streamlit as st
from docx import Document
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

st.set_page_config(
    page_title="Talent Lens | Resume Screening",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="expanded",
)

SKILL_ALIASES: dict[str, tuple[str, ...]] = {
    "Python": ("python",),
    "SQL": ("sql", "structured query language"),
    "Power BI": ("power bi", "powerbi"),
    "Tableau": ("tableau",),
    "Excel": ("excel", "microsoft excel"),
    "Machine Learning": ("machine learning", "ml"),
    "Deep Learning": ("deep learning",),
    "Data Analysis": ("data analysis", "data analytics", "analytics"),
    "Statistics": ("statistics", "statistical analysis"),
    "Data Visualization": ("data visualization", "data visualisation"),
    "Pandas": ("pandas",),
    "NumPy": ("numpy",),
    "R": ("r programming", "r language"),
    "Java": ("java",),
    "JavaScript": ("javascript",),
    "TypeScript": ("typescript",),
    "C++": ("c++", "cpp"),
    "C#": ("c#", "c sharp"),
    "HTML": ("html",),
    "CSS": ("css",),
    "React": ("react",),
    "Node.js": ("node.js", "nodejs"),
    "AWS": ("aws", "amazon web services"),
    "Azure": ("azure",),
    "Google Cloud": ("google cloud", "gcp"),
    "Docker": ("docker",),
    "Kubernetes": ("kubernetes", "k8s"),
    "Git": ("git",),
    "Linux": ("linux",),
    "Communication": ("communication", "communicator"),
    "Project Management": ("project management",),
    "Agile": ("agile", "scrum"),
    "Leadership": ("leadership", "team lead"),
    "Figma": ("figma",),
    "Product Management": ("product management",),
    "Cybersecurity": ("cybersecurity", "information security"),
    "ETL": ("etl", "extract transform load"),
    "Machine Vision": ("computer vision",),
    "NLP": ("natural language processing", "nlp"),
    "TensorFlow": ("tensorflow",),
    "PyTorch": ("pytorch",),
    "Spark": ("apache spark", "pyspark", "spark"),
}
DEPARTMENTS = [
    "Unspecified", "Engineering", "Data & Analytics", "Product", "Design",
    "Marketing", "Sales", "Operations", "Finance", "People & HR", "Other",
]
EDUCATION_TERMS = [
    "PhD", "Doctorate", "Master's", "Bachelor's", "Associate degree",
    "MBA", "Diploma", "High school",
]
STANDARD_RESUME_SECTIONS = (
    "experience", "employment", "education", "skills", "projects", "summary",
    "profile", "certifications", "technical skills", "work history",
)
ACTION_VERBS = (
    "built", "created", "delivered", "developed", "designed", "improved",
    "increased", "implemented", "launched", "led", "managed", "reduced",
    "automated", "analyzed", "analysed",
)


def _pattern_for(alias: str) -> re.Pattern[str]:
    if "\\b" in alias or "\\B" in alias:
        return re.compile(alias, re.IGNORECASE)
    escaped = re.escape(alias)
    return re.compile(r"(?<![\w+#.])" + escaped + r"(?![\w+#.])", re.IGNORECASE)


def find_skills(text: str) -> list[str]:
    found = []
    for skill, aliases in SKILL_ALIASES.items():
        if any(_pattern_for(alias).search(text) for alias in aliases):
            found.append(skill)
    return found


def extract_text(filename: str, data: bytes) -> str:
    suffix = filename.lower().rsplit(".", 1)[-1]
    if suffix == "pdf":
        reader = PdfReader(io.BytesIO(data))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    if suffix == "docx":
        document = Document(io.BytesIO(data))
        paragraphs = [paragraph.text for paragraph in document.paragraphs if paragraph.text.strip()]
        for table in document.tables:
            paragraphs.extend(cell.text for row in table.rows for cell in row.cells if cell.text.strip())
        return "\n".join(paragraphs)
    raise ValueError("Please upload a PDF or DOCX file.")


def extract_name(text: str, filename: str, email: str) -> str:
    for line in text.splitlines()[:12]:
        candidate = re.sub(r"\s+", " ", line).strip(" |\t•-_")
        if not candidate or len(candidate) > 70:
            continue
        if re.search(r"@|https?://|www\.|\d{3,}|resume|curriculum vitae|phone|email", candidate, re.I):
            continue
        words = candidate.split()
        if 2 <= len(words) <= 5 and all(re.search(r"[A-Za-z]", word) for word in words):
            return candidate.title() if candidate.isupper() else candidate
    if email:
        return email.split("@", 1)[0].replace(".", " ").replace("_", " ").title()
    return filename.rsplit(".", 1)[0].replace("_", " ").replace("-", " ").title()


def extract_candidate(filename: str, data: bytes) -> dict[str, Any]:
    text = extract_text(filename, data)
    if not text.strip():
        raise ValueError("No selectable text found. Scanned/image-only resumes need OCR before upload.")
    email_match = re.search(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", text, re.I)
    email = email_match.group(0) if email_match else "Not found"
    experience_matches = re.findall(
        r"\b(\d{1,2})\s*\+?\s*(?:years?|yrs?)\s+(?:of\s+)?(?:professional\s+)?experience\b",
        text,
        re.I,
    )
    experience = max((int(value) for value in experience_matches), default=None)
    education_terms = [term for term in EDUCATION_TERMS if re.search(re.escape(term), text, re.I)]
    education = education_terms[0] if education_terms else "Not found"
    digest = hashlib.sha256(data).hexdigest()
    return {
        "id": digest[:12], "filename": filename, "data": data,
        "name": extract_name(text, filename, email if email != "Not found" else ""),
        "email": email, "text": text, "skills": find_skills(text),
        "experience": experience, "education": education, "education_terms": education_terms,
        "department": "Unspecified", "status": "Pending review",
        "uploaded_at": datetime.now().strftime("%b %d, %H:%M"),
    }


def extract_required_skills(job_description: str, custom_skills: str) -> list[str]:
    required = find_skills(job_description)
    for item in re.split(r"[,;\n]", custom_skills):
        skill = item.strip()
        if skill and not any(skill.casefold() == known.casefold() for known in required):
            required.append(skill)
    return required


def score_candidate(candidate: dict[str, Any], required: list[str]) -> tuple[int, list[str], list[str]]:
    if not required:
        return 0, [], []
    resume_skills = {skill.casefold() for skill in candidate["skills"]}
    normalized_text = re.sub(r"[^a-z0-9+#]+", " ", candidate["text"].casefold())
    matched = []
    for skill in required:
        normalized_skill = re.sub(r"[^a-z0-9+#]+", " ", skill.casefold()).strip()
        if skill.casefold() in resume_skills or (
            normalized_skill and f" {normalized_skill} " in f" {normalized_text} "
        ):
            matched.append(skill)
    missing = [skill for skill in required if skill not in matched]
    score = round(100 * len(matched) / len(required))
    return score, matched, missing


def recommendation(score: int, required: list[str]) -> str:
    if not required:
        return "Add job skills to assess"
    if score >= 90:
        return "Strong alignment · review profile"
    if score >= 75:
        return "Good alignment · review profile"
    if score >= 50:
        return "Partial alignment · manual review"
    return "Limited keyword overlap · manual review"


def make_summary(candidate: dict[str, Any]) -> str:
    parts = []
    skills = candidate["skills"]
    if skills:
        parts.append("Resume text mentions " + ", ".join(skills[:6]) + ".")
    else:
        parts.append("No skills from the built-in vocabulary were detected; review the resume text.")
    if candidate["experience"] is not None:
        parts.append(f"It states {candidate['experience']} years of experience.")
    if candidate["education"] != "Not found":
        parts.append(f"Education text includes {candidate['education']}.")
    return " ".join(parts)


def required_experience_years(job_description: str) -> int | None:
    patterns = (
        r"\b(\d{1,2})\s*\+?\s*(?:years?|yrs?)\s+(?:of\s+)?(?:relevant\s+|professional\s+)?experience\b",
        r"\bminimum\s+(\d{1,2})\s*(?:years?|yrs?)\b",
    )
    matches = [int(value) for pattern in patterns for value in re.findall(pattern, job_description, re.I)]
    return max(matches) if matches else None


def ats_check(candidate: dict[str, Any]) -> tuple[int, list[str]]:
    text = candidate["text"]
    words = re.findall(r"\b[\w+#.-]+\b", text)
    sections = [section for section in STANDARD_RESUME_SECTIONS if re.search(rf"\b{re.escape(section)}\b", text, re.I)]
    issues: list[str] = []
    score = 100
    if len(words) < 120:
        issues.append("Extracted text is short; check whether the resume is complete or image-based.")
        score -= 20
    if candidate["email"] == "Not found":
        issues.append("No email address was detected in the extracted text.")
        score -= 15
    if not re.search(r"(?:\+?\d[\d ()-]{7,}\d)", text):
        issues.append("No phone number was detected in the extracted text.")
        score -= 10
    if len(sections) < 2:
        issues.append("Few standard resume section labels were detected; use clear headings such as Experience and Education.")
        score -= 20
    if not re.search(r"(?:^|\n)\s*(?:[-*•▪]|\d+[.)])\s+", text):
        issues.append("No text-based bullet markers detected; use simple bullets for accomplishments.")
        score -= 10
    if len(words) > 1600:
        issues.append("Resume text is lengthy; consider tightening repetitive details.")
        score -= 10
    if not issues:
        issues.append("Basic text, contact, section, and list checks passed. Verify the original document layout manually.")
    return max(0, score), issues


def improvement_suggestions(candidate: dict[str, Any], missing_skills: list[str]) -> list[str]:
    text = candidate["text"]
    suggestions: list[str] = []
    if not re.search(r"\b(summary|professional summary|profile|objective)\b", text, re.I):
        suggestions.append("Consider adding a concise professional summary tailored to the target role.")
    if not re.search(r"\b\d+(?:\.\d+)?%?\b", text):
        suggestions.append("Where accurate, add measurable outcomes (such as time saved, revenue, or project scale) to experience bullets.")
    if not any(re.search(rf"\b{verb}\w*\b", text, re.I) for verb in ACTION_VERBS):
        suggestions.append("Start experience bullets with clear action verbs and describe your own contribution.")
    if missing_skills:
        suggestions.append("If you genuinely have them, add evidence for these requested skills: " + ", ".join(missing_skills[:6]) + ".")
    if candidate["experience"] is None:
        suggestions.append("Make employment dates and total relevant experience easy to identify, if applicable.")
    if candidate["education"] == "Not found":
        suggestions.append("Add education or relevant training details if they are useful for this role.")
    if len(suggestions) < 3:
        suggestions.append("Use consistent dates and simple section headings, then proofread the final document.")
    return suggestions


def candidate_view(candidate: dict[str, Any], required: list[str], job_description: str = "") -> dict[str, Any]:
    score, matched, missing = score_candidate(candidate, required)
    required_years = required_experience_years(job_description)
    experience_match = None
    if required_years is not None:
        years = candidate["experience"]
        experience_match = min(100, round(100 * years / required_years)) if years is not None and required_years else 0
    required_education = [term for term in EDUCATION_TERMS if re.search(re.escape(term), job_description, re.I)]
    education_match = None
    if required_education:
        education_match = 100 if any(
            any(term.casefold() in found.casefold() or found.casefold() in term.casefold() for found in candidate.get("education_terms", [candidate["education"]]))
            for term in required_education
        ) else 0
    components: list[tuple[int, int]] = []
    if required:
        components.append((score, 70))
    if experience_match is not None:
        components.append((experience_match, 20))
    if education_match is not None:
        components.append((education_match, 10))
    resume_score = round(sum(value * weight for value, weight in components) / sum(weight for _, weight in components)) if components else 0
    ats_score, ats_issues = ats_check(candidate)
    result = {
        **candidate, "score": score, "resume_score": resume_score, "matched": matched,
        "missing": missing, "experience_match": experience_match,
        "required_experience": required_years, "education_match": education_match,
        "required_education": required_education, "ats_score": ats_score,
        "ats_issues": ats_issues,
    }
    result["suggestions"] = improvement_suggestions(result, missing)
    return result


def create_candidate_report_pdf(candidates: list[dict[str, Any]], role: str) -> bytes:
    buffer = io.BytesIO()
    document = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=18 * mm, leftMargin=18 * mm, topMargin=18 * mm, bottomMargin=18 * mm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="SmallBody", parent=styles["BodyText"], fontName="Helvetica", fontSize=8.5, leading=12, alignment=TA_LEFT, spaceAfter=3))
    story = [Paragraph("Talent Lens - Resume Analysis", styles["Title"]), Spacer(1, 4 * mm)]
    story.append(Paragraph(f"Role: {html.escape(role or 'Not specified')} | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles["SmallBody"]))
    story.append(Paragraph("Advisory text analysis only. Scores are not measures of candidate quality and must not make employment decisions.", styles["SmallBody"]))
    story.append(Spacer(1, 5 * mm))
    for rank, item in enumerate(sorted(candidates, key=lambda value: (-value["resume_score"], value["name"].casefold())), start=1):
        story.append(Paragraph(f"{rank}. {html.escape(item['name'])}", styles["Heading2"]))
        details = [
            ["Email", html.escape(item["email"]), "Review status", html.escape(item["status"])],
            ["Resume score", f"{item['resume_score']}%", "ATS checklist", f"{item['ats_score']}%"],
            ["Skills match", f"{item['score']}%", "Experience", f"{item['experience']} years" if item["experience"] is not None else "Not detected"],
            ["Education detected", html.escape(item["education"]), "Matched skills", html.escape(", ".join(item["matched"]) or "None detected")],
        ]
        table = Table(details, colWidths=[30 * mm, 55 * mm, 30 * mm, 55 * mm])
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#EEF0FA")),
            ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#EEF0FA")),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D9DEEA")),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.extend([table, Spacer(1, 2 * mm)])
        story.append(Paragraph(f"Missing role skills: {html.escape(', '.join(item['missing']) or 'None detected')}", styles["SmallBody"]))
        story.append(Paragraph("Improvement ideas (verify before applying):", styles["SmallBody"]))
        for suggestion in item["suggestions"]:
            story.append(Paragraph(f"- {html.escape(suggestion)}", styles["SmallBody"]))
        story.extend([Spacer(1, 4 * mm)])
    document.build(story)
    return buffer.getvalue()


def score_color(score: int) -> str:
    if score >= 90:
        return "#16a36a"
    if score >= 75:
        return "#3878f6"
    if score >= 50:
        return "#d79b17"
    return "#e45b5b"


def metric_card(label: str, value: str, note: str, icon: str) -> None:
    st.markdown(
        f'<div class="metric-card"><div class="metric-top"><span>{label}</span>'
        f'<span class="metric-icon">{icon}</span></div><div class="metric-value">{value}</div>'
        f'<div class="metric-note">{note}</div></div>',
        unsafe_allow_html=True,
    )


if "candidates" not in st.session_state:
    st.session_state.candidates = []
if "activity" not in st.session_state:
    st.session_state.activity = []

st.markdown(
    """<style>
    :root { --ink:#17243a; --muted:#758198; --line:#e7ebf2; --surface:#fff; --canvas:#f5f7fb; --accent:#6558e8; }
    .stApp { background:var(--canvas); font-family:'DM Sans',sans-serif; }
    [data-testid="stSidebar"] { background:linear-gradient(180deg,#171b34 0%,#202747 100%); }
    [data-testid="stSidebar"] * { color:#f2f4ff; }
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"] p { color:#d5d9ea; }
    [data-testid="stSidebar"] input,[data-testid="stSidebar"] textarea,[data-testid="stSidebar"] [data-baseweb="select"] > div { color:#17243a; }
    .block-container { max-width:1420px; padding-top:2rem; padding-bottom:3rem; }
    h1,h2,h3 { font-family:'Manrope',sans-serif; letter-spacing:-.035em; color:var(--ink); }
    .hero { border-radius:24px; padding:30px 34px; margin-bottom:22px; color:white;
      background:radial-gradient(circle at 90% 20%,rgba(151,139,255,.55),transparent 30%),linear-gradient(115deg,#252c53,#5146a9 67%,#7768ee); box-shadow:0 16px 38px rgba(49,43,117,.18); }
    .hero h1 { color:white; font-size:2rem; margin:4px 0 8px; }
    .hero p { color:#e4e6ff; margin:0; max-width:720px; }
    .eyebrow { color:#c8c8ff; text-transform:uppercase; font-weight:700; font-size:.72rem; letter-spacing:.14em; }
    .metric-card { background:var(--surface); border:1px solid var(--line); border-radius:18px; padding:18px 20px; min-height:130px; box-shadow:0 5px 16px rgba(31,43,73,.035); }
    .metric-top { display:flex; justify-content:space-between; color:#727d92; font-size:.85rem; font-weight:600; }
    .metric-icon { width:32px; height:32px; display:grid; place-items:center; background:#f0efff; border-radius:10px; color:#6459dd; font-size:1.05rem; }
    .metric-value { color:var(--ink); font-family:'Manrope',sans-serif; font-size:1.85rem; font-weight:800; margin:8px 0 1px; }
    .metric-note { color:#96a0b0; font-size:.76rem; }
    .section-label { color:#7f899e; font-weight:700; text-transform:uppercase; letter-spacing:.12em; font-size:.72rem; margin:20px 0 8px; }
    .candidate-card { border:1px solid var(--line); border-radius:18px; padding:18px 20px; background:var(--surface); margin:10px 0; box-shadow:0 5px 17px rgba(31,43,73,.035); }
    .score-track { height:9px; border-radius:10px; background:#edf0f6; overflow:hidden; margin:9px 0 5px; }
    .score-fill { height:100%; border-radius:10px; animation:grow .9s ease-out both; }
    @keyframes grow { from { width:0; } }
    .skill-chip { display:inline-block; padding:5px 10px; border-radius:99px; font-size:.75rem; font-weight:600; margin:3px 4px 3px 0; background:#e9f7f0; color:#168253; }
    .missing-chip { display:inline-block; padding:5px 10px; border-radius:99px; font-size:.75rem; font-weight:600; margin:3px 4px 3px 0; background:#fff3e7; color:#a66419; }
    .subtle { color:#818ba0; font-size:.84rem; }
    div[data-testid="stTabs"] button { font-weight:700; }
    div[data-testid="stButton"] button { border-radius:10px; font-weight:600; }
    .stProgress > div > div > div > div { background:#6558e8; }
    @media (prefers-color-scheme: dark) { .stApp { --ink:#edf0ff; --muted:#a6afc2; --line:#343b54; --surface:#20263a; --canvas:#151a2a; } }
    </style>""",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("## ✦ Talent Lens")
    st.caption("Resume review workspace")
    appearance = st.radio("Appearance", ["Light", "Dark"], horizontal=True, key="appearance")
    st.markdown("---")
    st.markdown("### Job criteria")
    job_title = st.text_input("Role title", placeholder="e.g. Data Analyst")
    department_filter_value = st.selectbox("Department for this role", DEPARTMENTS)
    job_description = st.text_area(
        "Job description", height=190,
        placeholder="Paste the role description. Detected skills are shown for your review.",
        help="Text is processed locally in this app session and is not sent to an AI service.",
    )
    custom_skills = st.text_input("Add or correct required skills", placeholder="Power BI, Tableau, Python")
    required_skills = extract_required_skills(job_description, custom_skills)
    st.markdown("**Detected criteria**")
    if required_skills:
        st.caption(" · ".join(required_skills))
    else:
        st.caption("Add a job description or skills to begin matching.")
    st.markdown("---")
    st.markdown("### Narrow the list")
    name_filter = st.text_input("Candidate name", placeholder="Search names")
    skill_filter = st.text_input("Resume skill", placeholder="e.g. SQL")
    exp_filter = st.selectbox("Experience", ["Any", "0–2 years", "3–5 years", "6+ years", "Not found"])
    education_filter = st.selectbox("Education", ["Any", *EDUCATION_TERMS, "Not found"])
    minimum_score = st.slider("Minimum resume score", 0, 100, 0, 5)
    status_filter = st.selectbox("Review status", ["All", "Pending review", "Shortlisted", "Rejected"])
    st.markdown("---")
    st.caption("🔒 Session-only demo · no database or external AI service")

if appearance == "Dark":
    st.markdown(
        """<style>
        .stApp { --ink:#edf0ff; --muted:#a6afc2; --line:#343b54; --surface:#20263a; --canvas:#151a2a; }
        .metric-card,.candidate-card { background:#20263a; border-color:#343b54; }
        .metric-value,.candidate-card div[style*="font-weight:800"] { color:#edf0ff !important; }
        .metric-icon { background:#343050; }
        .score-track { background:#343b54; }
        .stTabs [data-baseweb="tab-list"] { background:#20263a; }
        div[data-testid="stDataFrame"] { filter:brightness(.88); }
        </style>""",
        unsafe_allow_html=True,
    )

st.markdown(
    '<div class="hero"><div class="eyebrow">Responsible hiring workspace</div>'
    '<h1>Find the skills. Keep the human judgment.</h1>'
    '<p>Review resumes against role-specific criteria with transparent, editable keyword matching. Scores support review; they never make hiring decisions.</p></div>',
    unsafe_allow_html=True,
)

with st.expander("① Add a role description and upload resumes", expanded=not st.session_state.candidates):
    left, right = st.columns([1.05, 0.95], gap="large")
    with left:
        st.markdown("#### Build a role profile")
        if job_description:
            st.info(f"{len(required_skills)} required skill(s) detected. Check the criteria in the sidebar and add any missed skills.")
        else:
            st.markdown('<p class="subtle">Use the sidebar to paste the job description and set the role title.</p>', unsafe_allow_html=True)
        st.markdown("#### Upload a batch")
        uploads = st.file_uploader(
            "Choose resume files", type=["pdf", "docx"], accept_multiple_files=True,
            help="Multiple PDF and DOCX files are supported. Scanned PDFs require OCR before upload.",
        )
        if st.button("Analyze new resumes", type="primary", disabled=not uploads, use_container_width=True):
            known = {candidate["id"] for candidate in st.session_state.candidates}
            added = 0
            errors = []
            progress = st.progress(0, text="Reading resumes…")
            for index, uploaded in enumerate(uploads or []):
                try:
                    candidate = extract_candidate(uploaded.name, uploaded.getvalue())
                    if candidate["id"] not in known:
                        if department_filter_value != "Unspecified":
                            candidate["department"] = department_filter_value
                        st.session_state.candidates.append(candidate)
                        st.session_state.activity.insert(0, f"Added {candidate['name']} · {candidate['uploaded_at']}")
                        known.add(candidate["id"])
                        added += 1
                except Exception as exc:
                    errors.append(f"{uploaded.name}: {exc}")
                progress.progress((index + 1) / len(uploads), text=f"Processed {index + 1} of {len(uploads)}")
            progress.empty()
            if added:
                st.success(f"Added {added} resume(s). Results are ready for human review.")
            if errors:
                for error in errors:
                    st.warning(error)
            if added == 0 and not errors:
                st.info("These resumes are already in the current session.")
    with right:
        st.markdown("#### What the score means")
        st.markdown(
            "The match is the share of **your selected required skills** found as text in a resume. "
            "It does not measure overall ability, verify claims, or infer skills from context. "
            "Edit criteria before using results. No resume is automatically accepted or rejected."
        )
        st.markdown("#### Local, transparent processing")
        st.markdown("PDF/DOCX text extraction and alias-based skill matching run in this app. No external AI API or persistent storage is used. Close the session to clear its resumes.")
        st.caption("This demo does not perform OCR. Review original documents and follow your organization's privacy and hiring policies.")

candidates = st.session_state.candidates
views = [candidate_view(candidate, required_skills, job_description) for candidate in candidates]

# Apply filters consistently across the table and review list.
filtered = []
for item in views:
    if name_filter and name_filter.casefold() not in item["name"].casefold():
        continue
    if skill_filter and not any(skill_filter.casefold() in skill.casefold() for skill in item["skills"]):
        continue
    if item["resume_score"] < minimum_score:
        continue
    if status_filter != "All" and item["status"] != status_filter:
        continue
    years = item["experience"]
    if exp_filter == "Not found" and years is not None:
        continue
    if exp_filter == "0–2 years" and (years is None or years > 2):
        continue
    if exp_filter == "3–5 years" and (years is None or not 3 <= years <= 5):
        continue
    if exp_filter == "6+ years" and (years is None or years < 6):
        continue
    if education_filter != "Any" and item["education"].casefold() != education_filter.casefold():
        continue
    filtered.append(item)
filtered.sort(key=lambda item: (-item["resume_score"], item["name"].casefold()))

st.markdown('<div class="section-label">Workspace overview</div>', unsafe_allow_html=True)
shortlisted = sum(item["status"] == "Shortlisted" for item in views)
rejected = sum(item["status"] == "Rejected" for item in views)
average = round(sum(item["resume_score"] for item in views) / len(views)) if views else 0
average_ats = round(sum(item["ats_score"] for item in views) / len(views)) if views else 0
metrics = st.columns(5, gap="medium")
with metrics[0]: metric_card("Resumes uploaded", str(len(views)), "In this session", "▤")
with metrics[1]: metric_card("Shortlisted", str(shortlisted), "Set by a human reviewer", "✓")
with metrics[2]: metric_card("Average resume score", f"{average}%", "Transparent role criteria", "◉")
with metrics[3]: metric_card("Average ATS check", f"{average_ats}%", "Text-based checklist only", "⌕")
with metrics[4]: metric_card("Rejected", str(rejected), "Set by a human reviewer", "×")

if not candidates:
    st.markdown("<br>", unsafe_allow_html=True)
    st.info("Start by adding a job description in the sidebar, then upload one or more PDF/DOCX resumes.")
else:
    if not required_skills:
        st.warning("No role skills are selected, so match percentages are 0%. Add or correct the job criteria in the sidebar before comparing candidates.")
    tab_dashboard, tab_candidates, tab_compare, tab_activity = st.tabs(
        ["Overview", "Candidate review", "Compare candidates", "Activity & export"]
    )
    with tab_dashboard:
        chart_left, chart_right = st.columns([1.15, 0.85], gap="large")
        with chart_left:
            st.markdown("#### Match score overview")
            if views:
                score_df = pd.DataFrame([{"Candidate": item["name"], "Resume score": item["resume_score"], "Status": item["status"]} for item in filtered])
                if score_df.empty:
                    st.caption("No candidates match the current filters.")
                else:
                    fig = px.bar(score_df, x="Resume score", y="Candidate", color="Resume score", color_continuous_scale=["#ed7777", "#e8bc5c", "#7770e9"], range_x=[0, 100], text="Resume score", orientation="h", hover_data=["Status"])
                    fig.update_traces(texttemplate="%{text}%", textposition="outside", marker_line_width=0)
                    fig.update_layout(height=max(300, len(score_df) * 48), margin=dict(l=0, r=32, t=10, b=10), coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="Overall score (%)", yaxis_title="", font=dict(color="#758198"))
                    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})
        with chart_right:
            st.markdown("#### Reviewer decisions")
            status_df = pd.DataFrame({"Status": ["Pending review", "Shortlisted", "Rejected"], "Candidates": [len(views) - shortlisted - rejected, shortlisted, rejected]})
            fig = px.pie(status_df, values="Candidates", names="Status", hole=0.68, color="Status", color_discrete_map={"Pending review": "#a9b2c4", "Shortlisted": "#28a879", "Rejected": "#ef7373"})
            fig.update_layout(height=280, margin=dict(l=0, r=0, t=5, b=5), paper_bgcolor="rgba(0,0,0,0)", legend=dict(orientation="h", y=-0.08), font=dict(color="#758198"))
            st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})
        skill_left, skill_right = st.columns(2, gap="large")
        with skill_left:
            st.markdown("#### Most demanded skills · resume coverage")
            coverage = {
                skill: sum(skill.casefold() in {found.casefold() for found in item["matched"]} for item in views)
                for skill in required_skills
            }
            if coverage:
                skills_df = pd.DataFrame([{"Required skill": key, "Candidates mentioning skill": value} for key, value in sorted(coverage.items(), key=lambda pair: pair[1], reverse=True)])
                skill_fig = px.bar(skills_df, x="Candidates mentioning skill", y="Required skill", orientation="h", color_discrete_sequence=["#7669e9"])
                skill_fig.update_layout(height=300, margin=dict(l=0, r=5, t=5, b=5), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", yaxis={"categoryorder": "total ascending"}, font=dict(color="#758198"))
                st.plotly_chart(skill_fig, width="stretch", config={"displayModeBar": False})
            else:
                st.caption("Add job criteria in the sidebar to see skill coverage.")
        with skill_right:
            st.markdown("#### Candidates by department")
            department_counts: dict[str, int] = {}
            for item in views:
                department_counts[item["department"]] = department_counts.get(item["department"], 0) + 1
            dept_df = pd.DataFrame([{"Department": key, "Candidates": value} for key, value in department_counts.items()])
            dept_fig = px.bar(dept_df, x="Department", y="Candidates", color_discrete_sequence=["#35ad91"])
            dept_fig.update_layout(height=300, margin=dict(l=0, r=5, t=5, b=5), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#758198"))
            st.plotly_chart(dept_fig, width="stretch", config={"displayModeBar": False})
        with st.expander("Recent screening activity"):
            if st.session_state.activity:
                for activity in st.session_state.activity[:8]:
                    st.markdown(f"· {activity}")
            else:
                st.caption("No recent activity.")

    with tab_candidates:
        st.markdown("#### Candidate review queue")
        st.caption(f"Showing {len(filtered)} of {len(views)} resumes · sorted by the displayed advisory resume score")
        if not filtered:
            st.info("No candidates match these filters. Adjust the sidebar filters.")
        for rank, item in enumerate(filtered, start=1):
            color = score_color(item["resume_score"])
            status_color = {"Shortlisted": "#168253", "Rejected": "#b44b4b", "Pending review": "#68748a"}[item["status"]]
            st.markdown(
                f'<div class="candidate-card"><div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px">'
                f'<div><div style="font-size:1.08rem;font-weight:800;color:#17243a">#{rank} &nbsp;{item["name"]}</div>'
                f'<div class="subtle">{item["email"]} &nbsp;·&nbsp; {item["experience"] if item["experience"] is not None else "Experience not found"} years &nbsp;·&nbsp; {item["education"]}</div></div>'
                f'<div style="text-align:right;color:{status_color};font-weight:700">{item["status"]}<div style="font-size:1.45rem;color:{color}">{item["resume_score"]}%</div></div></div>'
                f'<div class="score-track"><div class="score-fill" style="width:{item["resume_score"]}%;background:{color}"></div></div>'
                f'<div class="subtle">Resume score · skills match {item["score"]}% · ATS text check {item["ats_score"]}% · human review required</div></div>',
                unsafe_allow_html=True,
            )
            with st.expander(f"Review {item['name']} · skills, evidence & resume"):
                detail_left, detail_right = st.columns([1.1, 0.9], gap="large")
                with detail_left:
                    score_cols = st.columns(4)
                    score_cols[0].metric("Resume score", f"{item['resume_score']}%")
                    score_cols[1].metric("Skills match", f"{item['score']}%")
                    score_cols[2].metric("Experience match", f"{item['experience_match']}%" if item["experience_match"] is not None else "Not specified")
                    score_cols[3].metric("Education match", f"{item['education_match']}%" if item["education_match"] is not None else "Not specified")
                    st.progress(item["resume_score"] / 100, text=f"Resume score: {item['resume_score']} / 100")
                    st.markdown("**ATS compatibility checklist (heuristic)**")
                    st.progress(item["ats_score"] / 100, text=f"Text-based ATS readiness: {item['ats_score']} / 100")
                    for issue in item["ats_issues"]:
                        st.markdown(f"- {issue}")
                    st.markdown("**Required skills found in resume text**")
                    if item["matched"]:
                        st.markdown("".join(f'<span class="skill-chip">✓ {skill}</span>' for skill in item["matched"]), unsafe_allow_html=True)
                    else:
                        st.caption("No selected role skills found in the extracted text.")
                    st.markdown("**Required skills not found in extracted text**")
                    if item["missing"]:
                        st.markdown("".join(f'<span class="missing-chip">+ {skill}</span>' for skill in item["missing"]), unsafe_allow_html=True)
                    else:
                        st.caption("No missing selected role skills.")
                    st.markdown("**Resume text summary**")
                    st.write(make_summary(item))
                    st.caption("Summary is generated from matched text only; it is not an assessment of candidate quality.")
                    st.markdown("**Improvement ideas**")
                    for suggestion in item["suggestions"]:
                        st.markdown(f"- {suggestion}")
                    st.caption("Suggestions are generated by simple checks; use only what is accurate and relevant.")
                    current_dept = item["department"] if item["department"] in DEPARTMENTS else "Other"
                    selected_department = st.selectbox("Department", DEPARTMENTS, index=DEPARTMENTS.index(current_dept), key=f"dept_{item['id']}")
                    next(candidate for candidate in candidates if candidate["id"] == item["id"])["department"] = selected_department
                with detail_right:
                    st.markdown("**Human review status**")
                    st.caption("Choose a workflow label after reviewing the resume. Scores never change this status automatically.")
                    button_cols = st.columns(3)
                    action = None
                    if button_cols[0].button("Shortlist", key=f"short_{item['id']}", disabled=item["status"] == "Shortlisted", use_container_width=True):
                        action = "Shortlisted"
                    if button_cols[1].button("Reject", key=f"reject_{item['id']}", disabled=item["status"] == "Rejected", use_container_width=True):
                        action = "Rejected"
                    if button_cols[2].button("Reset", key=f"reset_{item['id']}", disabled=item["status"] == "Pending review", use_container_width=True):
                        action = "Pending review"
                    if action:
                        next(candidate for candidate in candidates if candidate["id"] == item["id"])["status"] = action
                        st.session_state.activity.insert(0, f"{action}: {item['name']} · {datetime.now().strftime('%b %d, %H:%M')}")
                        st.rerun()
                    st.download_button("Download original resume", data=item["data"], file_name=item["filename"], mime="application/pdf" if item["filename"].lower().endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"download_{item['id']}", width="stretch")
                    if item["filename"].lower().endswith(".pdf"):
                        encoded = base64.b64encode(item["data"]).decode("ascii")
                        st.components.v1.html(f'<iframe title="Resume preview" src="data:application/pdf;base64,{encoded}" width="100%" height="470" style="border:1px solid #e7ebf2;border-radius:12px"></iframe>', height=480, scrolling=False)
                    else:
                        st.markdown("**DOCX text preview**")
                        st.text_area("Extracted resume text", item["text"][:12000], height=260, key=f"preview_{item['id']}", label_visibility="collapsed", disabled=True)

    with tab_compare:
        st.markdown("#### Side-by-side comparison")
        st.caption("Compare the role-related evidence that was extracted. This is not a recommendation engine.")
        options = {f"{item['name']} · {item['id']}": item for item in views}
        selected_labels = st.multiselect("Select two or three candidates", list(options), max_selections=3)
        if len(selected_labels) >= 2:
            selected = [options[label] for label in selected_labels]
            rows = [
                ("Resume score", *[f"{item['resume_score']}%" for item in selected]),
                ("Skills match", *[f"{item['score']}%" for item in selected]),
                ("ATS text checklist", *[f"{item['ats_score']}%" for item in selected]),
                ("Review status", *[item["status"] for item in selected]),
                ("Experience stated", *[f"{item['experience']} years" if item["experience"] is not None else "Not found" for item in selected]),
                ("Education text", *[item["education"] for item in selected]),
                ("Department", *[item["department"] for item in selected]),
            ]
            for skill in required_skills:
                rows.append((skill, *["✓ Found" if skill.casefold() in {s.casefold() for s in item["matched"]} else "— Not found" for item in selected]))
            rows.extend([
                ("Experience requirement", *[f"{item['required_experience']} years" if item["required_experience"] is not None else "Not set in role" for item in selected]),
                ("Education requirement", *[", ".join(item["required_education"]) or "Not set in role" for item in selected]),
            ])
            comparison = pd.DataFrame(rows, columns=["Feature", *[item["name"] for item in selected]])
            st.dataframe(comparison, hide_index=True, use_container_width=True)
        else:
            st.info("Select at least two candidates to compare their extracted role-related evidence.")

    with tab_activity:
        st.markdown("#### Screening report")
        report_rows = []
        for item in views:
            report_rows.append({
                "Candidate name": item["name"], "Email": item["email"],
                "Skills found": ", ".join(item["skills"]),
                "Experience (years)": item["experience"] if item["experience"] is not None else "",
                "Education text": item["education"], "Department": item["department"],
                "Resume score": item["resume_score"], "Skills match percentage": item["score"],
                "Experience match": item["experience_match"] if item["experience_match"] is not None else "",
                "Education match": item["education_match"] if item["education_match"] is not None else "",
                "ATS text checklist": item["ats_score"], "Matched skills": ", ".join(item["matched"]),
                "Missing skills": ", ".join(item["missing"]), "Review status": item["status"],
                "Recommendation (advisory)": recommendation(item["score"], required_skills),
            })
        report = pd.DataFrame(report_rows)
        csv_bytes = report.to_csv(index=False).encode("utf-8-sig")
        report_col, pdf_col = st.columns(2)
        with report_col:
            st.download_button("Download screening report (CSV)", data=csv_bytes, file_name="resume_screening_report.csv", mime="text/csv", type="primary", width="stretch")
        with pdf_col:
            pdf_bytes = create_candidate_report_pdf(views, job_title or "Not specified")
            st.download_button("Download analysis report (PDF)", data=pdf_bytes, file_name="resume_analysis_report.pdf", mime="application/pdf", width="stretch")
        st.dataframe(report, hide_index=True, use_container_width=True)
        st.markdown("#### Recent activity")
        if st.session_state.activity:
            for activity in st.session_state.activity[:20]:
                st.markdown(f"· {activity}")
        else:
            st.caption("No status changes or new uploads in this session.")
        st.warning("Use this report as a review aid only. Verify details in original documents, validate criteria for job relevance, and make all employment decisions through a fair, human-led process.")
