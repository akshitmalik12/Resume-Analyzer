#!/usr/bin/env python3
"""
Robust Resume Scanner (PyQt5)

Features:
- PDF (.pdf) and Word (.docx) input
- Normalization, fuzzy heading detection, regex extraction (email/phone/years/company/duration)
- Fuzzy skill detection (typo-tolerant)
- spaCy PhraseMatcher for section headings
- PyQt5 GUI with large fonts, table (multiline), summary (short paragraphs) and suggested job title

Install:
pip install PyQt5 python-docx PyMuPDF spacy rapidfuzz
python -m spacy download en_core_web_sm
"""

import sys
import re
import os
from collections import defaultdict
import docx
import fitz  # PyMuPDF
import math

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QTableWidget, QTableWidgetItem, QLabel, QTextEdit, QHeaderView, QMessageBox
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

# NLP / fuzzy imports
import spacy
from spacy.matcher import PhraseMatcher
from rapidfuzz import process, fuzz

# ---------------------------
# Load spaCy model (once)
# ---------------------------
try:
    nlp = spacy.load("en_core_web_sm")
except Exception as e:
    raise SystemExit("spaCy model not found. Run: python -m spacy download en_core_web_sm") from e

# ---------------------------
# Constants and pools
# ---------------------------
SECTION_KEYWORDS = {
    "experience": ["experience", "work experience", "employment history", "professional experience", "work history", "roles", "career"],
    "education": ["education", "academic", "academic background", "qualifications", "degrees", "education & qualifications"],
    "skills": ["skills", "technical skills", "key skills", "competencies", "expertise", "skillset"],
    "projects": ["projects", "personal projects", "academic projects", "selected projects"],
    "certifications": ["certifications", "licenses"],
    "summary": ["summary", "profile", "professional summary", "about", "about me", "career objective"],
    "interests": ["interests", "areas of interest", "hobbies"],
}

# flattened skill pool (extendable)
SKILL_POOL = [
    "python","sql","machine learning","deep learning","nlp","docker","git","tableau",
    "pandas","numpy","tensorflow","pytorch","scikit-learn","keras","aws","spark",
    "excel","power bi","c++","java","javascript","react","flask","fastapi","linux",
    "nosql","mongodb","hadoop","matplotlib","seaborn","etl","rest api","aws s3","ci/cd",
    "keras","spark","spark sql","statistics","r","sas","scala","kubernetes","tensorflow lite"
]

# regex patterns
RE_EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", re.I)
RE_PHONE = re.compile(r"(?:\+?\d{1,3}[\s\-\.]?)?(?:\(?\d{2,4}\)?[\s\-\.]?)?\d{6,12}")
RE_YEAR = re.compile(r"\b(19|20)\d{2}\b")
RE_COMPANY_DUR = re.compile(r"([A-Z][A-Za-z0-9&\.\- ,]{2,80}?)\s*\(\s*([A-Za-z0-9 ,\-\–\—\/]+)\s*\)")

# PhraseMatcher for headings
phrase_matcher = PhraseMatcher(nlp.vocab, attr="LOWER")
for sec, kws in SECTION_KEYWORDS.items():
    phrase_matcher.add(sec, [nlp.make_doc(k) for k in kws])


# ---------------------------
# Utility functions
# ---------------------------
def normalize_text(text: str) -> str:
    """Lowercase, normalize whitespace, convert smart quotes, keep line breaks for section splitting."""
    if not text:
        return ""
    # normalize unicode quotes and dashes
    text = text.replace("\u2013", "-").replace("\u2014", "-").replace("\u2018", "'").replace("\u2019", "'").replace("\u201c", '"').replace("\u201d", '"')
    # unify newlines and spaces
    text = re.sub(r"\r\n?", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    # strip leading/trailing spaces for each line
    lines = [ln.strip() for ln in text.splitlines()]
    return "\n".join(lines).strip()


def extract_text_from_docx(path: str) -> str:
    try:
        doc = docx.Document(path)
        paragraphs = [p.text for p in doc.paragraphs]
        return "\n".join(paragraphs)
    except Exception:
        # fallback to docx2txt if needed (not imported here to keep deps minimal)
        return ""


def extract_text_from_pdf(path: str) -> str:
    try:
        doc = fitz.open(path)
        pages = []
        for page in doc:
            pages.append(page.get_text("text") or "")
        return "\n".join(pages)
    except Exception:
        return ""


def split_into_lines(text: str):
    return [ln for ln in text.splitlines() if ln.strip()]


def fuzzy_section_heading(line: str, threshold=80):
    """Return matched section key if line looks like a heading via fuzzy match / phrase matcher"""
    if not line:
        return None
    doc = nlp(line)
    # phrase matcher first (fast, exact phrase variants)
    matches = phrase_matcher(doc)
    if matches:
        # return the label of first match
        match_id, start, end = matches[0]
        return nlp.vocab.strings[match_id]
    # fuzzy against keywords
    best = None
    best_score = 0
    for sec, kws in SECTION_KEYWORDS.items():
        # check each keyword phrase fuzzily
        for kw in kws:
            score = fuzz.token_set_ratio(line.lower(), kw.lower())
            if score > best_score:
                best_score = score
                best = sec
    if best_score >= threshold:
        return best
    return None


def find_name(lines, email_candidate=None):
    """Heuristic name extraction:
       - line above email (if exists)
       - any spaCy PERSON in top lines
       - first plausible header line
    """
    # 1) email neighbor
    if email_candidate:
        for i, ln in enumerate(lines):
            if email_candidate in ln:
                if i > 0:
                    cand = lines[i - 1].strip()
                    if plausible_name(cand):
                        return titlecase(cand)
    # 2) spaCy PERSON in header
    header = " ".join(lines[:12])
    doc = nlp(header)
    persons = [ent.text for ent in doc.ents if ent.label_ == "PERSON"]
    if persons:
        # prefer longer 2+ word names
        for p in persons:
            if len(p.split()) >= 2:
                return titlecase(p)
        return titlecase(persons[0])
    # 3) first plausible line
    for ln in lines[:6]:
        if plausible_name(ln):
            return titlecase(ln)
    return "Not Found"


def plausible_name(s: str) -> bool:
    if not s: return False
    s = s.strip()
    if len(s) < 3 or len(s) > 50: return False
    if re.search(r"(@|www\.|http|resume|curriculum|objective|linkedin|github)", s.lower()): return False
    words = [w for w in s.split() if re.search(r"[A-Za-z]", w)]
    return 1 < len(words) <= 4


def titlecase(s: str) -> str:
    return " ".join([w.capitalize() for w in re.sub(r"[^A-Za-z\s\-']", " ", s).split()])


def extract_emails(text: str):
    return list(dict.fromkeys(RE_EMAIL.findall(text)))  # unique preserve order


def extract_phones(text: str):
    raw = RE_PHONE.findall(text)
    # cleanup and pick plausible ones
    cleaned = []
    for r in raw:
        digits = re.sub(r"\D", "", r)
        if 8 <= len(digits) <= 15:
            cleaned.append(r.strip())
    # dedupe
    return list(dict.fromkeys(cleaned))


def extract_years(text: str):
    return list(dict.fromkeys(RE_YEAR.findall(text)))  # NOTE: findall returns tuples for groups, but okay for detection


def extract_company_duration(text: str):
    items = []
    for m in RE_COMPANY_DUR.finditer(text):
        comp = m.group(1).strip()
        dur = m.group(2).strip()
        items.append(f"{comp} ({dur})")
    return items


def fuzzy_find_skills(text: str, pool=SKILL_POOL, limit=20, threshold=80):
    """Use rapidfuzz to find best skill matches from pool allowing typos."""
    found = set()
    txt = text.lower()
    # Try direct substring first (fast)
    for s in pool:
        if s.lower() in txt:
            found.add(s)
    # If not many found, fuzzy-extract tokens/phrases
    if len(found) < 5:
        # consider splitting text into phrases (by comma/pipe/semicolon and by newline)
        candidates = set(re.split(r"[\n,;•\-–]+", text))
        for cand in candidates:
            cand = cand.strip()
            if not cand or len(cand) < 3: continue
            # find best match from pool
            best = process.extractOne(cand, pool, scorer=fuzz.token_sort_ratio)
            if best and best[1] >= threshold:
                found.add(best[0])
    # return sorted heuristically by pool order
    ordered = [s for s in pool if s in found]
    return ordered[:limit] if ordered else []


# ---------------------------
# High-level extraction function
# ---------------------------
def parse_resume_text(raw_text: str):
    """
    Parses resume text robustly and returns a dict with fields:
      Name, Email(s), Phone(s), Education (list), Experience (list), Skills (list),
      Summary (multi-paragraph), Specialization (suggested job title)
    """
    text = normalize_text(raw_text)
    lines = split_into_lines(text)

    # Extract emails & phones
    emails = extract_emails(raw_text)
    phones = extract_phones(raw_text)

    # Name detection
    name = find_name(lines, emails[0] if emails else None)

    # Section segmentation: find headings using fuzzy heading detection
    sections = defaultdict(list)
    current_section = "preamble"
    for i, ln in enumerate(lines):
        # if a heading-like short line, detect
        heading = fuzzy_section_heading(ln)
        if heading:
            current_section = heading
            continue
        sections[current_section].append(ln)

    # If sections empty, fall back to whole doc as preamble
    # Convert section lists to text
    section_text = {k: "\n".join(v).strip() for k, v in sections.items()}

    # Education extraction: look for degree keywords within education section first, then globally
    edu_lines = []
    if section_text.get("education"):
        edu_lines = [ln for ln in section_text["education"].splitlines() if len(ln) > 3]
    else:
        for ln in lines:
            if any(d.lower() in ln.lower() for d in ["b.tech","b.e","m.tech","m.e","bachelor","master","msc","bsc","phd","mba","degree","university","college","institute"]):
                edu_lines.append(ln)
    # dedupe & shorten
    education = list(dict.fromkeys(edu_lines))[:8]

    # Experience extraction: from experience section OR search for role-like lines / durations
    exp_lines = []
    if section_text.get("experience"):
        exp_lines = [ln for ln in section_text["experience"].splitlines() if len(ln) > 3]
    else:
        # heuristics: lines with 'intern', 'engineer', 'developer', 'worked', 'project' or containing years/date ranges
        for ln in lines:
            lnl = ln.lower()
            if any(k in lnl for k in ["intern", "engineer", "developer", "consultant", "analyst", "manager", "worked", "project", "research"]):
                exp_lines.append(ln)
            elif re.search(r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b|\d{4}", ln, re.I):
                # likely date mention, include nearby line
                exp_lines.append(ln)
    # also try to extract company(duration) patterns
    company_durs = extract_company_duration(raw_text)
    for cd in company_durs:
        if cd not in exp_lines:
            exp_lines.insert(0, cd)
    experience = list(dict.fromkeys(exp_lines))[:12]

    # Skills: try skills section first, then fuzzy in whole doc
    skills = []
    if section_text.get("skills"):
        skills = fuzzy_find_skills(section_text["skills"], SKILL_POOL, threshold=70)
    if not skills:
        skills = fuzzy_find_skills(raw_text, SKILL_POOL, threshold=78)

    # Specialization: determine job title suggestion from skills
    specialization = "Software Engineer / IT Professional"
    skset = set([s.lower() for s in skills])
    if any(k in skset for k in ["machine learning", "deep learning", "nlp", "tensorflow", "pytorch"]):
        specialization = "Machine Learning Engineer / Data Scientist"
    elif any(k in skset for k in ["sql", "tableau", "power bi", "excel", "bi"]):
        specialization = "Data Analyst / BI Engineer"
    elif any(k in skset for k in ["docker", "kubernetes", "ci/cd", "aws", "devops", "kubectl"]):
        specialization = "MLOps / DevOps Engineer"
    elif any(k in skset for k in ["react","javascript","flask","fastapi","django","api"]):
        specialization = "Full-stack / Backend Engineer"

    # Build short paragraph-style summary with insights
    summary_parts = []
    if name and name != "Not Found":
        summary_parts.append(f"{name} is a candidate with the following detected background.")
    if education:
        summary_parts.append(f"Education highlights include: {education[0] + (', ...' if len(education) > 1 else '')}.")
    if experience:
        summary_parts.append(f"Experience snippets: {experience[0] + (', ...' if len(experience) > 1 else '')}.")
    if skills:
        summary_parts.append(f"Key skills detected: {', '.join(skills[:6])}.")
    summary_parts.append(f"Suggested role: {specialization}.")

    summary = "\n\n".join(summary_parts)

    return {
        "Name": name,
        "Emails": emails if emails else ["Not Found"],
        "Phones": phones if phones else ["Not Found"],
        "Education": education if education else ["Not Found"],
        "Experience": experience if experience else ["Not Found"],
        "Skills": skills if skills else ["Not Found"],
        "Specialization": specialization,
        "Summary": summary
    }


# ---------------------------
# PyQt5 GUI
# ---------------------------
class ResumeApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Resume Analyzer — Robust")
        self.setGeometry(160, 120, 1100, 760)

        font_large = QFont("Segoe UI", 12)
        self.setFont(font_large)

        layout = QVBoxLayout()

        header_row = QHBoxLayout()
        self.btn_upload = QPushButton("Upload Resume (PDF / DOCX / TXT / Image)")
        self.btn_upload.setFont(QFont("Segoe UI", 12))
        self.btn_upload.clicked.connect(self.upload_resume)
        header_row.addWidget(self.btn_upload)

        self.btn_export = QPushButton("Export CSV")
        self.btn_export.setFont(QFont("Segoe UI", 11))
        self.btn_export.clicked.connect(self.export_csv)
        header_row.addWidget(self.btn_export)

        layout.addLayout(header_row)

        # Table for extracted fields
        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Field", "Details"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setWordWrap(True)
        self.table.setFont(QFont("Segoe UI", 11))
        layout.addWidget(self.table, stretch=2)

        # Summary
        layout.addWidget(QLabel("Summary & Insights:"))
        self.summary_box = QTextEdit()
        self.summary_box.setReadOnly(True)
        self.summary_box.setFont(QFont("Segoe UI", 12))
        self.summary_box.setMinimumHeight(180)
        layout.addWidget(self.summary_box, stretch=1)

        self.setLayout(layout)

        # state
        self.last_parsed = None

    def upload_resume(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Resume", "", "All Files (*);;PDF (*.pdf);;Word (*.docx);;Text (*.txt)")
        if not path:
            return
        # extract text robustly
        text = ""
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == ".pdf":
                text = extract_text_from_pdf(path)
            elif ext == ".docx":
                text = extract_text_from_docx(path)
            elif ext == ".txt":
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    text = f.read()
            else:
                # try pdf/docx anyway
                text = extract_text_from_pdf(path) or extract_text_from_docx(path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read file: {e}")
            return

        if not text.strip():
            QMessageBox.warning(self, "No text", "Could not extract text from this file.")
            return

        # parse
        parsed = parse_resume_text(text)
        self.last_parsed = parsed
        self.populate_table(parsed)
        self.summary_box.setPlainText(parsed["Summary"])

    def populate_table(self, parsed: dict):
        # fields to display (ordered)
        rows = [
            ("Name", parsed.get("Name", "")),
            ("Emails", ", ".join(parsed.get("Emails", []))),
            ("Phones", ", ".join(parsed.get("Phones", []))),
            ("Education", "\n".join(parsed.get("Education", []))),
            ("Experience", "\n".join(parsed.get("Experience", []))),
            ("Skills", ", ".join(parsed.get("Skills", []))),
            ("Specialization", parsed.get("Specialization", "")),
        ]
        self.table.setRowCount(len(rows))
        for r, (k, v) in enumerate(rows):
            key_item = QTableWidgetItem(k)
            key_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            val_item = QTableWidgetItem(v)
            val_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            val_item.setTextAlignment(Qt.AlignLeft | Qt.AlignTop)
            val_item.setToolTip(v)
            self.table.setItem(r, 0, key_item)
            self.table.setItem(r, 1, val_item)
        # Force resize rows to fit content
        self.table.resizeRowsToContents()

    def export_csv(self):
        if not self.last_parsed:
            QMessageBox.information(self, "No data", "Parse a resume first.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV", "", "CSV Files (*.csv)")
        if not path:
            return
        try:
            import csv
            p = self.last_parsed
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Field","Value"])
                w.writerow(["Name", p.get("Name","")])
                w.writerow(["Emails", "; ".join(p.get("Emails",[]))])
                w.writerow(["Phones", "; ".join(p.get("Phones",[]))])
                w.writerow(["Education", " | ".join(p.get("Education",[]))])
                w.writerow(["Experience", " | ".join(p.get("Experience",[]))])
                w.writerow(["Skills", " | ".join(p.get("Skills",[]))])
                w.writerow(["Specialization", p.get("Specialization","")])
            QMessageBox.information(self, "Saved", "CSV exported.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save CSV: {e}")


# ---------------------------
# Run app
# ---------------------------
def main():
    app = QApplication(sys.argv)
    win = ResumeApp()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
