#!/usr/bin/env python3
"""
AI Resume Scanner with PyQt5 — Dark Professional Dashboard
Features:
- Multi-threaded parsing for responsive UI
- Extracts Name, Emails, Phones, Education, Experience, Skills, Specialization
- Export results to CSV
- Dark-themed, modern dashboard
"""

import sys
import re
import os
from collections import defaultdict
import docx
import fitz  # PyMuPDF

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QTableWidget, QTableWidgetItem, QLabel, QTextEdit, QHeaderView, QMessageBox,
    QProgressBar, QSizePolicy
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, QThread, pyqtSignal

import spacy
from spacy.matcher import PhraseMatcher
from rapidfuzz import process, fuzz

# ---------------------------
# Load spaCy
# ---------------------------
try:
    nlp = spacy.load("en_core_web_sm")
except Exception as e:
    raise SystemExit("Run: python -m spacy download en_core_web_sm") from e

# ---------------------------
# Skill pools
# ---------------------------
SKILL_POOL = {
    "Software & Programming": {
        "score": 5, "keywords": ["python", "java", "c++", "c", "c#", "javascript", "sql", "git", "rest api", "ci/cd", "agile", "scrum", "jira"]
    },
    "Data Science & AI": {
        "score": 10, "keywords": ["machine learning", "deep learning", "nlp", "tensorflow", "pytorch", "scikit-learn", "pandas", "numpy", "matplotlib", "seaborn", "r", "spark", "hadoop", "bi", "tableau", "power bi"]
    },
    "Electrical & Electronics Engineering": {
        "score": 10, "keywords": ["verilog", "embedded c", "matlab", "xilinx vivado", "altium designer", "eagle", "pcb design", "circuit design", "vsim", "arduino", "raspberry pi", "microcontrollers", "vlsi", "fpga", "hfss", "ece"]
    },
    "Mechanical & Aerospace Engineering": {
        "score": 8, "keywords": ["autocad", "solidworks", "ansys", "catia", "cad", "cam", "thermodynamics", "fluid mechanics", "materials science", "fea", "robotics", "3d printing", "propulsion", "aerodynamics"]
    },
    "Civil Engineering": {
        "score": 8, "keywords": ["autocad civil 3d", "revit", "etabs", "staad pro", "safe", "gis", "structural analysis", "geotechnical engineering", "transportation engineering"]
    },
    "Web Development": {
        "score": 5, "keywords": ["html", "css", "javascript", "react", "angular", "vue.js", "node.js", "django", "flask"]
    },
    "Project Management & Business": {
        "score": 4, "keywords": ["project management", "scrum", "agile", "jira", "confluence", "microsoft office", "excel", "powerpoint", "finance", "marketing"]
    }
}

FLATTENED_SKILL_POOL = [skill for sublist in [d['keywords'] for d in SKILL_POOL.values()] for skill in sublist]

# ---------------------------
# Regex patterns
# ---------------------------
RE_EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", re.I)
RE_PHONE = re.compile(r"(?:\+?\d{1,3}[\s\-\.]?)?(?:\(?\d{2,4}\)?[\s\-\.]?)?\d{6,12}")
RE_DATE_RANGE = re.compile(
    r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+\d{4}\s*[\-–]\s*(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+\d{4}|(present|current|until now)", re.I)

SECTION_KEYWORDS = {
    "experience": ["experience", "work experience", "employment history", "professional experience", "work history", "roles", "career"],
    "education": ["education", "academic", "academic background", "qualifications", "degrees", "education & qualifications"],
    "skills": ["skills", "technical skills", "key skills", "competencies", "expertise", "skillset"],
    "projects": ["projects", "personal projects", "academic projects", "selected projects", "research"],
    "certifications": ["certifications", "licenses"],
    "summary": ["summary", "profile", "professional summary", "about", "about me", "career objective"],
    "awards": ["awards", "honors", "achievements"],
    "publications": ["publications", "research papers"]
}
phrase_matcher = PhraseMatcher(nlp.vocab, attr="LOWER")
for sec, kws in SECTION_KEYWORDS.items():
    phrase_matcher.add(sec, [nlp.make_doc(k) for k in kws])

# ---------------------------
# Utility functions
# ---------------------------
def normalize_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\u2013", "-").replace("\u2014", "-").replace("\u2018", "'").replace("\u2019", "'").replace("\u201c", '"').replace("\u201d", '"')
    text = re.sub(r"\r\n?", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    lines = [ln.strip() for ln in text.splitlines()]
    return "\n".join(lines).strip()

def extract_text_from_docx(path: str) -> str:
    try:
        doc = docx.Document(path)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception:
        return ""

def extract_text_from_pdf(path: str) -> str:
    try:
        doc = fitz.open(path)
        return "\n".join([page.get_text("text") or "" for page in doc])
    except Exception:
        return ""

def split_into_lines(text: str):
    return [ln for ln in text.splitlines() if ln.strip()]

def fuzzy_section_heading(line: str, threshold=80):
    if not line:
        return None
    doc_line = nlp(line)
    matches = phrase_matcher(doc_line)
    if matches:
        match_id, start, end = matches[0]
        return nlp.vocab.strings[match_id]
    best = None
    best_score = 0
    for sec, kws in SECTION_KEYWORDS.items():
        for kw in kws:
            score = fuzz.token_set_ratio(line.lower(), kw.lower())
            if score > best_score:
                best_score = score
                best = sec
    if best_score >= threshold:
        return best
    return None

def fuzzy_find_skills(text: str, pool=FLATTENED_SKILL_POOL, limit=30, threshold=80):
    found = set()
    doc = nlp(text.lower())
    skill_matcher = PhraseMatcher(nlp.vocab, attr="LOWER")
    skill_patterns = [nlp.make_doc(s) for s in pool]
    skill_matcher.add("SKILLS", skill_patterns)
    matches = skill_matcher(doc)
    for match_id, start, end in matches:
        found.add(doc[start:end].text)
    if len(found) < 10:
        tokens = [token.text for token in doc if not token.is_punct and not token.is_space and len(token) > 2]
        for token in tokens:
            best = process.extractOne(token, pool, scorer=fuzz.token_set_ratio)
            if best and best[1] >= threshold:
                found.add(best[0].lower())
    return sorted(list(found))[:limit]

def get_specialization(skills, education_lines, experience_lines):
    skill_set = set(skills)
    scores = defaultdict(int)
    for category, data in SKILL_POOL.items():
        for keyword in data['keywords']:
            if keyword in skill_set:
                scores[category] += data['score']
    education_text = " ".join(education_lines).lower()
    experience_text = " ".join(experience_lines).lower()
    if any(s in education_text for s in ["computer science", "cs", "software"]):
        scores["Software & Programming"] += 15
    if any(s in education_text for s in ["electronics", "ece"]):
        scores["Electrical & Electronics Engineering"] += 15
    if any(s in education_text for s in ["mechanical", "aerospace"]):
        scores["Mechanical & Aerospace Engineering"] += 15
    if any(s in education_text for s in ["civil engineering", "civil"]):
        scores["Civil Engineering"] += 15
    if any(s in experience_text for s in ["data scientist", "machine learning"]):
        scores["Data Science & AI"] += 15
    if any(s in experience_text for s in ["web developer", "full-stack"]):
        scores["Web Development"] += 10
    total_score = sum(scores.values())
    if not total_score:
        return "General Professional", {}
    weighted_scores = {k: (v / total_score) * 100 for k, v in scores.items()}
    best_category = max(weighted_scores, key=weighted_scores.get)
    mapping = {
        "Software & Programming": "Software Engineer / Developer",
        "Data Science & AI": "Machine Learning Engineer / Data Scientist",
        "Electrical & Electronics Engineering": "Electronics Engineer / ECE Professional",
        "Mechanical & Aerospace Engineering": "Mechanical Engineer",
        "Civil Engineering": "Civil Engineer",
        "Web Development": "Web Developer",
        "Project Management & Business": "Project Management / Business Analyst"
    }
    return mapping.get(best_category, "General Professional"), weighted_scores

# ---------------------------
# Parser Worker
# ---------------------------
class ParserWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, text):
        super().__init__()
        self.text = text

    def run(self):
        try:
            parsed = self.parse_resume_text(self.text)
            self.finished.emit(parsed)
        except Exception as e:
            self.error.emit(f"Parsing failed: {e}")

    def parse_resume_text(self, raw_text: str):
        text = normalize_text(raw_text)
        lines = split_into_lines(text)

        # Improved Name Extraction
        def extract_name(lines):
            for ln in lines[:15]:
                ln = ln.strip()
                # Title case
                match = re.match(r"^([A-Z][a-z]+(?:\s[A-Z][a-z]+)+)$", ln)
                if match:
                    return match.group(1).strip()
                # All caps
                if ln.isupper() and len(ln.split()) >= 2 and all(w.isalpha() for w in ln.split()):
                    return ln.title()
            for ln in lines[:20]:
                if "name" in ln.lower():
                    parts = ln.split(":")
                    if len(parts) > 1:
                        return parts[1].strip().title()
            return "Not Found"

        def extract_emails(text):
            return re.findall(RE_EMAIL, text) or ["Not Found"]

        def extract_phones(text):
            phone_pattern = r"(\+?\d{1,3}[-\s]?)?\(?\d{2,4}\)?[-\s]?\d{3,4}[-\s]?\d{4}"
            matches = re.findall(phone_pattern, text)
            results = []
            for m in matches:
                num = "".join(m) if isinstance(m, tuple) else m
                digits = re.sub(r"\D", "", num)
                if 8 <= len(digits) <= 15:
                    results.append(num.strip())
            return list(dict.fromkeys(results)) or ["Not Found"]

        name = extract_name(lines)
        emails = extract_emails(raw_text)
        phones = extract_phones(raw_text)

        # Sections
        sections = defaultdict(list)
        current_section = "preamble"
        for ln in lines:
            heading = fuzzy_section_heading(ln)
            if heading:
                current_section = heading
                continue
            sections[current_section].append(ln)
        section_text = {k: "\n".join(v).strip() for k, v in sections.items()}

        edu_lines = [ln for ln in section_text.get("education", "").splitlines() if len(ln) > 3]
        if not edu_lines:
            for ln in lines:
                if any(d.lower() in ln.lower() for d in ["bachelor", "master", "phd", "degree", "university", "college", "institute", "school"]):
                    edu_lines.append(ln)
        education = list(dict.fromkeys(edu_lines))[:8]

        experience_lines = [ln for ln in section_text.get("experience", "").splitlines() if len(ln) > 3]
        if not experience_lines:
            for ln in lines:
                if any(k in ln.lower() for k in ["engineer", "developer", "analyst", "intern", "project", "worked"]) or re.search(RE_DATE_RANGE, ln):
                    experience_lines.append(ln)
        experience = list(dict.fromkeys(experience_lines))[:15]

        skills = fuzzy_find_skills(raw_text, FLATTENED_SKILL_POOL, threshold=78)
        specialization, scores = get_specialization(skills, education, experience)

        summary_parts = []
        if name != "Not Found":
            summary_parts.append(f"A profile for {name} has been successfully parsed. The system uses an advanced NLP pipeline.")
        if education:
            summary_parts.append(f"Education highlights include: {education[0] + (', ...' if len(education) > 1 else '')}.")
        if experience:
            summary_parts.append(f"The experience section details roles such as '{experience[0].split('(')[0].strip()}' and mentions key projects.")
        if skills:
            summary_parts.append(f"The candidate's core skills are in **{specialization}**, with a strong command of: {', '.join(skills[:6]) + '...' if len(skills)>6 else ''}.")
        summary_parts.append(f"Overall, the resume suggests a strong fit for a **{specialization}** role.")

        summary = "\n\n".join(summary_parts)

        return {
            "Name": name,
            "Emails": emails,
            "Phones": phones,
            "Education": education,
            "Experience": experience,
            "Skills": skills,
            "Specialization": specialization,
            "SpecializationScores": scores,
            "Summary": summary
        }

# ---------------------------
# PyQt5 GUI
# ---------------------------
class ResumeApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI Resume Scanner — Professional Dashboard")
        self.setGeometry(100, 100, 1000, 700)
        self.setStyleSheet("background-color: #2c3e50; color: #ecf0f1;")

        self.initUI()
        self.current_resume_text = ""

    def initUI(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(10,10,10,10)

        # Header
        header_layout = QHBoxLayout()
        self.upload_btn = QPushButton("Upload Resume")
        self.upload_btn.setStyleSheet("background-color: #3498db; color: #ecf0f1; font-weight:bold; font-size:16px; padding:8px;")
        self.upload_btn.clicked.connect(self.upload_resume)

        self.export_btn = QPushButton("Export to CSV")
        self.export_btn.setStyleSheet("background-color: #1abc9c; color: #2c3e50; font-weight:bold; font-size:16px; padding:8px;")
        self.export_btn.clicked.connect(self.export_csv)

        header_layout.addWidget(self.upload_btn)
        header_layout.addWidget(self.export_btn)
        header_layout.addStretch()

        # Progress Bar
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 2px solid #34495e;
                border-radius: 5px;
                text-align: center;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #3498db;
            }
        """)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Field", "Value"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setStyleSheet("background-color: #34495e; color: #ecf0f1; font-size:14px;")
        self.table.verticalHeader().setVisible(False)

        # Summary
        self.summary_label = QLabel("AI Insights & Summary")
        self.summary_label.setFont(QFont("Arial", 14, QFont.Bold))
        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        self.summary_text.setStyleSheet("background-color: #34495e; color: #ecf0f1; font-size:14px;")

        layout.addLayout(header_layout)
        layout.addWidget(self.progress)
        layout.addWidget(self.table)
        layout.addWidget(self.summary_label)
        layout.addWidget(self.summary_text)

        self.setLayout(layout)

    def upload_resume(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Resume", "", "PDF Files (*.pdf);;Word Files (*.docx);;Text Files (*.txt)")
        if not path:
            return

        ext = os.path.splitext(path)[1].lower()
        if ext == ".pdf":
            self.current_resume_text = extract_text_from_pdf(path)
        elif ext == ".docx":
            self.current_resume_text = extract_text_from_docx(path)
        else:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                self.current_resume_text = f.read()

        if not self.current_resume_text.strip():
            QMessageBox.warning(self, "Error", "Could not extract text from resume.")
            return

        self.progress.setVisible(True)
        self.progress.setRange(0,0)  # Busy indicator

        self.worker = ParserWorker(self.current_resume_text)
        self.worker.finished.connect(self.display_result)
        self.worker.error.connect(self.show_error)
        self.worker.start()

    def display_result(self, result):
        self.progress.setVisible(False)
        self.table.setRowCount(0)

        for field, value in result.items():
            if isinstance(value, list):
                value = ", ".join(value)
            elif isinstance(value, dict):
                value = ", ".join([f"{k}:{v:.1f}%" for k,v in value.items()])
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(field))
            self.table.setItem(row, 1, QTableWidgetItem(str(value)))

        self.summary_text.setPlainText(result.get("Summary", ""))

    def show_error(self, msg):
        self.progress.setVisible(False)
        QMessageBox.critical(self, "Error", msg)

    def export_csv(self):
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "Warning", "No data to export")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV", "", "CSV Files (*.csv)")
        if not path:
            return
        import csv
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Field", "Value"])
            for r in range(self.table.rowCount()):
                field = self.table.item(r,0).text()
                value = self.table.item(r,1).text()
                writer.writerow([field,value])
        QMessageBox.information(self, "Success", f"Data exported to {path}")

# ---------------------------
# Run App
# ---------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ResumeApp()
    window.show()
    sys.exit(app.exec_())
