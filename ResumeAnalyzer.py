import sys
import re
import docx
import fitz  # PyMuPDF
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog,
    QTableWidget, QTableWidgetItem, QLabel, QTextEdit, QHeaderView
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt


# ----------- Resume Extractor ------------
class ResumeAnalyzer:
    def extract_text(self, filepath):
        text = ""
        if filepath.endswith(".pdf"):
            doc = fitz.open(filepath)
            for page in doc:
                text += page.get_text()
        elif filepath.endswith(".docx"):
            doc = docx.Document(filepath)
            for para in doc.paragraphs:
                text += para.text + "\n"
        return text

    def extract_info(self, text):
        info = {}

        # Email
        email_match = re.search(r'[\w\.-]+@[\w\.-]+', text)
        info["Email"] = email_match.group(0) if email_match else "Not Found"

        # Phone
        phone_match = re.search(r'\+?\d[\d\s-]{8,}\d', text)
        info["Phone"] = phone_match.group(0) if phone_match else "Not Found"

        # Name (heuristic: first line or before email)
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        if lines:
            info["Name"] = lines[0]
        else:
            info["Name"] = "Not Found"

        # Education
        edu_keywords = ["B.Tech", "B.E", "M.Tech", "M.E", "Bachelor", "Master", "University", "School", "College"]
        info["Education"] = "\n".join([line for line in lines if any(k in line for k in edu_keywords)]) or "Not Found"

        # Experience
        exp_keywords = ["Intern", "Experience", "Project", "Work", "Research"]
        info["Experience"] = "\n".join([line for line in lines if any(k in line for k in exp_keywords)]) or "Not Found"

        # Skills
        skills_keywords = ["Python", "SQL", "Machine Learning", "AI", "NLP", "Docker", "Tableau", "Deep Learning", "Git"]
        info["Skills"] = ", ".join([skill for skill in skills_keywords if skill.lower() in text.lower()]) or "Not Found"

        # Summary (longer, multi-line)
        info["Summary"] = (
            f"{info['Name']} has professional and academic background in:\n\n"
            f"Education:\n{info['Education']}\n\n"
            f"Experience:\n{info['Experience']}\n\n"
            f"Skills:\n{info['Skills']}\n"
        )

        # Specialization (Ideal Job Title)
        if "Machine Learning" in info["Skills"] or "Deep Learning" in info["Skills"]:
            info["Specialization"] = "Machine Learning Engineer / Data Scientist"
        elif "SQL" in info["Skills"] or "Tableau" in info["Skills"]:
            info["Specialization"] = "Data Analyst / BI Engineer"
        elif "Docker" in info["Skills"]:
            info["Specialization"] = "MLOps Engineer / DevOps for AI"
        else:
            info["Specialization"] = "Software Engineer / IT Professional"

        return info


# ----------- PyQt5 GUI ------------
class ResumeApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI Resume Analyzer")
        self.setGeometry(200, 200, 1000, 700)

        layout = QVBoxLayout()

        self.label = QLabel("Upload a Resume (PDF/DOCX)")
        self.label.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(self.label)

        self.button = QPushButton("Upload Resume")
        self.button.setFont(QFont("Arial", 11))
        self.button.clicked.connect(self.load_resume)
        layout.addWidget(self.button)

        # Table for details
        self.table = QTableWidget()
        self.table.setFont(QFont("Arial", 10))
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        layout.addWidget(self.table, stretch=2)

        # Summary box (bigger and multi-line)
        self.summary_box = QTextEdit()
        self.summary_box.setReadOnly(True)
        self.summary_box.setFont(QFont("Arial", 11))
        layout.addWidget(QLabel("Resume Summary:"))
        layout.addWidget(self.summary_box, stretch=1)

        self.setLayout(layout)

    def load_resume(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Open Resume", "", "Documents (*.pdf *.docx)")
        if filepath:
            self.label.setText(f"Loaded: {filepath}")
            analyzer = ResumeAnalyzer()
            text = analyzer.extract_text(filepath)
            info = analyzer.extract_info(text)

            # Fill Table
            table_fields = {k: v for k, v in info.items() if k != "Summary"}
            self.table.setRowCount(len(table_fields))
            self.table.setColumnCount(2)
            self.table.setHorizontalHeaderLabels(["Field", "Details"])

            for i, (key, value) in enumerate(table_fields.items()):
                item_key = QTableWidgetItem(key)
                item_value = QTableWidgetItem(value)
                item_value.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                item_value.setTextAlignment(Qt.AlignTop)
                self.table.setItem(i, 0, item_key)
                self.table.setItem(i, 1, item_value)

            # Show summary in text box
            self.summary_box.setText(info["Summary"])


# ----------- Main Run ------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ResumeApp()
    window.show()
    sys.exit(app.exec_())
