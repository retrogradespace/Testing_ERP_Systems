# -*- coding: utf-8 -*-
"""ResumeGenerator.ipynb

# Building Artifacts for Testing ERP Systems

In this notebook, I provide a method for generating resume artifacts that can be used when conducting testing on human resources and enterprise resource planning (ERP) systems. I construct the underlying data used in this resume generator using a different notebook so please review the ApplicantData.ipynb if you would like to generate your own test corpus.
"""

!pip install fpdf python-docx

import pandas as pd
import os
import json
import zipfile
from fpdf import FPDF
from docx import Document
from google.colab import files

from google.colab import drive
drive.mount('/content/drive', force_remount=True)  # wait for auth flow

import os
print(os.listdir('/content/drive/MyDrive/Workflows')[:10])  # should list your Drive items I placed this file in a specific folder so
                                                                                 # I am pointing to it you will need to update with a file path that is true for your drive folder if you choose this method

import os
from pathlib import Path

OUT_FIG_DIR = 'output/figures'
OUT_TAB_DIR = 'output/tables'
OUT_CKP_DIR = 'output/checkpoints'

for d in (OUT_FIG_DIR, OUT_TAB_DIR, OUT_CKP_DIR):
    Path(d).mkdir(parents=True, exist_ok=True)

def safe_save_csv(df, path, **kwargs):
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(p, **kwargs)
    print(f"Saved: {p}")

def safe_save_fig(path, dpi=300):
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    import matplotlib.pyplot as plt
    plt.savefig(p, dpi=dpi)
    print(f"Saved: {p}")

# 3. Load and Merge Data
applicants_df = pd.read_csv('/content/drive/MyDrive/Workflows/final_applicants_unique.csv', encoding='latin1')
history_df = pd.read_csv('/content/drive/MyDrive/Workflows/applicantshistory_consolidated.csv', encoding='latin1')


# Group work history into a list of dictionaries for each applicant
history_grouped = history_df.groupby('ApplicantID').apply(
    lambda x: x[['Organization', 'Role', 'Start Date', 'End Date', 'Achievements']].to_dict('records')
).reset_index()
history_grouped.columns = ['ApplicantID', 'Work_History_List']

# Merge with unique applicant data
df = pd.merge(applicants_df, history_grouped, on='ApplicantID', how='left')

def generate_bio(row):
    """Creates a professional 3-sentence biography based on applicant data."""
    name = row['Full Name']
    role = row['Position Applied For']
    skills = row['Key Skills']
    edu = row['Education']

    bio = (f"{name} is a dedicated professional currently applying for the {role} position. "
           f"With a strong background in {skills}, they bring a wealth of expertise to the team. "
           f"Educated at {edu}, they have consistently demonstrated a commitment to excellence and professional growth.")
    return bio

def enhance_achievement(job):
    """Converts a single achievement bullet into a descriptive job narrative."""
    role = job['Role']
    org = job['Organization']
    ach = job['Achievements']

    narrative = (f"As a {role} at {org}, they were responsible for driving key organizational initiatives. "
                 f"Most notably, they successfully {ach[0].lower() + ach[1:] if ach else 'contributed to team goals'}.")
    return narrative

def clean_text(text):
    if pd.isna(text) or str(text).lower() == 'nan': return "N/A"
    return str(text).encode('latin-1', 'replace').decode('latin-1')

print(df.head())

history_df.head()

# 4. Helper for text encoding/enhancement/cleaning
def generate_pdf(row, filename):
    pdf = FPDF()
    pdf.add_page()

    # Header
    pdf.set_font("Arial", 'B', 20)
    pdf.cell(0, 12, clean_text(row['Full Name']), ln=True, align='L')
    pdf.set_font("Arial", size=10)
    pdf.cell(0, 5, clean_text(f"{row['Email']} | {row['Phone']} | {row['City']}, {row['State']}"), ln=True)
    pdf.line(10, 32, 200, 32)
    pdf.ln(10)

    # NARRATIVE BIO SECTION
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "Professional Profile", ln=True)
    pdf.set_font("Arial", 'I', 10)
    pdf.multi_cell(0, 6, clean_text(generate_bio(row)))
    pdf.ln(5)

    # EXPERIENCE WITH JOB DESCRIPTIONS
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "Professional Experience", ln=True)
    if isinstance(row['Work_History_List'], list):
        for job in row['Work_History_List']:
            pdf.set_font("Arial", 'B', 11)
            pdf.cell(0, 6, clean_text(f"{job['Role']} | {job['Organization']}"), ln=True)
            pdf.set_font("Arial", size=9)
            pdf.cell(0, 5, clean_text(f"{job['Start Date']} - {job['End Date']}"), ln=True)

            # Narrative job description
            pdf.set_font("Arial", size=10)
            pdf.multi_cell(0, 5, clean_text(enhance_achievement(job)))
            pdf.ln(4)

    # OTHER FIELDS
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, "Additional Accomplishments", ln=True)
    pdf.set_font("Arial", size=10)
    pdf.multi_cell(0, 5, f"Awards: {clean_text(row['Awards'])}\nVolunteer: {clean_text(row['Volunteer Work'])}\nPublications: {clean_text(row['Publications'])}")

    pdf.output(filename)

# 5. PDF Generation Function
def generate_docx(row, filename):
    doc = Document()
    doc.add_heading(row['Full Name'], 0)
    doc.add_paragraph(f"{row['Email']} | {row['Phone']} | {row['City']}, {row['State']}")

    doc.add_heading('Professional Profile', level=1)
    doc.add_paragraph(generate_bio(row), style='Intense Quote')

    doc.add_heading('Work History & Descriptions', level=1)
    if isinstance(row['Work_History_List'], list):
        for job in row['Work_History_List']:
            p = doc.add_heading(f"{job['Role']} at {job['Organization']}", level=2)
            doc.add_paragraph(enhance_achievement(job))

    doc.add_heading('Education & Activities', level=1)
    doc.add_paragraph(f"Education: {row['Education']}", style='List Bullet')
    doc.add_paragraph(f"Volunteer Work: {row['Volunteer Work']}", style='List Bullet')
    doc.add_paragraph(f"Publications: {row['Publications']}", style='List Bullet')

    doc.save(filename)

# 6. DOCX Generation Function
output_dir = 'narrative_resumes'
os.makedirs(output_dir, exist_ok=True)
for i, (index, row) in enumerate(df.iterrows()):
    safe_name = "".join([c for c in str(row['Full Name']) if c.isalnum() or c==' ']).strip()
    file_base = os.path.join(output_dir, f"{row['ApplicantID']}_{safe_name}")
    if i < len(df) // 2: generate_pdf(row, f"{file_base}.pdf")
    else: generate_docx(row, f"{file_base}.docx")

zip_name = 'Narrative_Resumes_Pack.zip'
with zipfile.ZipFile(zip_name, 'w') as zipf:
    for root, dirs, files_in_dir in os.walk(output_dir):
        for file in files_in_dir: zipf.write(os.path.join(root, file), file)

files.download(zip_name)



safe_save_csv(df, os.path.join(OUT_TAB_DIR, 'merged_applicants_data.csv'))