import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import json
import urllib.parse
import io
import base64
import sqlite3
import random
import zipfile
import hashlib
import shutil
import requests
from datetime import datetime
from PIL import Image
import plotly.graph_objects as go

# ReportLab Imports for PDF, Watermark, Signatures, Borders & Certificate Generation
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="PM SHRI JNV CHHOTAUDEPUR - RESULT PORTAL", page_icon="🎓", layout="wide")

# Custom CSS & Responsive Mobile Optimization
st.markdown("""
    <style>
    @media (max-width: 768px) {
        .stApp { padding: 4px !important; }
        h1 { font-size: 1.1rem !important; }
        h2 { font-size: 0.95rem !important; }
        h3 { font-size: 0.85rem !important; }
        .stButton>button { width: 100% !important; padding: 6px 10px !important; }
        .mobile-hide { display: none !important; }
        .topper-card { width: 160px !important; padding: 8px !important; margin-right: 8px !important; }
        .topper-card img { width: 60px !important; height: 60px !important; }
    }
    .header-logo-left {
        display: flex;
        justify-content: flex-start;
        align-items: center;
    }
    .header-logo-right {
        display: flex;
        justify-content: flex-end;
        align-items: center;
    }
    .main-title {
        text-transform: uppercase;
        font-weight: 800;
        color: #1E88E5;
        margin: 0 !important;
        padding: 0 !important;
        font-size: 1.45rem;
        text-align: center;
    }
    .sub-title {
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #2E7D32;
        font-weight: 700;
        margin-top: 4px !important;
        padding: 0 !important;
        font-size: 1.05rem;
        text-align: center;
    }
    .weak-badge {
        background-color: #FFEBEE;
        border-left: 5px solid #E53935;
        padding: 10px;
        border-radius: 5px;
        color: #B71C1C;
        font-weight: bold;
        margin: 10px 0;
    }
    .topper-card {
        display: inline-block;
        width: 210px;
        background: #ffffff;
        padding: 12px;
        margin-right: 15px;
        border-radius: 10px;
        border: 2px solid #1E88E5;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        vertical-align: top;
    }
    .hall-of-fame-box {
        background: linear-gradient(135deg, #E3F2FD, #FFF9C4);
        border: 2px solid #1565C0;
        border-radius: 10px;
        padding: 12px;
        margin-bottom: 20px;
    }
    .notice-box {
        background-color: #E8F5E9;
        border-left: 5px solid #2E7D32;
        padding: 10px 15px;
        border-radius: 5px;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# Directory Setup
for folder in ["photos/students", "photos/gallery", "photos/board", "photos/system", "backups"]:
    os.makedirs(folder, exist_ok=True)

# System Image Paths
BG_PATH = "photos/system/background.png"
LOGO_PATH = "photos/system/logo.png"
SEAL_PATH = "photos/system/seal.png"
SIGN_PATH = "photos/system/signature.png"
NOTICES_FILE = "notices.json"
BOARD_TOPPERS_FILE = "board_toppers.json"
LOG_FILE = "result_logs.csv"

# Password Hashing & Security Helper
DEFAULT_PASS = "Jnvcu@me2"
ADMIN_PASS_HASH = hashlib.sha256(st.secrets.get("ADMIN_PASSWORD", os.getenv("ADMIN_PASSWORD", DEFAULT_PASS)).encode()).hexdigest()

DB_FILE = "school_database.db"
EXCEL_FILE_PATH = "JNV_Student_Marks.xlsx"

ALL_SUBJECTS = [
    'Gujarati', 'Hindi', 'English', 'Mathematics', 
    'Science', 'Social_Science', 'Physics', 'Chemistry', 'Biology'
]

LANG_TEXTS = {
    "English": {
        "title": "STUDENT PERFORMANCE & RESULT PORTAL",
        "search_lbl": "🔎 CHECK STUDENT RESULT",
        "cert_btn": "🏆 Download Merit Certificate",
        "admit_btn": "🪪 Download Admit Card",
        "weak_alert": "⚠️ Needs Special Attention in the following subjects (< 33%):",
        "chart_title": "Subject-wise Performance Breakdown",
        "combined_title": "🌟 COMBINED OVERALL PERFORMANCE ACROSS ALL EXAMS",
        "trend_title": "📈 Multi-Exam Performance Trend Line"
    },
    "Hindi": {
        "title": "विद्यार्थी प्रदर्शन एवं परिणाम पोर्टल",
        "search_lbl": "🔎 छात्र परिणाम खोजें",
        "cert_btn": "🏆 योग्यता प्रमाण पत्र (Merit Certificate) डाउनलोड करें",
        "admit_btn": "🪪 प्रवेश पत्र (Admit Card) डाउनलोड करें",
        "weak_alert": "⚠️ निम्नलिखित विषयों में विशेष ध्यान देने की आवश्यकता है (< 33%):",
        "chart_title": "विषय-वार अंक विश्लेषण",
        "combined_title": "🌟 सभी परीक्षाओं का संयुक्त प्रदर्शन (Combined Performance)",
        "trend_title": "📈 मल्टी-एग्जाम प्रोग्रेस ट्रेंड ग्राफ"
    },
    "Gujarati": {
        "title": "વિદ્યાર્થી પ્રદર્શન અને પરિણામ પોર્ટલ",
        "search_lbl": "🔎 વિદ્યાર્થીનું પરિણામ જુઓ",
        "cert_btn": "🏆 મેરિટ સર્ટિફિકેટ ડાઉનલોડ કરો",
        "admit_btn": "🪪 એડમિટ કાર્ડ (Admit Card) ડાઉનલોડ કરો",
        "weak_alert": "⚠️ નીચેના વિષયોમાં વિશેષ ધ્યાન આપવાની જરૂર છે (< 33%):",
        "chart_title": "વિષયવાર ગુણ વિશ્લેષણ ગ્રાફ",
        "combined_title": "🌟 તમામ પરીક્ષાઓનું સંયુક્ત પ્રદર્શન (Combined Performance)",
        "trend_title": "📈 મલ્ટિ-પરીક્ષા પ્રગતિ ગ્રાફ"
    }
}

def mask_aadhaar(val):
    return "[Aadhaar Redacted]"

def format_clean_number(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if not val_str:
        return ""
    try:
        f = float(val_str)
        if f.is_integer():
            return str(int(f))
    except (ValueError, TypeError):
        pass
    if '.' in val_str:
        val_str = val_str.split('.')[0]
    return val_str

def clean_val(val):
    s = format_clean_number(val)
    return re.sub(r'[^a-zA-Z0-9]', '', s).lower().strip()

# Enhanced DOB Normalizer (Fixes 26.12.2011, 2011-12-26, 26/12/2011 variations)
def clean_dob_str(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if not s:
        return ""
    formats_to_try = [
        "%d.%m.%Y", "%d-%m-%Y", "%d/%m/%Y", 
        "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
        "%d%m%Y", "%m/%d/%Y"
    ]
    for fmt in formats_to_try:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%Y%m%d")
        except ValueError:
            pass
    return re.sub(r'[^a-zA-Z0-9]', '', s).lower().strip()

def clean_mobile_for_wa(val):
    s = format_clean_number(val)
    digits = re.sub(r'[^0-9]', '', s)
    if len(digits) >= 10:
        return "91" + digits[-10:]
    return digits

def get_exam_priority(exam_name):
    e = str(exam_name).upper().strip()
    if 'TERM END' in e or 'ANNUAL' in e or 'FINAL' in e:
        return 6
    elif 'PWT-4' in e or 'PWT 4' in e or 'PWT4' in e:
        return 5
    elif 'PWT-3' in e or 'PWT 3' in e or 'PWT3' in e:
        return 4
    elif 'TERM-1' in e or 'TERM 1' in e or 'HALF YEARLY' in e or 'MID' in e:
        return 3
    elif 'PWT-2' in e or 'PWT 2' in e or 'PWT2' in e:
        return 2
    elif 'PWT-1' in e or 'PWT 1' in e or 'PWT1' in e:
        return 1
    return 0

@st.cache_data
def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode()
    return None

encoded_bg = get_base64_image(BG_PATH)
if encoded_bg:
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: linear-gradient(rgba(255, 255, 255, 0.90), rgba(255, 255, 255, 0.90)), url("data:image/png;base64,{encoded_bg}");
            background-size: cover;
            background-attachment: fixed;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

def get_and_increment_visits():
    counter_file = "visit_count.txt"
    if not os.path.exists(counter_file):
        with open(counter_file, "w") as f:
            f.write("0")
    with open(counter_file, "r+") as f:
        try:
            count = int(f.read().strip())
        except Exception:
            count = 0
        count += 1
        f.seek(0)
        f.write(str(count))
        f.truncate()
    return count

total_visits = get_and_increment_visits()

# Database Handlers with Auto Migration Feature
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS student_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            Class TEXT, Roll_No TEXT, Student_Name TEXT, Father_Name TEXT,
            DOB TEXT, Aadhaar_No TEXT, Mobile_No TEXT, Exam_Type TEXT,
            Max_Marks REAL, Class_Teacher TEXT, Total_Marks REAL,
            Percentage REAL, Class_Rank INTEGER, Subject_Data TEXT,
            Attendance TEXT, Discipline TEXT, Remarks TEXT
        )
    ''')
    
    cursor.execute("PRAGMA table_info(student_results)")
    existing_columns = [col[1] for col in cursor.fetchall()]
    
    new_cols = [("Attendance", "TEXT"), ("Discipline", "TEXT"), ("Remarks", "TEXT")]
    for col_name, col_type in new_cols:
        if col_name not in existing_columns:
            try:
                cursor.execute(f"ALTER TABLE student_results ADD COLUMN {col_name} {col_type}")
            except Exception:
                pass
                
    conn.commit()
    conn.close()

init_db()

def sync_df_to_sqlite(df):
    if os.path.exists(DB_FILE):
        backup_file = f"backups/backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
        shutil.copyfile(DB_FILE, backup_file)
    
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM student_results")
    for _, row in df.iterrows():
        sub_json = json.dumps({s: row[s] for s in ALL_SUBJECTS if s in row and pd.notna(row[s])})
        cursor.execute('''
            INSERT INTO student_results (Class, Roll_No, Student_Name, Father_Name, DOB, Aadhaar_No, Mobile_No, Exam_Type, Max_Marks, Class_Teacher, Total_Marks, Percentage, Class_Rank, Subject_Data, Attendance, Discipline, Remarks)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            str(row.get('Class', '')), format_clean_number(row.get('Roll_No', '')), str(row.get('Student_Name', '')), str(row.get('Father_Name', '')), 
            str(row.get('DOB', '')), format_clean_number(row.get('Aadhaar_No', '')), format_clean_number(row.get('Mobile_No', '')), str(row.get('Exam_Type', '')), 
            float(row.get('Max_Marks', 600)), str(row.get('Class_Teacher', '')), float(row.get('Total_Marks', 0)), 
            float(row.get('Percentage', 0)), int(row.get('Class_Rank', 0)), sub_json,
            str(row.get('Attendance', '95%')), str(row.get('Discipline', 'A')), str(row.get('Remarks', 'Good Performance'))
        ))
    conn.commit()
    conn.close()

def load_sqlite_to_df():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Class, Roll_No, Student_Name, Father_Name, DOB, Aadhaar_No, Mobile_No, Exam_Type, Max_Marks, Class_Teacher, Total_Marks, Percentage, Class_Rank, Subject_Data, Attendance, Discipline, Remarks FROM student_results")
        rows = cursor.fetchall()
    except sqlite3.OperationalError:
        conn.close()
        init_db()
        return None
    conn.close()
    
    if not rows:
        return None
    
    records = []
    for r in rows:
        rec = {
            'Class': r[0], 
            'Roll_No': format_clean_number(r[1]), 
            'Student_Name': r[2], 
            'Father_Name': r[3],
            'DOB': r[4], 
            'Aadhaar_No': format_clean_number(r[5]), 
            'Mobile_No': format_clean_number(r[6]), 
            'Exam_Type': r[7],
            'Max_Marks': r[8], 
            'Class_Teacher': r[9], 
            'Total_Marks': r[10],
            'Percentage': r[11], 
            'Class_Rank': r[12], 
            'Attendance': r[14] if r[14] else '95%',
            'Discipline': r[15] if r[15] else 'A',
            'Remarks': r[16] if r[16] else 'Good Performance'
        }
        sub_dict = json.loads(r[13]) if r[13] else {}
        for sub in ALL_SUBJECTS:
            rec[sub] = sub_dict.get(sub, np.nan)
        records.append(rec)
    return pd.DataFrame(records)

# Notices Helpers
def load_notices():
    if os.path.exists(NOTICES_FILE):
        with open(NOTICES_FILE, "r") as f:
            return json.load(f)
    return ["Welcome to the PM SHRI JNV Chhotaudepur Student Portal!"]

def save_notices(notices_list):
    with open(NOTICES_FILE, "w") as f:
        json.dump(notices_list, f)

def log_parent_search(roll_no, student_name, selected_class):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_data = pd.DataFrame([{"Timestamp": timestamp, "Roll_No": str(roll_no), "Student_Name": student_name, "Class": selected_class}])
    if os.path.exists(LOG_FILE):
        new_data.to_csv(LOG_FILE, mode='a', header=False, index=False)
    else:
        new_data.to_csv(LOG_FILE, mode='w', header=True, index=False)

def process_data_excel(excel_file_source):
    xls = pd.ExcelFile(excel_file_source)
    sheet_names = xls.sheet_names
    if len(sheet_names) > 1:
        df_basic = pd.read_excel(xls, sheet_name=sheet_names[0])
        df_marks = pd.read_excel(xls, sheet_name=sheet_names[1])
        merge_keys = [c for c in ['Roll_No', 'Class'] if c in df_basic.columns and c in df_marks.columns]
        if not merge_keys:
            merge_keys = ['Roll_No']
        df = pd.merge(df_marks, df_basic, on=merge_keys, how='left', suffixes=('', '_basic'))
    else:
        df = pd.read_excel(xls, sheet_name=sheet_names[0])

    meta_cols = ['Class', 'Roll_No', 'Student_Name', 'Father_Name', 'DOB', 'Aadhaar_No', 'Mobile_No', 'Exam_Type', 'Max_Marks', 'Class_Teacher', 'Attendance', 'Discipline', 'Remarks']
    for col in meta_cols:
        if col not in df.columns:
            df[col] = "95%" if col == 'Attendance' else ("A" if col == 'Discipline' else ("Good" if col == 'Remarks' else ""))

    for col in ['Roll_No', 'Aadhaar_No', 'Mobile_No']:
        if col in df.columns:
            df[col] = df[col].apply(format_clean_number)

    for sub in ALL_SUBJECTS:
        if sub not in df.columns:
            df[sub] = np.nan
        df[sub] = pd.to_numeric(df[sub], errors='coerce')

    df['Total_Marks'] = df[ALL_SUBJECTS].sum(axis=1, skipna=True)
    if 'Max_Marks' not in df.columns or df['Max_Marks'].isnull().all():
        df['Max_Marks'] = df['Exam_Type'].apply(lambda x: 150 if 'PWT' in str(x).upper() else 600)

    df['Max_Marks'] = pd.to_numeric(df['Max_Marks'], errors='coerce').fillna(600)
    df['Percentage'] = (df['Total_Marks'] / df['Max_Marks']) * 100
    df['Percentage'] = df['Percentage'].round(2)
    df['Class_Rank'] = df.groupby(['Class', 'Exam_Type'])['Total_Marks'].rank(ascending=False, method='min').fillna(0).astype(int)
    
    sync_df_to_sqlite(df)
    return df

if "student_data" not in st.session_state or st.session_state["student_data"] is None:
    if os.path.exists(EXCEL_FILE_PATH):
        try:
            st.session_state["student_data"] = process_data_excel(EXCEL_FILE_PATH)
        except Exception as e:
            st.error(f"Excel read error: {e}")
            st.session_state["student_data"] = load_sqlite_to_df()
    else:
        st.session_state["student_data"] = load_sqlite_to_df()

if "admin_logged_in" not in st.session_state:
    st.session_state["admin_logged_in"] = False

def load_board_toppers():
    if os.path.exists(BOARD_TOPPERS_FILE):
        with open(BOARD_TOPPERS_FILE, "r") as f:
            return json.load(f)
    return []

# ReportLab Decorative Canvas Callbacks
def create_watermark_callback(watermark_text, with_border=True):
    def draw_canvas(canvas, doc):
        canvas.saveState()
        if with_border:
            canvas.setStrokeColor(colors.HexColor('#1565C0'))
            canvas.setLineWidth(3)
            canvas.rect(15, 15, doc.pagesize[0]-30, doc.pagesize[1]-30)
            canvas.setLineWidth(1)
            canvas.rect(19, 19, doc.pagesize[0]-38, doc.pagesize[1]-38)

        canvas.setFont('Helvetica-Bold', 36)
        canvas.setFillColor(colors.HexColor('#E0E0E0'), alpha=0.25)
        canvas.translate(doc.pagesize[0] / 2.0, doc.pagesize[1] / 2.0)
        canvas.rotate(35)
        canvas.drawCentredString(0, 0, watermark_text)
        canvas.restoreState()
    return draw_canvas

def generate_merit_certificate_pdf(student_info, exam_type, percentage, rank):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=35, leftMargin=35, topMargin=35, bottomMargin=35)
    story = []
    styles = getSampleStyleSheet()

    cert_title = ParagraphStyle('CertTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=22, leading=26, alignment=1, textColor=colors.HexColor('#1565C0'))
    sub_title = ParagraphStyle('SubTitle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=15, leading=18, alignment=1, textColor=colors.HexColor('#2E7D32'))
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontName='Helvetica', fontSize=13, leading=22, alignment=1)

    if os.path.exists(LOGO_PATH):
        try:
            logo_img = RLImage(LOGO_PATH, width=55, height=55)
            logo_img.hAlign = 'CENTER'
            story.append(logo_img)
            story.append(Spacer(1, 4))
        except Exception:
            pass

    story.append(Paragraph("PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR", cert_title))
    story.append(Spacer(1, 4))
    story.append(Paragraph("🏆 CERTIFICATE OF ACADEMIC EXCELLENCE 🏆", sub_title))
    story.append(Spacer(1, 15))

    cert_text = f"""
    This is to proudly certify that <b>{student_info['Student_Name']}</b>, Son/Daughter of <b>{student_info['Father_Name']}</b>, 
    studying in <b>Class {student_info['Class']}</b> (Roll No: <b>{student_info['Roll_No']}</b>), has secured 
    <font color="#1565C0"><b>RANK #{rank}</b></font> with an outstanding score of <b>{percentage:.2f}%</b> 
    in the <b>{exam_type}</b> Examination (Academic Year 2025-26).
    """
    story.append(Paragraph(cert_text, body_style))
    story.append(Spacer(1, 20))

    seal_element = RLImage(SEAL_PATH, width=50, height=50) if os.path.exists(SEAL_PATH) else Paragraph("<b>[OFFICIAL SEAL]</b>", body_style)
    sign_element = RLImage(SIGN_PATH, width=70, height=35) if os.path.exists(SIGN_PATH) else Paragraph("<b>____________________</b>", body_style)

    sig_data = [
        [Paragraph("<b>____________________</b>", body_style), seal_element, sign_element],
        [Paragraph("<b>Class Teacher</b>", body_style), Paragraph("<b>School Seal</b>", body_style), Paragraph("<b>Principal Signature</b>", body_style)]
    ]
    sig_table = Table(sig_data, colWidths=[240, 180, 240])
    sig_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(sig_table)

    watermark_fn = create_watermark_callback("PM SHRI JNV CHHOTAUDEPUR", with_border=True)
    doc.build(story, onFirstPage=watermark_fn, onLaterPages=watermark_fn)
    buffer.seek(0)
    return buffer.getvalue()

def generate_pdf_scorecard(student_info, filtered_df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=15, leading=18, alignment=1, textColor=colors.HexColor('#1E88E5'))
    subtitle_style = ParagraphStyle('DocSubtitle', parent=styles['Heading3'], fontName='Helvetica-Bold', fontSize=11, leading=14, alignment=1, textColor=colors.HexColor('#333333'))
    normal_style = styles['Normal']

    if os.path.exists(LOGO_PATH):
        try:
            logo_img = RLImage(LOGO_PATH, width=45, height=45)
            logo_img.hAlign = 'CENTER'
            story.append(logo_img)
            story.append(Spacer(1, 4))
        except Exception:
            pass

    story.append(Paragraph("PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR", title_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph("STUDENT ACADEMIC PERFORMANCE REPORT CARD", subtitle_style))
    story.append(Spacer(1, 8))

    info_data = [
        [Paragraph(f"<b>Student Name:</b> {student_info['Student_Name']}", normal_style), Paragraph(f"<b>Roll No:</b> {student_info['Roll_No']}", normal_style)],
        [Paragraph(f"<b>Class:</b> {student_info['Class']}", normal_style), Paragraph(f"<b>Class Teacher:</b> {student_info['Class_Teacher']}", normal_style)],
        [Paragraph(f"<b>Father's Name:</b> {student_info['Father_Name']}", normal_style), Paragraph(f"<b>Aadhaar:</b> [Aadhaar Redacted]", normal_style)],
        [Paragraph(f"<b>Attendance:</b> {student_info.get('Attendance', '95%')}", normal_style), Paragraph(f"<b>Discipline Grade:</b> {student_info.get('Discipline', 'A')}", normal_style)]
    ]
    info_table = Table(info_data, colWidths=[260, 260])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8F9FA')),
        ('PADDING', (0, 0), (-1, -1), 5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E0E0E0')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 10))

    for _, row in filtered_df.iterrows():
        exam_header = f"<b>Exam:</b> {row['Exam_Type']} &nbsp;|&nbsp; <b>Score:</b> {int(row['Total_Marks'])}/{int(row['Max_Marks'])} ({row['Percentage']:.2f}%) &nbsp;|&nbsp; <b>Rank:</b> #{row['Class_Rank']}"
        story.append(Paragraph(exam_header, styles['Heading3']))
        story.append(Spacer(1, 4))

        table_data = [["S.No.", "Subject Name", "Marks Obtained"]]
        active_subs = [s for s in ALL_SUBJECTS if s in row and pd.notna(row[s])]
        for idx, sub_name in enumerate(active_subs, start=1):
            val = row[sub_name]
            val_str = str(int(val)) if pd.notna(val) and float(val).is_integer() else str(val)
            table_data.append([str(idx), str(sub_name), val_str])

        score_table = Table(table_data, colWidths=[50, 320, 150])
        score_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E88E5')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
            ('PADDING', (0, 0), (-1, -1), 4),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('ALIGN', (2, 0), (2, -1), 'CENTER'),
        ]))
        story.append(score_table)
        story.append(Spacer(1, 8))

    story.append(Paragraph(f"<b>Teacher Remarks:</b> {student_info.get('Remarks', 'Keep up the good work!')}", normal_style))
    story.append(Spacer(1, 10))

    seal_element = RLImage(SEAL_PATH, width=45, height=45) if os.path.exists(SEAL_PATH) else Paragraph("<b>[SEAL]</b>", normal_style)
    sign_element = RLImage(SIGN_PATH, width=65, height=30) if os.path.exists(SIGN_PATH) else Paragraph("<b>________________</b>", normal_style)

    sig_data = [
        [Paragraph("<b>____________________</b>", normal_style), seal_element, sign_element],
        [Paragraph("<b>Class Teacher Sign</b>", normal_style), Paragraph("<b>School Seal</b>", normal_style), Paragraph("<b>Principal Signature</b>", normal_style)]
    ]
    sig_table = Table(sig_data, colWidths=[180, 160, 180])
    sig_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(sig_table)

    watermark_fn = create_watermark_callback("PM SHRI JNV CHHOTAUDEPUR", with_border=True)
    doc.build(story, onFirstPage=watermark_fn, onLaterPages=watermark_fn)
    buffer.seek(0)
    return buffer.getvalue()

def generate_admit_card_pdf(student_info):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=16, leading=18, alignment=1, textColor=colors.HexColor('#1565C0'))
    sub_style = ParagraphStyle('DocSub', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=12, leading=14, alignment=1, textColor=colors.HexColor('#2E7D32'))
    normal_style = styles['Normal']

    if os.path.exists(LOGO_PATH):
        try:
            logo_img = RLImage(LOGO_PATH, width=50, height=50)
            logo_img.hAlign = 'CENTER'
            story.append(logo_img)
            story.append(Spacer(1, 4))
        except Exception:
            pass

    story.append(Paragraph("PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR", title_style))
    story.append(Paragraph("EXAMINATION ADMIT CARD / HALL TICKET (2025-26)", sub_style))
    story.append(Spacer(1, 10))

    photo_path = f"photos/students/{student_info['Roll_No']}.png"
    photo_elem = RLImage(photo_path, width=70, height=80) if os.path.exists(photo_path) else Paragraph("<b>[PHOTO]</b>", normal_style)

    info_data = [
        [Paragraph(f"<b>Student Name:</b> {student_info['Student_Name']}", normal_style), photo_elem],
        [Paragraph(f"<b>Roll No:</b> {student_info['Roll_No']}", normal_style), ""],
        [Paragraph(f"<b>Class:</b> {student_info['Class']}", normal_style), ""],
        [Paragraph(f"<b>Father's Name:</b> {student_info['Father_Name']}", normal_style), ""],
        [Paragraph(f"<b>Exam Center:</b> JNV Chhotaudepur Main Campus", normal_style), ""]
    ]
    t = Table(info_data, colWidths=[380, 120])
    t.setStyle(TableStyle([
        ('SPAN', (1, 0), (1, 4)),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
        ('PADDING', (0,0), (-1,-1), 5)
    ]))
    story.append(t)
    story.append(Spacer(1, 15))

    rules_text = """
    <b>Exam Instructions:</b><br/>
    1. Student must carry this Admit Card to the Examination Hall.<br/>
    2. Electronic devices (Mobile, Smart Watches) are strictly prohibited.<br/>
    3. Be present in the exam hall 15 minutes before the scheduled time.
    """
    story.append(Paragraph(rules_text, normal_style))
    story.append(Spacer(1, 20))

    seal_element = RLImage(SEAL_PATH, width=45, height=45) if os.path.exists(SEAL_PATH) else Paragraph("<b>[SEAL]</b>", normal_style)
    sign_element = RLImage(SIGN_PATH, width=65, height=30) if os.path.exists(SIGN_PATH) else Paragraph("<b>________________</b>", normal_style)

    sig_data = [
        [Paragraph("<b>____________________</b>", normal_style), seal_element, sign_element],
        [Paragraph("<b>Student Signature</b>", normal_style), Paragraph("<b>School Seal</b>", normal_style), Paragraph("<b>Principal Signature</b>", normal_style)]
    ]
    sig_table = Table(sig_data, colWidths=[180, 160, 180])
    sig_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(sig_table)

    watermark_fn = create_watermark_callback("PM SHRI JNV CHHOTAUDEPUR", with_border=True)
    doc.build(story, onFirstPage=watermark_fn, onLaterPages=watermark_fn)
    buffer.seek(0)
    return buffer.getvalue()


# Header Layout
h_col1, h_col2, h_col3 = st.columns([1.2, 5.6, 1.2], vertical_alignment="center")
with h_col1:
    if os.path.exists(LOGO_PATH):
        st.markdown(f'<div class="header-logo-left"><img src="data:image/png;base64,{get_base64_image(LOGO_PATH)}" width="80"></div>', unsafe_allow_html=True)
with h_col2:
    st.markdown("<h2 class='main-title'>PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR</h2>", unsafe_allow_html=True)
with h_col3:
    if os.path.exists(LOGO_PATH):
        st.markdown(f'<div class="header-logo-right mobile-hide"><img src="data:image/png;base64,{get_base64_image(LOGO_PATH)}" width="80"></div>', unsafe_allow_html=True)

st.markdown("---")

# Navigation Menu Bar
nav_col1, nav_col2 = st.columns([4, 1.2], vertical_alignment="center")
with nav_col1:
    menu = st.radio(
        "NAVIGATION_MENU",
        ["👨‍🎓 PARENT PORTAL", "🖼️ SCHOOL GALLERY", "🏆 BOARD EXAM RESULTS", "⚙️ ADMIN PORTAL"],
        horizontal=True,
        label_visibility="collapsed"
    )
with nav_col2:
    selected_lang = st.selectbox("🌐 Language / भाषा", ["English", "Hindi", "Gujarati"], label_visibility="collapsed")

txt = LANG_TEXTS[selected_lang]
st.markdown(f"<h4 class='sub-title'>{txt['title']}</h4>", unsafe_allow_html=True)
st.markdown("---")

def render_topper_marquee(topper_list):
    if not topper_list:
        st.info("Top performers details will be displayed here once available.")
        return
    cards_html = ""
    for t in topper_list:
        img_b64 = get_base64_image(t.get("photo", ""))
        img_src = f"data:image/png;base64,{img_b64}" if img_b64 else "https://via.placeholder.com/80?text=Topper"
        card = (
            f'<div class="topper-card">'
            f'<img src="{img_src}" style="width: 75px; height: 75px; border-radius: 50%; object-fit: cover; border: 2px solid #1565C0;">'
            f'<div style="font-weight: bold; color: #0D47A1; margin-top: 5px; font-size: 13px;">{t["name"]}</div>'
            f'<div style="font-size: 11px; color: #333;">Class {t["class"]} ({t.get("year", "2025-26")})</div>'
            f'<div style="font-size: 14px; font-weight: bold; color: #2E7D32; background: #E8F5E9; margin-top: 4px; border-radius: 4px; padding: 2px 0;">🏆 {t["percentage"]}</div>'
            f'</div>'
        )
        cards_html += card
    st.markdown(f'<marquee direction="left" scrollamount="6" onmouseover="this.stop();" onmouseout="this.start();">{cards_html}</marquee>', unsafe_allow_html=True)

# ==============================================================================
# 👨‍🎓 PARENT PORTAL
# ==============================================================================
if menu == "👨‍🎓 PARENT PORTAL":
    # Dynamic Digital Notice Board
    notices = load_notices()
    if notices:
        st.markdown(f"""
            <div class="notice-box">
                <span style="font-weight: bold; color: #1B5E20;">📢 DIGITAL NOTICE BOARD:</span>
                <marquee direction="left" scrollamount="5" behavior="scroll" style="vertical-align: middle; margin-left: 10px;">
                    {" &nbsp;&nbsp;&nbsp;&nbsp; 🔹 &nbsp;&nbsp;&nbsp;&nbsp; ".join(notices)}
                </marquee>
            </div>
        """, unsafe_allow_html=True)

    # Hall of Fame
    st.markdown("<div class='hall-of-fame-box'>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; color: #0D47A1; margin-bottom: 8px;'>🏆 ACADEMIC HALL OF FAME (SCHOOL TOPPERS) 🏆</h4>", unsafe_allow_html=True)
    
    all_toppers_list = load_board_toppers()
    if st.session_state["student_data"] is not None and not st.session_state["student_data"].empty:
        df_top = st.session_state["student_data"]
        for c_val in ['12', '10']:
            c_df = df_top[df_top['Class'].astype(str).str.contains(c_val, na=False)]
            if not c_df.empty:
                top_student = c_df.sort_values(by='Percentage', ascending=False).iloc[0]
                photo_p = f"photos/students/{top_student['Roll_No']}.png"
                all_toppers_list.append({
                    "name": top_student['Student_Name'],
                    "class": str(top_student['Class']),
                    "percentage": f"{top_student['Percentage']:.1f}%",
                    "year": "Current Exam",
                    "photo": photo_p if os.path.exists(photo_p) else ""
                })

    hof_tab12, hof_tab10 = st.tabs(["🎓 Class 12 Toppers", "🎓 Class 10 Toppers"])
    with hof_tab12:
        top_12 = [t for t in all_toppers_list if "12" in str(t.get("class", ""))]
        render_topper_marquee(top_12)
    with hof_tab10:
        top_10 = [t for t in all_toppers_list if "10" in str(t.get("class", ""))]
        render_topper_marquee(top_10)
    st.markdown("</div>", unsafe_allow_html=True)
    
    if st.session_state["student_data"] is not None:
        df_data = st.session_state["student_data"]
        if not df_data.empty and 'Exam_Type' in df_data.columns:
            unique_exams = df_data['Exam_Type'].dropna().unique()
            if len(unique_exams) > 0:
                sorted_exams = sorted(unique_exams, key=get_exam_priority, reverse=True)
                latest_exam = sorted_exams[0]
                latest_df = df_data[df_data['Exam_Type'] == latest_exam].copy()
                if not latest_df.empty:
                    school_topper = latest_df.sort_values(by='Percentage', ascending=False).iloc[0]
                    ticker_items = [f"🏆 <b>OVERALL SCHOOL TOPPER ({latest_exam}):</b> {school_topper['Student_Name']} (Class {school_topper['Class']}) - {school_topper['Percentage']:.2f}%"]
                    
                    classes = sorted(latest_df['Class'].astype(str).unique())
                    for cls in classes:
                        cls_toppers = latest_df[latest_df['Class'].astype(str) == cls].sort_values(by='Percentage', ascending=False).head(3)
                        top_list = [f"{idx+1}. {r['Student_Name']} ({r['Percentage']:.1f}%)" for idx, (_, r) in enumerate(cls_toppers.iterrows())]
                        ticker_items.append(f"🥇 <b>Class {cls} Top 3:</b> {' | '.join(top_list)}")
                        
                    st.markdown(f"""
                        <div style="background-color: #FFF9C4; border-left: 5px solid #FBC02D; padding: 7px 10px; border-radius: 4px; color: #000; font-size: 15px; margin-bottom: 10px;">
                            <marquee direction="left" scrollamount="6" behavior="scroll">{" &nbsp;&nbsp;&nbsp; ✦ &nbsp;&nbsp;&nbsp; ".join(ticker_items)}</marquee>
                        </div>
                    """, unsafe_allow_html=True)

    st.header(txt['search_lbl'])
    
    if st.session_state["student_data"] is None:
        st.warning("⚠️ Data file not found. Kripya Admin Portal se Data Upload karein.")
    else:
        df = st.session_state["student_data"]
        search_method = st.radio("Choose Verification Method:", [
            "Option 1: Roll No & Date of Birth (DOB)", 
            "Option 2: Roll No & Aadhaar Number", 
            "Option 3: OTP Based Mobile Verification (SMS/WhatsApp)"
        ], horizontal=True)
        
        filtered_df = pd.DataFrame()
        
        if "Option 3" in search_method:
            c1, c2, c3 = st.columns([2, 2, 1])
            with c1:
                selected_class = st.selectbox("Select Class", sorted(df['Class'].astype(str).unique()), key="otp_cls")
                roll_no = st.text_input("Roll No", key="otp_roll")
            with c2:
                mobile_input = st.text_input("Registered Mobile No", key="otp_mob")
            with c3:
                st.write(" ")
                st.write(" ")
                if st.button("📲 Send OTP"):
                    if roll_no and mobile_input:
                        gen_otp = str(random.randint(1000, 9999))
                        st.session_state["current_otp"] = gen_otp
                        st.info(f"🔑 [DEMO OTP]: Your OTP for Verification is **{gen_otp}**")
                    else:
                        st.error("Enter Details first!")
            
            user_otp = st.text_input("Enter 4-Digit OTP Received", key="entered_otp")
            if st.button("🔍 Verify & View Result"):
                if "current_otp" in st.session_state and user_otp == st.session_state["current_otp"]:
                    filtered_df = df[
                        (df['Class'].astype(str).str.strip().str.lower() == selected_class.strip().lower()) &
                        (df['Roll_No'].apply(clean_val) == clean_val(roll_no)) &
                        (df['Mobile_No'].apply(clean_val) == clean_val(mobile_input))
                    ]
                else:
                    st.error("❌ Incorrect OTP entered!")
        else:
            with st.form("search_form"):
                c1, c2 = st.columns(2)
                with c1:
                    selected_class = st.selectbox("Select Class", sorted(df['Class'].astype(str).unique()))
                    roll_no = st.text_input("Roll No")
                
                if "Option 1" in search_method:
                    with c2:
                        dob_input = st.text_input("Date of Birth (e.g. 26.12.2011, 26/12/2011 or 2011-12-26)")
                else:
                    with c2:
                        aadhaar_input = st.text_input("Aadhaar Number")
                
                submit_btn = st.form_submit_button("🔍 View Result")

            if submit_btn:
                if "Option 1" in search_method:
                    filtered_df = df[
                        (df['Class'].astype(str).str.strip().str.lower() == selected_class.strip().lower()) &
                        (df['Roll_No'].apply(clean_val) == clean_val(roll_no)) &
                        (df['DOB'].apply(clean_dob_str) == clean_dob_str(dob_input))
                    ]
                else:
                    filtered_df = df[
                        (df['Class'].astype(str).str.strip().str.lower() == selected_class.strip().lower()) &
                        (df['Roll_No'].apply(clean_val) == clean_val(roll_no)) &
                        (df['Aadhaar_No'].apply(clean_val) == clean_val(aadhaar_input))
                    ]
            
        if not filtered_df.empty:
            student_info = filtered_df.iloc[0]
            log_parent_search(student_info['Roll_No'], student_info['Student_Name'], student_info['Class'])
            
            pdf_bytes = generate_pdf_scorecard(student_info, filtered_df)
            admit_bytes = generate_admit_card_pdf(student_info)

            st.success(f"🎓 Result Found for: **{student_info['Student_Name']}**")
            
            r_col1, r_col2 = st.columns([1, 4])
            with r_col1:
                photo_file = f"photos/students/{student_info['Roll_No']}.png"
                if os.path.exists(photo_file):
                    st.image(photo_file, width=130)
                else:
                    st.info("📷 No Photo")
            
            with r_col2:
                st.write(f"**Student:** {student_info['Student_Name']} | **Roll No:** {student_info['Roll_No']}")
                st.write(f"**Class:** {student_info['Class']} | **Aadhaar:** {mask_aadhaar(student_info['Aadhaar_No'])}")
                st.write(f"**Attendance:** {student_info.get('Attendance', '95%')} | **Discipline:** Grade {student_info.get('Discipline', 'A')}")
                
                b_c1, b_c2, b_c3 = st.columns(3)
                with b_c1:
                    st.download_button("📥 Report Card (PDF)", data=pdf_bytes, file_name=f"Report_{student_info['Roll_No']}.pdf", mime="application/pdf", use_container_width=True)
                with b_c2:
                    st.download_button(txt['admit_btn'], data=admit_bytes, file_name=f"AdmitCard_{student_info['Roll_No']}.pdf", mime="application/pdf", use_container_width=True)
                
                best_row = filtered_df.sort_values('Class_Rank').iloc[0]
                if best_row['Class_Rank'] <= 3:
                    with b_c3:
                        cert_pdf = generate_merit_certificate_pdf(student_info, best_row['Exam_Type'], best_row['Percentage'], best_row['Class_Rank'])
                        st.download_button(txt['cert_btn'], data=cert_pdf, file_name=f"Merit_Certificate_{student_info['Roll_No']}.pdf", mime="application/pdf", use_container_width=True)

            st.markdown("---")
            st.subheader(txt['combined_title'])
            tot_obtained = filtered_df['Total_Marks'].sum()
            tot_max = filtered_df['Max_Marks'].sum()
            overall_pct = (tot_obtained / tot_max * 100) if tot_max > 0 else 0.0
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Exams Taken", len(filtered_df))
            m2.metric("Total Marks Obtained", f"{int(tot_obtained)}")
            m3.metric("Combined Maximum Marks", f"{int(tot_max)}")
            m4.metric("Overall Percentage", f"{overall_pct:.2f}%")

            st.markdown("---")
            st.subheader(txt['trend_title'])
            
            if len(filtered_df) > 0:
                trend_fig = go.Figure()
                trend_fig.add_trace(go.Scatter(
                    x=filtered_df['Exam_Type'], 
                    y=filtered_df['Percentage'],
                    mode='lines+markers+text',
                    name='Percentage',
                    text=[f"{p:.1f}%" for p in filtered_df['Percentage']],
                    textposition="top center",
                    line=dict(color='#1E88E5', width=3),
                    marker=dict(size=10, color='#1565C0')
                ))
                trend_fig.update_layout(
                    title=f"Academic Progress Timeline for {student_info['Student_Name']}",
                    xaxis_title="Examinations",
                    yaxis_title="Percentage Score (%)",
                    yaxis=dict(range=[0, 105]),
                    margin=dict(l=20, r=20, t=40, b=20)
                )
                st.plotly_chart(trend_fig, use_container_width=True)

            st.markdown("---")
            st.subheader(f"📊 {txt['chart_title']}")
            latest_row = filtered_df.iloc[-1]
            sub_names = [s for s in ALL_SUBJECTS if s in latest_row and pd.notna(latest_row[s])]
            sub_marks = [latest_row[s] for s in sub_names]
            
            class_df = df[(df['Class'].astype(str) == str(latest_row['Class'])) & (df['Exam_Type'] == latest_row['Exam_Type'])]
            class_avgs = [class_df[s].mean() for s in sub_names]

            fig = go.Figure()
            fig.add_trace(go.Bar(x=sub_names, y=sub_marks, name="Student Score", marker_color='#1E88E5', text=sub_marks, textposition='auto'))
            fig.add_trace(go.Bar(x=sub_names, y=class_avgs, name="Class Average", marker_color='#FFA726', text=[f"{v:.1f}" for v in class_avgs], textposition='auto'))
            
            fig.update_layout(barmode='group', title=f"Marks Comparison vs Class Average ({latest_row['Exam_Type']})", xaxis_title="Subjects", yaxis_title="Marks", yaxis=dict(range=[0, 100]))
            st.plotly_chart(fig, use_container_width=True)

            weak_subs = [f"{s} ({latest_row[s]} marks)" for s in sub_names if float(latest_row[s]) < 33.0]
            if weak_subs:
                st.markdown(f"<div class='weak-badge'>{txt['weak_alert']} {', '.join(weak_subs)}</div>", unsafe_allow_html=True)

            st.markdown("---")
            st.subheader("📝 INDIVIDUAL EXAM SCORECARDS")
            for index, row in filtered_df.iterrows():
                with st.expander(f"📌 **{row['Exam_Type']}** | Score: {int(row['Total_Marks'])}/{int(row['Max_Marks'])} ({row['Percentage']:.2f}%) | Rank: #{row['Class_Rank']}", expanded=True):
                    subject_rows = []
                    s_no = 1
                    for sub_name in ALL_SUBJECTS:
                        if pd.notna(row[sub_name]):
                            val = row[sub_name]
                            subject_rows.append({'S.No.': s_no, 'Subject Name': sub_name, 'Marks Obtained': int(val) if float(val).is_integer() else val})
                            s_no += 1
                    st.dataframe(pd.DataFrame(subject_rows), hide_index=True, use_container_width=True)
        elif submit_btn:
            st.error("❌ No student record found matching the provided Class, Roll Number, and Credentials. Kripya details re-check karein.")

# ==============================================================================
# 🖼️ GALLERY & BOARD RESULTS
# ==============================================================================
elif menu == "🖼️ SCHOOL GALLERY":
    st.header("🏫 GALLERY")
    gallery_files = [f for f in os.listdir("photos/gallery") if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
    if not gallery_files:
        st.info("ℹ️ Gallery empty.")
    else:
        images_html = "".join([f'<img src="data:image/png;base64,{get_base64_image(os.path.join("photos/gallery", img))}" style="height: 200px; margin-right: 15px; border-radius: 8px; border: 2px solid #1E88E5;">' for img in gallery_files])
        st.markdown(f'<marquee direction="left" scrollamount="7">{images_html}</marquee>', unsafe_allow_html=True)

elif menu == "🏆 BOARD EXAM RESULTS":
    st.header("🎓 CBSE BOARD TOPPERS HALL OF FAME")
    toppers_data = load_board_toppers()
    if not toppers_data:
        st.info("No toppers uploaded yet. Add toppers from Admin Dashboard.")
    else:
        tab12, tab10 = st.tabs(["🎓 Class 12 Toppers", "🎓 Class 10 Toppers"])
        with tab12:
            t12_list = [t for t in toppers_data if "12" in str(t.get("class", ""))]
            render_topper_marquee(t12_list)

        with tab10:
            t10_list = [t for t in toppers_data if "10" in str(t.get("class", ""))]
            render_topper_marquee(t10_list)

# ==============================================================================
# ⚙️ ADMIN PORTAL
# ==============================================================================
elif menu == "⚙️ ADMIN PORTAL":
    st.header("🔒 ADMIN DASHBOARD")
    st.info(f"👁️ **Total Website Visits:** `{total_visits}`")
    
    if not st.session_state["admin_logged_in"]:
        with st.form("login_form"):
            st.subheader("🔐 Secure Admin Login")
            admin_user = st.text_input("Username")
            admin_pass = st.text_input("Password", type="password")
            if st.form_submit_button("Login"):
                hashed_input = hashlib.sha256(admin_pass.encode()).hexdigest()
                if admin_user == "admin" and hashed_input == ADMIN_PASS_HASH:
                    st.session_state["admin_logged_in"] = True
                    st.success("✅ Logged In Successfully!")
                    st.rerun()
                else:
                    st.error("❌ Invalid Admin Credentials!")
    else:
        st.success("🔓 Admin Logged In (Secure Session Active)")
        if st.button("Logout"):
            st.session_state["admin_logged_in"] = False
            st.rerun()

        st.markdown("---")

        with st.expander("📦 1. BULK CLASS RESULT PDF EXPORT & EXCEL MERIT LIST", expanded=False):
            if st.session_state["student_data"] is not None:
                df_bulk = st.session_state["student_data"]
                c_col1, c_col2 = st.columns(2)
                with c_col1:
                    bulk_cls = st.selectbox("Select Class", sorted(df_bulk['Class'].astype(str).unique()), key="bulk_cls")
                with c_col2:
                    bulk_exam = st.selectbox("Select Exam Type", sorted(df_bulk['Exam_Type'].astype(str).unique()), key="bulk_exam")

                ex_col1, ex_col2 = st.columns(2)
                with ex_col1:
                    if st.button("🚀 Generate Bulk ZIP File"):
                        filtered_bulk = df_bulk[(df_bulk['Class'].astype(str) == str(bulk_cls)) & (df_bulk['Exam_Type'] == bulk_exam)]
                        if filtered_bulk.empty:
                            st.warning("No records found.")
                        else:
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                                for idx, s_row in filtered_bulk.iterrows():
                                    s_df = filtered_bulk[filtered_bulk['Roll_No'] == s_row['Roll_No']]
                                    pdf_data = generate_pdf_scorecard(s_row, s_df)
                                    pdf_filename = f"Class_{bulk_cls}_{s_row['Roll_No']}_{s_row['Student_Name'].replace(' ', '_')}.pdf"
                                    zip_file.writestr(pdf_filename, pdf_data)
                            zip_buffer.seek(0)
                            st.download_button(
                                label=f"📥 Download ZIP (Class {bulk_cls})",
                                data=zip_buffer,
                                file_name=f"Class_{bulk_cls}_{bulk_exam}_ReportCards.zip",
                                mime="application/zip",
                                use_container_width=True
                            )
                with ex_col2:
                    filtered_bulk_ex = df_bulk[(df_bulk['Class'].astype(str) == str(bulk_cls)) & (df_bulk['Exam_Type'] == bulk_exam)].sort_values('Class_Rank')
                    if not filtered_bulk_ex.empty:
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            filtered_bulk_ex.to_excel(writer, sheet_name=f"Class_{bulk_cls}_Merit", index=False)
                        excel_buffer.seek(0)
                        st.download_button(
                            label=f"📊 Export Merit List (Excel)",
                            data=excel_buffer,
                            file_name=f"MeritList_Class_{bulk_cls}_{bulk_exam}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

        with st.expander("📊 2. TEACHER-WISE & ADVANCED SUBJECT ANALYTICS", expanded=False):
            if st.session_state["student_data"] is not None:
                df_t = st.session_state["student_data"]
                t_tab1, t_tab2 = st.tabs(["👩‍🏫 Class Teacher Summary", "📈 Subject-Wise Analytics"])
                
                with t_tab1:
                    if 'Class_Teacher' in df_t.columns and not df_t['Class_Teacher'].isnull().all():
                        teachers = sorted(df_t['Class_Teacher'].dropna().astype(str).unique())
                        teacher_summary = []
                        for t in teachers:
                            t_df = df_t[df_t['Class_Teacher'].astype(str) == t]
                            tot_students = len(t_df)
                            passed_students = len(t_df[t_df['Percentage'] >= 33.0])
                            pass_pct = (passed_students / tot_students * 100) if tot_students > 0 else 0.0
                            avg_score = t_df['Percentage'].mean()
                            teacher_summary.append({
                                "Class Teacher": t,
                                "Class(es) Assigned": ", ".join(t_df['Class'].astype(str).unique()),
                                "Total Students": tot_students,
                                "Passed Students": passed_students,
                                "Pass Percentage (%)": f"{pass_pct:.2f}%",
                                "Average Class Score (%)": f"{avg_score:.2f}%"
                            })
                        st.dataframe(pd.DataFrame(teacher_summary), hide_index=True, use_container_width=True)
                
                with t_tab2:
                    sub_stats = []
                    for s in ALL_SUBJECTS:
                        if s in df_t.columns and not df_t[s].dropna().empty:
                            s_data = df_t[s].dropna()
                            sub_stats.append({
                                "Subject": s,
                                "Highest Score": s_data.max(),
                                "Lowest Score": s_data.min(),
                                "Average Score": round(s_data.mean(), 2),
                                "Pass % (>=33)": f"{round((s_data[s_data >= 33].count() / len(s_data)) * 100, 1)}%"
                            })
                    st.dataframe(pd.DataFrame(sub_stats), hide_index=True, use_container_width=True)

        with st.expander("✏️ 3. EDIT STUDENT DATA, RANKS & BULK PHOTO UPLOAD", expanded=False):
            if st.session_state["student_data"] is not None:
                st.markdown("##### 📷 Single Student Photo Upload")
                up_p_col1, up_p_col2 = st.columns([2, 2])
                with up_p_col1:
                    roll_to_photo = st.selectbox("Select Student Roll No", sorted(st.session_state["student_data"]['Roll_No'].astype(str).unique()), key="photo_roll_sel")
                with up_p_col2:
                    stu_photo_file = st.file_uploader("Upload Single Photo", type=["png", "jpg", "jpeg"], key="stu_photo_up")
                    if st.button("🖼️ Save Photo"):
                        if stu_photo_file and roll_to_photo:
                            Image.open(stu_photo_file).save(f"photos/students/{roll_to_photo}.png")
                            st.success(f"✅ Photo uploaded for Roll No: {roll_to_photo}!")

                st.markdown("---")
                # NEW FEATURE 4: BULK STUDENT PHOTO UPLOAD (ZIP EXTRACTOR)
                st.markdown("##### 📦 Bulk Student Photo Upload (ZIP Extractor)")
                st.info("Upload a `.zip` file containing images named by Roll Number (e.g., `101.png`, `102.jpg`).")
                zip_photo_file = st.file_uploader("Upload Photos Archive (.zip)", type=["zip"], key="zip_photos_up")
                if st.button("🚀 Process & Extract ZIP Photos"):
                    if zip_photo_file:
                        try:
                            extracted_count = 0
                            with zipfile.ZipFile(zip_photo_file, 'r') as z:
                                for file_info in z.infolist():
                                    if not file_info.is_dir():
                                        ext = os.path.splitext(file_info.filename)[1].lower()
                                        if ext in ['.png', '.jpg', '.jpeg']:
                                            raw_name = os.path.basename(file_info.filename)
                                            roll_stem = os.path.splitext(raw_name)[0]
                                            clean_roll = format_clean_number(roll_stem)
                                            if clean_roll:
                                                out_path = f"photos/students/{clean_roll}.png"
                                                img_data = z.read(file_info.filename)
                                                img = Image.open(io.BytesIO(img_data))
                                                img.save(out_path)
                                                extracted_count += 1
                            st.success(f"✅ Successfully extracted and linked {extracted_count} student photo(s)!")
                        except Exception as ex:
                            st.error(f"❌ Error extracting ZIP: {ex}")

                st.markdown("---")
                st.markdown("##### 📝 Realtime Data & Rank Modifier")
                
                r_calc_col1, r_calc_col2 = st.columns([2, 2])
                with r_calc_col1:
                    if st.button("🔄 Auto-Recalculate Class Ranks"):
                        df_mod = st.session_state["student_data"].copy()
                        df_mod['Total_Marks'] = df_mod[ALL_SUBJECTS].sum(axis=1, skipna=True)
                        df_mod['Percentage'] = ((df_mod['Total_Marks'] / df_mod['Max_Marks']) * 100).round(2)
                        df_mod['Class_Rank'] = df_mod.groupby(['Class', 'Exam_Type'])['Total_Marks'].rank(ascending=False, method='min').fillna(0).astype(int)
                        st.session_state["student_data"] = df_mod
                        sync_df_to_sqlite(df_mod)
                        st.success("✅ Class ranks recalculated & saved!")
                        st.rerun()

                edited_df = st.data_editor(st.session_state["student_data"], num_rows="dynamic", use_container_width=True, key="db_realtime_editor")

                if "show_save_confirm" not in st.session_state:
                    st.session_state["show_save_confirm"] = False

                if st.button("💾 Save Edits to SQLite Database"):
                    st.session_state["show_save_confirm"] = True

                if st.session_state.get("show_save_confirm", False):
                    st.warning("❓ **ARE YOU SURE YOU WANT TO UPDATE THE DATABASE?**")
                    confirm_col1, confirm_col2 = st.columns(2)
                    with confirm_col1:
                        if st.button("✅ YES, SAVE DATA", use_container_width=True):
                            for sub in ALL_SUBJECTS:
                                if sub in edited_df.columns:
                                    edited_df[sub] = pd.to_numeric(edited_df[sub], errors='coerce')
                            edited_df['Total_Marks'] = edited_df[ALL_SUBJECTS].sum(axis=1, skipna=True)
                            edited_df['Percentage'] = ((edited_df['Total_Marks'] / edited_df['Max_Marks']) * 100).round(2)
                            
                            st.session_state["student_data"] = edited_df
                            sync_df_to_sqlite(edited_df)
                            st.session_state["show_save_confirm"] = False
                            st.success("✅ Database updated successfully!")
                            st.rerun()
                    with confirm_col2:
                        if st.button("❌ NO, CANCEL", use_container_width=True):
                            st.session_state["show_save_confirm"] = False
                            st.info("Update cancelled.")
                            st.rerun()

        with st.expander("📢 4. DIGITAL NOTICE BOARD MANAGEMENT", expanded=False):
            current_notices = load_notices()
            st.subheader("Add New Notice Announcement")
            new_notice = st.text_input("Notice Text")
            if st.button("➕ Post Notice") and new_notice:
                current_notices.insert(0, new_notice)
                save_notices(current_notices)
                st.success("✅ Notice posted!")
                st.rerun()

            st.markdown("---")
            st.write("**Active Notices:**")
            for n_idx, n_text in enumerate(current_notices):
                nc1, nc2 = st.columns([5, 1])
                nc1.write(f"🔹 {n_text}")
                if nc2.button("🗑️ Delete", key=f"del_notice_{n_idx}"):
                    current_notices.pop(n_idx)
                    save_notices(current_notices)
                    st.rerun()

        with st.expander("✒️ 5. DIGITAL SEAL & SIGNATURES MANAGEMENT", expanded=False):
            s_col1, s_col2 = st.columns(2)
            with s_col1:
                st.subheader("Upload School Stamp / Seal")
                seal_file = st.file_uploader("Upload Official Seal", type=["png", "jpg", "jpeg"], key="seal_up")
                if st.button("Save School Seal") and seal_file:
                    Image.open(seal_file).save(SEAL_PATH)
                    st.success("✅ School Seal updated!")
            
            with s_col2:
                st.subheader("Upload Principal Signature")
                sign_file = st.file_uploader("Upload Digital Sign", type=["png", "jpg", "jpeg"], key="sign_up")
                if st.button("Save Principal Sign") and sign_file:
                    Image.open(sign_file).save(SIGN_PATH)
                    st.success("✅ Principal Signature updated!")

        with st.expander("🖼️ 6. SCHOOL GALLERY MANAGEMENT", expanded=False):
            gallery_upload = st.file_uploader("Upload Image to Gallery", type=["png", "jpg", "jpeg"], key="gal_upload")
            if st.button("➕ Add Image to Gallery") and gallery_upload:
                gal_path = os.path.join("photos/gallery", gallery_upload.name)
                Image.open(gallery_upload).save(gal_path)
                st.success("✅ Gallery image added successfully!")
                st.rerun()

            gallery_files = [f for f in os.listdir("photos/gallery") if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            if gallery_files:
                st.write("**Current Gallery Photos:**")
                cols = st.columns(4)
                for idx, g_file in enumerate(gallery_files):
                    with cols[idx % 4]:
                        g_path = os.path.join("photos/gallery", g_file)
                        st.image(g_path, use_container_width=True)
                        if st.button(f"🗑️ Delete", key=f"del_gal_{idx}"):
                            os.remove(g_path)
                            st.success(f"Deleted {g_file}")
                            st.rerun()

        with st.expander("🏆 7. CBSE TOPPERS MANAGEMENT", expanded=False):
            st.subheader("Add New Board Topper")
            b_class = st.selectbox("Class", ["Class 12", "Class 10"])
            b_name = st.text_input("Name")
            b_percent = st.text_input("Percentage (e.g. 98.4%)")
            b_year = st.text_input("Year", value="2025-26")
            b_photo = st.file_uploader("Photo", type=["jpg", "png", "jpeg"])
            if st.button("Add Board Topper") and b_name and b_photo:
                photo_file = f"photos/board/{b_class.replace(' ', '_')}_{clean_val(b_name)}.png"
                Image.open(b_photo).save(photo_file)
                toppers = load_board_toppers()
                toppers.append({"class": b_class, "name": b_name, "percentage": b_percent, "year": b_year, "photo": photo_file})
                with open(BOARD_TOPPERS_FILE, "w") as f:
                    json.dump(toppers, f)
                st.success("✅ Board Topper Added!")
                st.rerun()

            toppers_list = load_board_toppers()
            if toppers_list:
                st.markdown("---")
                st.write("**Current Board Toppers:**")
                for idx, t in enumerate(toppers_list):
                    t_col1, t_col2 = st.columns([4, 1])
                    with t_col1:
                        st.write(f"🏆 **{t['name']}** ({t['class']}, {t['year']}) - {t['percentage']}")
                    with t_col2:
                        if st.button(f"🗑️ Remove", key=f"del_top_{idx}"):
                            if t.get("photo") and os.path.exists(t["photo"]):
                                try:
                                    os.remove(t["photo"])
                                except Exception:
                                    pass
                            toppers_list.pop(idx)
                            with open(BOARD_TOPPERS_FILE, "w") as f:
                                json.dump(toppers_list, f)
                            st.rerun()

        with st.expander("🎨 8. BRANDING & BACKGROUND MANAGEMENT", expanded=False):
            st.subheader("Logo Management")
            up_logo = st.file_uploader("Upload Logo", type=["png", "jpg", "jpeg"], key="logo_up")
            if st.button("Save Logo") and up_logo:
                Image.open(up_logo).save(LOGO_PATH)
                st.success("✅ Logo Updated!")

            st.markdown("---")
            st.subheader("Background Image Management")
            bg_upload = st.file_uploader("Upload Background Image", type=["png", "jpg", "jpeg"], key="bg_up")
            if st.button("🖼️ Set Background") and bg_upload:
                Image.open(bg_upload).save(BG_PATH)
                st.success("✅ Background image updated!")
                st.rerun()

            if os.path.exists(BG_PATH):
                if st.button("🗑️ Remove Background Image"):
                    os.remove(BG_PATH)
                    st.success("✅ Background image removed!")
                    st.rerun()

        # UPDATED FEATURE 3: AUTOMATED WHATSAPP API & MANUAL LINK GENERATOR
        with st.expander("📲 9. AUTOMATED WHATSAPP API & NOTIFICATIONS", expanded=False):
            if st.session_state["student_data"] is not None:
                df_notif = st.session_state["student_data"]
                if 'Mobile_No' in df_notif.columns:
                    n_cls = st.selectbox("Select Class", sorted(df_notif['Class'].astype(str).unique()), key="wa_cls")
                    filtered_notif = df_notif[df_notif['Class'].astype(str) == str(n_cls)]
                    msg_template = st.text_area("Message Content", "Dear Parent, your child's exam results are live on the portal. Check now!")
                    
                    wa_mode = st.radio("Dispatch Mode:", ["Automated Background API Call (Twilio/Meta)", "Manual wa.me Links"], horizontal=True)

                    if wa_mode == "Automated Background API Call (Twilio/Meta)":
                        st.caption("Configure your Cloud API Endpoint & Token to send background SMS/WhatsApp messages automatically.")
                        api_endpoint = st.text_input("API Endpoint URL", value="https://api.twilio.com/2010-04-01/Accounts/YOUR_ACCOUNT_SID/Messages.json")
                        api_token = st.text_input("API Key / Bearer Token", type="password")
                        sender_id = st.text_input("Sender ID / WhatsApp Number", value="+14155238886")

                        if st.button("🚀 Send Automated WhatsApp Messages"):
                            sent_count = 0
                            for _, row in filtered_notif.iterrows():
                                mob = clean_mobile_for_wa(row['Mobile_No'])
                                if len(mob) >= 10:
                                    msg_body = f"Hello {row['Student_Name']},\n\n{msg_template}\nTotal Score: {row['Total_Marks']} ({row['Percentage']}%)"
                                    try:
                                        payload = {"To": f"whatsapp:+{mob}", "From": f"whatsapp:{sender_id}", "Body": msg_body}
                                        headers = {"Authorization": f"Bearer {api_token}"}
                                        # Background API dispatch call
                                        res = requests.post(api_endpoint, data=payload, headers=headers, timeout=5)
                                        sent_count += 1
                                    except Exception:
                                        pass
                            st.success(f"✅ Processed automated dispatch for {sent_count} student(s)!")
                    else:
                        if st.button("🚀 Generate WhatsApp Dispatch Links"):
                            for _, row in filtered_notif.iterrows():
                                mob = clean_mobile_for_wa(row['Mobile_No'])
                                if len(mob) >= 10:
                                    msg_body = f"Hello {row['Student_Name']},\n\n{msg_template}\nTotal Score: {row['Total_Marks']} ({row['Percentage']}%)"
                                    encoded_msg = urllib.parse.quote(msg_body)
                                    wa_link = f"https://api.whatsapp.com/send?phone={mob}&text={encoded_msg}"
                                    st.markdown(f"👉 **{row['Student_Name']}** ({row['Roll_No']}) -> [Click to Send WhatsApp Alert]({wa_link})")
                else:
                    st.error("Mobile_No column missing.")

        with st.expander("📤 10. EXCEL DATA UPLOAD & SQLITE SYNC", expanded=False):
            uploaded_file = st.file_uploader("Upload Excel Sheet (.xlsx)", type=["xlsx", "xls"])
            if st.button("Process & Sync Database") and uploaded_file:
                with open(EXCEL_FILE_PATH, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                st.session_state["student_data"] = process_data_excel(EXCEL_FILE_PATH)
                st.success("Excel & SQLite Database Successfully Updated!")

        # NEW FEATURE 1: AUDIT LOG VIEWER IN ADMIN UI
        with st.expander("📋 11. AUDIT LOG VIEWER (PARENT SEARCH LOGS)", expanded=False):
            if os.path.exists(LOG_FILE):
                log_df = pd.read_csv(LOG_FILE)
                st.dataframe(log_df, use_container_width=True)
                
                search_term = st.text_input("Filter Log by Name or Roll No:", key="log_search")
                if search_term:
                    filtered_logs = log_df[
                        log_df['Student_Name'].astype(str).str.contains(search_term, case=False, na=False) |
                        log_df['Roll_No'].astype(str).str.contains(search_term, case=False, na=False)
                    ]
                    st.write("**Filtered Search Results:**")
                    st.dataframe(filtered_logs, use_container_width=True)

                csv_bytes = log_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Export Audit Logs (CSV)",
                    data=csv_bytes,
                    file_name=f"Result_Search_AuditLogs_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv"
                )
            else:
                st.info("ℹ️ No parent search logs recorded yet.")

        # NEW FEATURE 2: DATABASE BACKUP & RESTORE MANAGER
        with st.expander("🗄️ 12. DATABASE BACKUP & RESTORE MANAGER", expanded=False):
            b_col1, b_col2 = st.columns(2)
            with b_col1:
                st.subheader("📥 Download DB Snapshot")
                if os.path.exists(DB_FILE):
                    with open(DB_FILE, "rb") as f:
                        db_bytes = f.read()
                    st.download_button(
                        label="💾 Download Current SQLite DB (.db)",
                        data=db_bytes,
                        file_name=f"school_database_snapshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db",
                        mime="application/x-sqlite3",
                        use_container_width=True
                    )
                else:
                    st.warning("No Database file found.")

            with b_col2:
                st.subheader("📤 Restore DB Snapshot")
                restored_db_file = st.file_uploader("Upload Backup Database File (.db)", type=["db"], key="restore_db_file")
                if st.button("⚠️ Restore Database Snapshot", use_container_width=True):
                    if restored_db_file:
                        if os.path.exists(DB_FILE):
                            shutil.copyfile(DB_FILE, f"backups/pre_restore_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
                        with open(DB_FILE, "wb") as f:
                            f.write(restored_db_file.getbuffer())
                        st.session_state["student_data"] = load_sqlite_to_df()
                        st.success("✅ Database Snapshot successfully restored & loaded!")
                        st.rerun()

# Footer Text
st.markdown("---")
st.markdown("""
    <div style='text-align: center; color: #555555; padding: 12px; font-size: 14px;'>
        <b>© 2026 PM SHRI JNV CHHOTAUDEPUR | Designed & Developed by <i>Anil Chaudhary</i></b>
    </div>
""", unsafe_allow_html=True)