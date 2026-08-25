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
CBSE_LOGO_PATH = "photos/system/cbse_logo.png"
SEAL_PATH = "photos/system/seal.png"
SIGN_PATH = "photos/system/signature.png"
NOTICES_FILE = "notices.json"
BOARD_TOPPERS_FILE = "board_toppers.json"
LOG_FILE = "result_logs.csv"

# Session State Initialization for Signature & Stamp Offsets
if "sig_width" not in st.session_state:
    st.session_state["sig_width"] = 80
if "sig_height" not in st.session_state:
    st.session_state["sig_height"] = 35
if "seal_width" not in st.session_state:
    st.session_state["seal_width"] = 55
if "seal_height" not in st.session_state:
    st.session_state["seal_height"] = 55
if "sig_y_offset" not in st.session_state:
    st.session_state["sig_y_offset"] = 0

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

def calculate_grade(pct):
    if pct is None or np.isnan(pct):
        return "-"
    if pct >= 91:
        return "A1"
    elif pct >= 81:
        return "A2"
    elif pct >= 71:
        return "B1"
    elif pct >= 61:
        return "B2"
    elif pct >= 51:
        return "C1"
    elif pct >= 41:
        return "C2"
    elif pct >= 33:
        return "D"
    else:
        return "E"

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

# Database Handlers
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS student_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            Class TEXT, Roll_No TEXT, Student_Name TEXT, Father_Name TEXT, Mother_Name TEXT,
            DOB TEXT, Aadhaar_No TEXT, Mobile_No TEXT, Exam_Type TEXT,
            Max_Marks REAL, Class_Teacher TEXT, Total_Marks REAL,
            Percentage REAL, Class_Rank INTEGER, Subject_Data TEXT,
            Attendance TEXT, Discipline TEXT, Remarks TEXT, School_Name TEXT
        )
    ''')
    
    cursor.execute("PRAGMA table_info(student_results)")
    existing_columns = [col[1] for col in cursor.fetchall()]
    
    new_cols = [("Attendance", "TEXT"), ("Discipline", "TEXT"), ("Remarks", "TEXT"), ("Mother_Name", "TEXT"), ("School_Name", "TEXT")]
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
            INSERT INTO student_results (Class, Roll_No, Student_Name, Father_Name, Mother_Name, DOB, Aadhaar_No, Mobile_No, Exam_Type, Max_Marks, Class_Teacher, Total_Marks, Percentage, Class_Rank, Subject_Data, Attendance, Discipline, Remarks, School_Name)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            str(row.get('Class', '')), format_clean_number(row.get('Roll_No', '')), str(row.get('Student_Name', '')), str(row.get('Father_Name', '')), str(row.get('Mother_Name', '')),
            str(row.get('DOB', '')), format_clean_number(row.get('Aadhaar_No', '')), format_clean_number(row.get('Mobile_No', '')), str(row.get('Exam_Type', '')), 
            float(row.get('Max_Marks', 600)), str(row.get('Class_Teacher', '')), float(row.get('Total_Marks', 0)), 
            float(row.get('Percentage', 0)), int(row.get('Class_Rank', 0)), sub_json,
            str(row.get('Attendance', '95%')), str(row.get('Discipline', 'A')), str(row.get('Remarks', 'Good Performance')),
            str(row.get('School_Name', 'PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR'))
        ))
    conn.commit()
    conn.close()

def load_sqlite_to_df():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT Class, Roll_No, Student_Name, Father_Name, DOB, Aadhaar_No, Mobile_No, Exam_Type, Max_Marks, Class_Teacher, Total_Marks, Percentage, Class_Rank, Subject_Data, Attendance, Discipline, Remarks, Mother_Name, School_Name FROM student_results")
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
            'Remarks': r[16] if r[16] else 'Good Performance',
            'Mother_Name': r[17] if len(r) > 17 and r[17] else '',
            'School_Name': r[18] if len(r) > 18 and r[18] else 'PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR'
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

    meta_cols = ['Class', 'Roll_No', 'Student_Name', 'Father_Name', 'Mother_Name', 'DOB', 'Aadhaar_No', 'Mobile_No', 'Exam_Type', 'Max_Marks', 'Class_Teacher', 'Attendance', 'Discipline', 'Remarks', 'School_Name']
    for col in meta_cols:
        if col not in df.columns:
            if col == 'Attendance':
                df[col] = "95%"
            elif col == 'Discipline':
                df[col] = "A"
            elif col == 'Remarks':
                df[col] = "Good Performance"
            elif col == 'School_Name':
                df[col] = "PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR"
            else:
                df[col] = ""

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

# Custom Canvas Callback for Watermark & Decorative Outer Border
def draw_report_card_border_and_watermark(canvas_obj, doc, school_name="SCHOOL"):
    canvas_obj.saveState()
    w, h = doc.pagesize
    
    # Outer Maroon Decorative Double Border
    canvas_obj.setStrokeColor(colors.HexColor('#A91D22'))
    canvas_obj.setLineWidth(2.5)
    canvas_obj.rect(12, 12, w - 24, h - 24)
    
    canvas_obj.setLineWidth(1.0)
    canvas_obj.rect(16, 16, w - 32, h - 32)
    
    # Corner Accents
    for x, y in [(20, 20), (w - 20, 20), (20, h - 20), (w - 20, h - 20)]:
        canvas_obj.circle(x, y, 3, stroke=1, fill=1)
        
    # Background Watermark
    canvas_obj.setFont('Helvetica-Bold', 36)
    canvas_obj.setFillColor(colors.HexColor('#CCCCCC'), alpha=0.15)
    canvas_obj.saveState()
    canvas_obj.translate(w / 2.0, h / 2.0)
    canvas_obj.rotate(32)
    canvas_obj.drawCentredString(0, 0, str(school_name).upper())
    canvas_obj.restoreState()
    
    canvas_obj.restoreState()

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

    school_title = student_info.get('School_Name', 'PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR')
    story.append(Paragraph(f"<b>{school_title.upper()}</b>", cert_title))
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

    seal_w = st.session_state.get("seal_width", 50)
    seal_h = st.session_state.get("seal_height", 50)
    sign_w = st.session_state.get("sig_width", 70)
    sign_h = st.session_state.get("sig_height", 35)

    seal_element = RLImage(SEAL_PATH, width=seal_w, height=seal_h) if os.path.exists(SEAL_PATH) else Paragraph("<b>[OFFICIAL SEAL]</b>", body_style)
    sign_element = RLImage(SIGN_PATH, width=sign_w, height=sign_h) if os.path.exists(SIGN_PATH) else Paragraph("<b>____________________</b>", body_style)

    sig_data = [
        [Paragraph("<b>____________________</b>", body_style), seal_element, sign_element],
        [Paragraph("<b>Class Teacher</b>", body_style), Paragraph("<b>School Seal</b>", body_style), Paragraph("<b>Principal Signature</b>", body_style)]
    ]
    sig_table = Table(sig_data, colWidths=[240, 180, 240])
    sig_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(sig_table)

    watermark_fn = lambda c, d: draw_report_card_border_and_watermark(c, d, school_title)
    doc.build(story, onFirstPage=watermark_fn, onLaterPages=watermark_fn)
    buffer.seek(0)
    return buffer.getvalue()

def generate_pdf_scorecard(student_info, filtered_df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=22, leftMargin=22, topMargin=22, bottomMargin=22)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=14, leading=16, alignment=1, textColor=colors.HexColor('#A91D22'))
    sub_title_style = ParagraphStyle('DocSub', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=1, textColor=colors.HexColor('#333333'))
    term_style = ParagraphStyle('TermTitle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1, textColor=colors.HexColor('#0D47A1'))
    
    cell_head = ParagraphStyle('CellHead', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, leading=9, alignment=1)
    cell_left = ParagraphStyle('CellLeft', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10, alignment=0)
    cell_text = ParagraphStyle('CellText', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=1)
    bold_center = ParagraphStyle('BoldCenter', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=10, alignment=1)

    # 1. Header (School Logo + Title + CBSE Logo)
    logo_left = RLImage(LOGO_PATH, width=50, height=50) if os.path.exists(LOGO_PATH) else Paragraph("<b>[SCHOOL LOGO]</b>", bold_center)
    logo_right = RLImage(CBSE_LOGO_PATH, width=50, height=50) if os.path.exists(CBSE_LOGO_PATH) else (RLImage(LOGO_PATH, width=50, height=50) if os.path.exists(LOGO_PATH) else Paragraph("<b>[CBSE LOGO]</b>", bold_center))
    
    school_name_str = str(student_info.get('School_Name', 'PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR')).strip()
    if not school_name_str or school_name_str == 'nan':
        school_name_str = "PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR"

    header_data = [
        [
            logo_left,
            [
                Paragraph(f"<b>{school_name_str.upper()}</b>", title_style),
                Paragraph("CBSE Affiliated Senior Secondary Residential School<br/>Helpline: +91 2669 222020 | Email: jnvchhotaudepur@gmail.com", sub_title_style),
                Spacer(1, 2),
                Paragraph("<b>ACADEMIC PERFORMANCE REPORT CARD (SESSION 2025-2026)</b>", term_style)
            ],
            logo_right
        ]
    ]
    t_header = Table(header_data, colWidths=[60, 430, 60])
    t_header.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(t_header)
    story.append(Spacer(1, 4))

    # 2. Student Profile Block & Photo
    photo_file = f"photos/students/{student_info['Roll_No']}.png"
    photo_elem = RLImage(photo_file, width=70, height=80) if os.path.exists(photo_file) else Paragraph("<b>[STUDENT<br/>PHOTO]</b>", bold_center)

    mother_name = student_info.get('Mother_Name', '')
    if not mother_name or pd.isna(mother_name):
        mother_name = "N/A"

    info_data = [
        [Paragraph(f"<b>Student's Name :</b> {student_info['Student_Name']}", cell_left), Paragraph(f"<b>Class & Sec :</b> Class {student_info['Class']}", cell_left)],
        [Paragraph(f"<b>Father's Name :</b> {student_info['Father_Name']}", cell_left), Paragraph(f"<b>Roll No :</b> {student_info['Roll_No']}", cell_left)],
        [Paragraph(f"<b>Mother's Name :</b> {mother_name}", cell_left), Paragraph(f"<b>Aadhaar No :</b> [Aadhaar Redacted]", cell_left)],
        [Paragraph(f"<b>D.O.B. :</b> {student_info['DOB']}", cell_left), Paragraph(f"<b>Attendance :</b> {student_info.get('Attendance', '95%')}", cell_left)],
        [Paragraph(f"<b>Class Teacher :</b> {student_info.get('Class_Teacher', 'Teacher Incharge')}", cell_left), Paragraph(f"<b>Discipline Grade :</b> {student_info.get('Discipline', 'A')}", cell_left)]
    ]
    t_info = Table(info_data, colWidths=[250, 200])
    t_info.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2),
    ]))

    profile_block = Table([
        [t_info, photo_elem]
    ], colWidths=[450, 100])
    profile_block.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#A91D22')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E0E0E0')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (1,0), (1,0), 'CENTER'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(profile_block)
    story.append(Spacer(1, 6))

    # 3. Comprehensive Multi-Term Scholastic Marks Table
    sch_rows = [
        [
            Paragraph("<b>Scholastic Subjects</b>", cell_head),
            Paragraph("<b>Term 1 Performance</b>", cell_head), "", "", "", "",
            Paragraph("<b>Term 2 Performance</b>", cell_head), "", "", "", "",
            Paragraph("<b>Overall Performance</b>", cell_head), ""
        ],
        [
            "",
            Paragraph("<b>PWT-1<br/>(10)</b>", cell_head), Paragraph("<b>PWT-2<br/>(10)</b>", cell_head), Paragraph("<b>Term-1<br/>(80)</b>", cell_head), Paragraph("<b>Total<br/>(100)</b>", cell_head), Paragraph("<b>Grade</b>", cell_head),
            Paragraph("<b>PWT-3<br/>(10)</b>", cell_head), Paragraph("<b>PWT-4<br/>(10)</b>", cell_head), Paragraph("<b>Term End<br/>(80)</b>", cell_head), Paragraph("<b>Total<br/>(100)</b>", cell_head), Paragraph("<b>Grade</b>", cell_head),
            Paragraph("<b>Total Marks</b>", cell_head), Paragraph("<b>Grade</b>", cell_head)
        ]
    ]

    active_subs = []
    for sub in ALL_SUBJECTS:
        if sub in filtered_df.columns and filtered_df[sub].notna().any():
            active_subs.append(sub)
    if not active_subs:
        active_subs = ['Gujarati', 'Hindi', 'English', 'Mathematics', 'Science', 'Social_Science']

    def get_exam_mark(exam_kw, sub_col):
        for _, r in filtered_df.iterrows():
            if exam_kw.lower() in str(r['Exam_Type']).lower():
                v = r.get(sub_col, np.nan)
                if pd.notna(v):
                    return float(v)
        return None

    tot_overall_obtained = 0
    tot_overall_max = 0

    for sub in active_subs:
        pwt1 = get_exam_mark('pwt-1', sub) or get_exam_mark('pwt 1', sub)
        pwt2 = get_exam_mark('pwt-2', sub) or get_exam_mark('pwt 2', sub)
        term1 = get_exam_mark('term-1', sub) or get_exam_mark('term 1', sub) or get_exam_mark('half', sub)
        
        t1_parts = [v for v in [pwt1, pwt2, term1] if v is not None]
        t1_total = sum(t1_parts) if t1_parts else None
        t1_grade = calculate_grade(t1_total) if t1_total is not None else "-"

        pwt3 = get_exam_mark('pwt-3', sub) or get_exam_mark('pwt 3', sub)
        pwt4 = get_exam_mark('pwt-4', sub) or get_exam_mark('pwt 4', sub)
        term2 = get_exam_mark('term end', sub) or get_exam_mark('annual', sub) or get_exam_mark('final', sub)
        
        t2_parts = [v for v in [pwt3, pwt4, term2] if v is not None]
        t2_total = sum(t2_parts) if t2_parts else None
        t2_grade = calculate_grade(t2_total) if t2_total is not None else "-"

        comb_parts = []
        if t1_total is not None: comb_parts.append(t1_total)
        if t2_total is not None: comb_parts.append(t2_total)

        if comb_parts:
            overall_sub_score = sum(comb_parts) / len(comb_parts)
            tot_overall_obtained += overall_sub_score
            tot_overall_max += 100
            overall_grade = calculate_grade(overall_sub_score)
        else:
            overall_sub_score = None
            overall_grade = "-"

        def fmt(val):
            if val is None: return "-"
            return f"{val:.1f}" if not float(val).is_integer() else str(int(val))

        sch_rows.append([
            Paragraph(f"<b>{sub.replace('_', ' ')}</b>", cell_left),
            Paragraph(fmt(pwt1), cell_text), Paragraph(fmt(pwt2), cell_text), Paragraph(fmt(term1), cell_text), Paragraph(fmt(t1_total), cell_text), Paragraph(t1_grade, cell_text),
            Paragraph(fmt(pwt3), cell_text), Paragraph(fmt(pwt4), cell_text), Paragraph(fmt(term2), cell_text), Paragraph(fmt(t2_total), cell_text), Paragraph(t2_grade, cell_text),
            Paragraph(fmt(overall_sub_score), cell_text), Paragraph(overall_grade, cell_text)
        ])

    t_sch = Table(sch_rows, colWidths=[92, 38, 38, 44, 42, 34, 38, 38, 44, 42, 34, 42, 24])
    t_sch.setStyle(TableStyle([
        ('SPAN', (0,0), (0,1)),
        ('SPAN', (1,0), (5,0)),
        ('SPAN', (6,0), (10,0)),
        ('SPAN', (11,0), (12,0)),
        ('BACKGROUND', (0,0), (-1,1), colors.HexColor('#F8E8E8')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#555555')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('PADDING', (0,0), (-1,-1), 2.5),
    ]))
    story.append(t_sch)
    story.append(Spacer(1, 6))

    # 4. Overall Summary Box
    pct_val = (tot_overall_obtained / tot_overall_max * 100) if tot_overall_max > 0 else filtered_df['Percentage'].mean()
    ov_grade = calculate_grade(pct_val)

    summary_data = [
        [
            Paragraph(f"<b>TOTAL MARKS:</b> {int(tot_overall_obtained)} / {int(tot_overall_max if tot_overall_max > 0 else 600)}", cell_left),
            Paragraph(f"<b>PERCENTAGE:</b> {pct_val:.2f}%", cell_head),
            Paragraph(f"<b>OVERALL GRADE:</b> {ov_grade}", cell_head)
        ]
    ]
    t_sum = Table(summary_data, colWidths=[240, 160, 150])
    t_sum.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#E8F5E9')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#2E7D32')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#A5D6A7')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_sum)
    story.append(Spacer(1, 6))

    # 5. Co-Scholastic & Grade Scale Legend
    coscholastic_data = [
        [Paragraph("<b>Co-Scholastic Areas (5 Point Scale)</b>", cell_head), Paragraph("<b>Grade</b>", cell_head)],
        [Paragraph("Work Education (Pre-Vocational)", cell_left), Paragraph("A", cell_text)],
        [Paragraph("Art Education", cell_left), Paragraph("A", cell_text)],
        [Paragraph("Health & Physical Education", cell_left), Paragraph("A", cell_text)],
        [Paragraph("Discipline & Conduct", cell_left), Paragraph("A", cell_text)]
    ]
    t_cos = Table(coscholastic_data, colWidths=[195, 50])
    t_cos.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#EEEEEE')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#777777')),
        ('PADDING', (0,0), (-1,-1), 2.5),
    ]))

    scale_data = [
        [Paragraph("<b>Grading Scale for Scholastic Areas</b>", cell_head), ""],
        [Paragraph("91 - 100 : A1 &nbsp;&nbsp;|&nbsp;&nbsp; 81 - 90 : A2", cell_text), Paragraph("71 - 80 : B1 &nbsp;&nbsp;|&nbsp;&nbsp; 61 - 70 : B2", cell_text)],
        [Paragraph("51 - 60 : C1 &nbsp;&nbsp;|&nbsp;&nbsp; 41 - 50 : C2", cell_text), Paragraph("33 - 40 : D &nbsp;&nbsp;|&nbsp;&nbsp; Below 33 : E (Needs Attention)", cell_text)]
    ]
    t_scale = Table(scale_data, colWidths=[150, 150])
    t_scale.setStyle(TableStyle([
        ('SPAN', (0,0), (1,0)),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#EEEEEE')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#777777')),
        ('PADDING', (0,0), (-1,-1), 2.5),
    ]))

    co_scale_table = Table([[t_cos, t_scale]], colWidths=[250, 300])
    co_scale_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(co_scale_table)
    story.append(Spacer(1, 6))

    # 6. Teacher Remarks & Promotion
    remarks_data = [
        [Paragraph(f"<b>Class Teacher's Remarks:</b> {student_info.get('Remarks', 'Good Performance and active participation.')}", cell_left)],
        [Paragraph("<b>Result Status:</b> PASSED AND PROMOTED TO NEXT HIGHER CLASS", ParagraphStyle('Pass', parent=cell_left, textColor=colors.HexColor('#1B5E20')))]
    ]
    t_rem = Table(remarks_data, colWidths=[550])
    t_rem.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#888888')),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#FAFAFA')),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(t_rem)
    story.append(Spacer(1, 10 + st.session_state.get("sig_y_offset", 0)))

    # 7. Signatures Table (Parent, Teacher, Principal & Seal)
    seal_w = st.session_state.get("seal_width", 50)
    seal_h = st.session_state.get("seal_height", 50)
    sign_w = st.session_state.get("sig_width", 70)
    sign_h = st.session_state.get("sig_height", 35)

    seal_elem = RLImage(SEAL_PATH, width=seal_w, height=seal_h) if os.path.exists(SEAL_PATH) else Paragraph("<b>[SEAL]</b>", bold_center)
    sign_elem = RLImage(SIGN_PATH, width=sign_w, height=sign_h) if os.path.exists(SIGN_PATH) else Paragraph("<b>________________</b>", bold_center)

    principal_block = Table([[seal_elem, sign_elem]], colWidths=[seal_w + 5, sign_w + 5])
    principal_block.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER')
    ]))

    sig_data = [
        [
            Paragraph("<b>______________________</b>", cell_head),
            Paragraph("<b>______________________</b>", cell_head),
            principal_block
        ],
        [
            Paragraph("<b>Parent's Signature</b>", cell_head),
            Paragraph("<b>Class Incharge Signature</b>", cell_head),
            Paragraph("<b>Principal Signature & School Seal</b>", cell_head)
        ]
    ]
    t_sig = Table(sig_data, colWidths=[180, 185, 185])
    t_sig.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(t_sig)

    watermark_fn = lambda c, d: draw_report_card_border_and_watermark(c, d, school_name_str)
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

    school_name_str = student_info.get('School_Name', 'PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR')
    story.append(Paragraph(f"<b>{school_name_str.upper()}</b>", title_style))
    story.append(Paragraph("EXAMINATION ADMIT CARD / HALL TICKET (2025-26)", sub_style))
    story.append(Spacer(1, 10))

    photo_path = f"photos/students/{student_info['Roll_No']}.png"
    photo_elem = RLImage(photo_path, width=70, height=80) if os.path.exists(photo_path) else Paragraph("<b>[PHOTO]</b>", normal_style)

    info_data = [
        [Paragraph(f"<b>Student Name:</b> {student_info['Student_Name']}", normal_style), photo_elem],
        [Paragraph(f"<b>Roll No:</b> {student_info['Roll_No']}", normal_style), ""],
        [Paragraph(f"<b>Class:</b> {student_info['Class']}", normal_style), ""],
        [Paragraph(f"<b>Father's Name:</b> {student_info['Father_Name']}", normal_style), ""],
        [Paragraph(f"<b>Exam Center:</b> Main School Campus", normal_style), ""]
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

    seal_w = st.session_state.get("seal_width", 45)
    seal_h = st.session_state.get("seal_height", 45)
    sign_w = st.session_state.get("sig_width", 65)
    sign_h = st.session_state.get("sig_height", 30)

    seal_element = RLImage(SEAL_PATH, width=seal_w, height=seal_h) if os.path.exists(SEAL_PATH) else Paragraph("<b>[SEAL]</b>", normal_style)
    sign_element = RLImage(SIGN_PATH, width=sign_w, height=sign_h) if os.path.exists(SIGN_PATH) else Paragraph("<b>________________</b>", normal_style)

    sig_data = [
        [Paragraph("<b>____________________</b>", normal_style), seal_element, sign_element],
        [Paragraph("<b>Student Signature</b>", normal_style), Paragraph("<b>School Seal</b>", normal_style), Paragraph("<b>Principal Signature</b>", normal_style)]
    ]
    sig_table = Table(sig_data, colWidths=[180, 160, 180])
    sig_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(sig_table)

    watermark_fn = lambda c, d: draw_report_card_border_and_watermark(c, d, school_name_str)
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
    if os.path.exists(CBSE_LOGO_PATH):
        st.markdown(f'<div class="header-logo-right mobile-hide"><img src="data:image/png;base64,{get_base64_image(CBSE_LOGO_PATH)}" width="80"></div>', unsafe_allow_html=True)
    elif os.path.exists(LOGO_PATH):
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
                    st.download_button("📥 Comprehensive Report Card (PDF)", data=pdf_bytes, file_name=f"Report_{student_info['Roll_No']}.pdf", mime="application/pdf", use_container_width=True)
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

        with st.expander("✒️ REPORT CARD STAMP & SIGNATURE ALIGNMENT SETTINGS", expanded=True):
            st.write("🔧 Adjust size and offsets of Principal Signature and School Stamp on PDF Report Cards:")
            sc1, sc2, sc3, sc4 = st.columns(4)
            with sc1:
                st.session_state["sig_width"] = st.slider("Signature Width (px)", 40, 150, st.session_state.get("sig_width", 80))
                st.session_state["sig_height"] = st.slider("Signature Height (px)", 20, 80, st.session_state.get("sig_height", 35))
            with sc2:
                st.session_state["seal_width"] = st.slider("Stamp/Seal Width (px)", 30, 120, st.session_state.get("seal_width", 55))
                st.session_state["seal_height"] = st.slider("Stamp/Seal Height (px)", 30, 120, st.session_state.get("seal_height", 55))
            with sc3:
                st.session_state["sig_y_offset"] = st.slider("Vertical Shift Above Signatures (px)", -20, 40, st.session_state.get("sig_y_offset", 0))
            with sc4:
                st.write(" ")
                st.write(" ")
                if st.button("💾 Save Alignment Settings"):
                    st.success("✅ Signature and Stamp positions updated!")

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
                                "Pass % (>=33)": f"{round((s_data[s_data >= 33].count() / len(s_data))*100, 2)}%"
                            })
                    st.dataframe(pd.DataFrame(sub_stats), hide_index=True, use_container_width=True)