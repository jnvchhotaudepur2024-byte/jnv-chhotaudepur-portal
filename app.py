import streamlit as st
import pandas as pd
import numpy as np
import os
import shutil
import re
import json
import urllib.parse
import io
import base64
import sqlite3
from datetime import datetime
from PIL import Image
import plotly.graph_objects as go

# ReportLab Imports for PDF & Certificate Generation
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="PM SHRI JNV CHHOTAUDEPUR - RESULT PORTAL", page_icon="🎓", layout="wide")

# Custom CSS for UI Enhancement
st.markdown("""
    <style>
    @media (max-width: 768px) {
        .stApp { padding: 5px !important; }
        h1 { font-size: 1.4rem !important; }
        h3 { font-size: 1.1rem !important; }
        .stButton>button { width: 100% !important; }
    }
    .main-title {
        text-transform: uppercase;
        font-weight: 800;
        color: #1E88E5;
        margin: 0 !important;
        padding: 0 !important;
    }
    .sub-title {
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #333333;
        margin-top: 2px !important;
        padding: 0 !important;
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
    </style>
""", unsafe_allow_html=True)

# Directory Setup
os.makedirs("photos/students", exist_ok=True)
os.makedirs("photos/gallery", exist_ok=True)
os.makedirs("photos/board", exist_ok=True)
os.makedirs("photos/system", exist_ok=True)
os.makedirs("backups", exist_ok=True)

# Security: Fetch Admin Password from Secrets or Env Variable
try:
    ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", os.getenv("ADMIN_PASSWORD", "Jnvcu@me2"))
except Exception:
    ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "Jnvcu@me2")
DB_FILE = "school_database.db"
EXCEL_FILE_PATH = "JNV_Student_Marks.xlsx"

ALL_SUBJECTS = [
    'Gujarati', 'Hindi', 'English', 'Mathematics', 
    'Science', 'Social_Science', 'Physics', 'Chemistry', 'Biology'
]

# 🌐 MODIFICATION 3: Multi-Language Dictionary Support
LANG_TEXTS = {
    "English": {
        "title": "STUDENT PERFORMANCE & RESULT PORTAL",
        "search_lbl": "🔎 CHECK STUDENT RESULT",
        "cert_btn": "🏆 Download Merit Certificate",
        "weak_alert": "⚠️ Needs Special Attention in the following subjects (< 33%):",
        "overall_status": "Overall Performance Status",
        "pass": "PASS / EXCELLENT",
        "needs_imp": "NEEDS IMPROVEMENT",
        "chart_title": "Subject-wise Performance Breakdown"
    },
    "Hindi": {
        "title": "विद्यार्थी प्रदर्शन एवं परिणाम पोर्टल",
        "search_lbl": "🔎 छात्र परिणाम खोजें",
        "cert_btn": "🏆 योग्यता प्रमाण पत्र (Merit Certificate) डाउनलोड करें",
        "weak_alert": "⚠️ निम्नलिखित विषयों में विशेष ध्यान देने की आवश्यकता है (< 33%):",
        "overall_status": "कुल प्रदर्शन स्थिति",
        "pass": "उत्तीर्ण / उत्कृष्ट",
        "needs_imp": "सुधार की आवश्यकता है",
        "chart_title": "विषय-वार अंक विश्लेषण graph"
    },
    "Gujarati": {
        "title": "વિદ્યાર્થી પ્રદર્શન અને પરિણામ પોર્ટલ",
        "search_lbl": "🔎 વિદ્યાર્થીનું પરિણામ જુઓ",
        "cert_btn": "🏆 મેરિટ સર્ટિફિકેટ ડાઉનલોડ કરો",
        "weak_alert": "⚠️ નીચેના વિષયોમાં વિશેષ ધ્યાન આપવાની જરૂર છે (< 33%):",
        "overall_status": "સમગ્ર પ્રદર્શન સ્થિતિ",
        "pass": "ઉત્તીર્ણ / ઉત્કૃષ્ટ",
        "needs_imp": "સુધારાની જરૂર છે",
        "chart_title": "વિષયવાર ગુણ વિશ્લેષણ ગ્રાફ"
    }
}

# 🔒 MODIFICATION 1: Aadhaar Masking Function
def mask_aadhaar(val):
    clean_a = re.sub(r'[^0-9]', '', str(val))
    if len(clean_a) >= 4:
        return f"XXXX-XXXX-{clean_a[-4:]}"
    return "XXXX-XXXX-XXXX"

def clean_val(val):
    if pd.isna(val) or val is None:
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(val)).lower().strip()

def get_rank_comment(rank):
    if rank == 1:
        return "🥇 Outstanding Performance! Class Rank #1. Keep it up! 🌟"
    elif rank <= 3:
        return "🥈 Excellent Performance! Top 3 in Class. Keep pushing! 💪"
    elif rank <= 10:
        return "🌟 Very Good Effort! Top 10 Performer. Work hard for Top 3!"
    elif rank <= 20:
        return "👍 Good Effort! Consistent practice will boost your rank."
    else:
        return "💪 Scope for Improvement! Focus on weaker subjects."

@st.cache_data
def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode()
    return None

BG_PATH = "photos/system/background.png"
LOGO_PATH = "photos/system/logo.png"

encoded_string = get_base64_image(BG_PATH)
if encoded_string:
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: linear-gradient(rgba(255, 255, 255, 0.90), rgba(255, 255, 255, 0.90)), url("data:image/png;base64,{encoded_string}");
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
        except:
            count = 0
        count += 1
        f.seek(0)
        f.write(str(count))
        f.truncate()
    return count

total_visits = get_and_increment_visits()

# 🗄️ MODIFICATION 6: SQLite Database Integration
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS student_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            Class TEXT, Roll_No TEXT, Student_Name TEXT, Father_Name TEXT,
            DOB TEXT, Aadhaar_No TEXT, Mobile_No TEXT, Exam_Type TEXT,
            Max_Marks REAL, Class_Teacher TEXT, Total_Marks REAL,
            Percentage REAL, Class_Rank INTEGER, Subject_Data TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def sync_df_to_sqlite(df):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM student_results")
    for _, row in df.iterrows():
        sub_json = json.dumps({s: row[s] for s in ALL_SUBJECTS if s in row and pd.notna(row[s])})
        cursor.execute('''
            INSERT INTO student_results (Class, Roll_No, Student_Name, Father_Name, DOB, Aadhaar_No, Mobile_No, Exam_Type, Max_Marks, Class_Teacher, Total_Marks, Percentage, Class_Rank, Subject_Data)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (str(row['Class']), str(row['Roll_No']), str(row['Student_Name']), str(row['Father_Name']), str(row['DOB']), str(row['Aadhaar_No']), str(row['Mobile_No']), str(row['Exam_Type']), float(row['Max_Marks']), str(row['Class_Teacher']), float(row['Total_Marks']), float(row['Percentage']), int(row['Class_Rank']), sub_json))
    conn.commit()
    conn.close()

def log_parent_search(roll_no, student_name, selected_class):
    log_file = "result_logs.csv"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_data = pd.DataFrame([{"Timestamp": timestamp, "Roll_No": str(roll_no), "Student_Name": student_name, "Class": selected_class}])
    if os.path.exists(log_file):
        new_data.to_csv(log_file, mode='a', header=False, index=False)
    else:
        new_data.to_csv(log_file, mode='w', header=True, index=False)

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

    meta_cols = ['Class', 'Roll_No', 'Student_Name', 'Father_Name', 'DOB', 'Aadhaar_No', 'Mobile_No', 'Exam_Type', 'Max_Marks', 'Class_Teacher']
    for col in meta_cols:
        if col not in df.columns:
            df[col] = ""

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
    else:
        st.session_state["student_data"] = None

if "admin_logged_in" not in st.session_state:
    st.session_state["admin_logged_in"] = False

BOARD_TOPPERS_FILE = "board_toppers.json"
def load_board_toppers():
    if os.path.exists(BOARD_TOPPERS_FILE):
        with open(BOARD_TOPPERS_FILE, "r") as f:
            return json.load(f)
    return []

# 🏆 MODIFICATION 5: Auto Merit Certificate PDF Generator
def generate_merit_certificate_pdf(student_info, exam_type, percentage, rank):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []
    styles = getSampleStyleSheet()

    cert_title = ParagraphStyle('CertTitle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=26, leading=30, alignment=1, textColor=colors.HexColor('#1565C0'))
    sub_title = ParagraphStyle('SubTitle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=18, leading=22, alignment=1, textColor=colors.HexColor('#2E7D32'))
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontName='Helvetica', fontSize=14, leading=22, alignment=1)

    story.append(Spacer(1, 20))
    story.append(Paragraph("PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR", cert_title))
    story.append(Spacer(1, 10))
    story.append(Paragraph("🏆 CERTIFICATE OF ACADEMIC EXCELLENCE 🏆", sub_title))
    story.append(Spacer(1, 25))

    cert_text = f"""
    This is to proudly certify that <b>{student_info['Student_Name']}</b>, Son/Daughter of <b>{student_info['Father_Name']}</b>, 
    studying in <b>Class {student_info['Class']}</b> (Roll No: <b>{student_info['Roll_No']}</b>), has secured 
    <font color="#1565C0"><b>RANK #{rank}</b></font> with an outstanding score of <b>{percentage:.2f}%</b> 
    in the <b>{exam_type}</b> Examination (Academic Year 2025-26).
    """
    story.append(Paragraph(cert_text, body_style))
    story.append(Spacer(1, 40))

    sig_data = [
        [Paragraph("<b>____________________</b>", body_style), Paragraph("<b>____________________</b>", body_style)],
        [Paragraph("<b>Class Teacher</b>", body_style), Paragraph("<b>Principal Signature</b>", body_style)]
    ]
    sig_table = Table(sig_data, colWidths=[350, 350])
    sig_table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    story.append(sig_table)

    doc.build(story)
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

    story.append(Paragraph("PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR", title_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph("STUDENT ACADEMIC PERFORMANCE REPORT CARD", subtitle_style))
    story.append(Spacer(1, 15))

    masked_a = mask_aadhaar(student_info['Aadhaar_No'])
    info_data = [
        [Paragraph(f"<b>Student Name:</b> {student_info['Student_Name']}", normal_style), Paragraph(f"<b>Roll No:</b> {student_info['Roll_No']}", normal_style)],
        [Paragraph(f"<b>Class:</b> {student_info['Class']}", normal_style), Paragraph(f"<b>Class Teacher:</b> {student_info['Class_Teacher']}", normal_style)],
        [Paragraph(f"<b>Father's Name:</b> {student_info['Father_Name']}", normal_style), Paragraph(f"<b>Aadhaar:</b> {masked_a}", normal_style)]
    ]
    info_table = Table(info_data, colWidths=[260, 260])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8F9FA')),
        ('PADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E0E0E0')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 15))

    for _, row in filtered_df.iterrows():
        exam_header = f"<b>Exam:</b> {row['Exam_Type']} &nbsp;|&nbsp; <b>Score:</b> {int(row['Total_Marks'])}/{int(row['Max_Marks'])} ({row['Percentage']:.2f}%) &nbsp;|&nbsp; <b>Rank:</b> #{row['Class_Rank']}"
        story.append(Paragraph(exam_header, styles['Heading3']))
        story.append(Spacer(1, 5))

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
            ('PADDING', (0, 0), (-1, -1), 5),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('ALIGN', (2, 0), (2, -1), 'CENTER'),
        ]))
        story.append(score_table)
        story.append(Spacer(1, 12))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

@st.dialog("⚠️ CONFIRMATION / क्या आप आश्वस्त हैं?")
def confirm_action_dialog(title, callback, *args, **kwargs):
    st.write(f"**{title}**")
    st.write("Do you really want to perform this update?")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ Yes, Proceed", use_container_width=True):
            callback(*args, **kwargs)
            st.rerun()
    with c2:
        if st.button("❌ No, Cancel", use_container_width=True):
            st.rerun()

# 🌐 Sidebar Navigation & Language Switcher
st.sidebar.title("☰ NAVIGATION")
selected_lang = st.sidebar.selectbox("🌐 Choose Language / भाषा चुनें:", ["English", "Hindi", "Gujarati"])
txt = LANG_TEXTS[selected_lang]

menu = st.sidebar.radio("SELECT PORTAL / PAGE:", ["👨‍🎓 PARENT PORTAL", "🖼️ SCHOOL GALLERY", "🏆 BOARD EXAM RESULTS", "⚙️ ADMIN PORTAL"])

head_col1, head_col2 = st.columns([1, 5], vertical_alignment="center")
with head_col1:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=95)
with head_col2:
    st.markdown("<h1 class='main-title'>🏫 PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR</h1>", unsafe_allow_html=True)
    st.markdown(f"<h3 class='sub-title'>{txt['title']}</h3>", unsafe_allow_html=True)

st.markdown("---")

# ==============================================================================
# 👨‍🎓 PARENT PORTAL
# ==============================================================================
if menu == "👨‍🎓 PARENT PORTAL":
    educational_quotes = [
        "🎓 'Education is the most powerful weapon which you can use to change the world.' – Nelson Mandela",
        "🌟 'Live as if you were to die tomorrow. Learn as if you were to live forever.' – Mahatma Gandhi",
        "💡 'The mind is not a vessel to be filled, but a fire to be kindled.' – Plutarch"
    ]
    quotes_ticker_text = " &nbsp;&nbsp;&nbsp;&nbsp; ✦ &nbsp;&nbsp;&nbsp;&nbsp; ".join(educational_quotes)

    st.markdown(
        f"""
        <div style="background: linear-gradient(90deg, #1565C0, #1E88E5); border-radius: 6px; padding: 8px 12px; color: #FFFFFF; font-size: 14px; font-weight: 600; margin-bottom: 15px;">
            <marquee direction="left" scrollamount="6" behavior="scroll">{quotes_ticker_text}</marquee>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    if st.session_state["student_data"] is not None:
        df_data = st.session_state["student_data"]
        if not df_data.empty and 'Exam_Type' in df_data.columns:
            latest_exam = df_data['Exam_Type'].dropna().iloc[-1]
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
        search_method = st.radio("Choose Search Method:", ["Option 1: Roll No & Date of Birth (DOB)", "Option 2: Roll No & Aadhaar Number"], horizontal=True)
        
        with st.form("search_form"):
            c1, c2 = st.columns(2)
            with c1:
                selected_class = st.selectbox("Select Class", sorted(df['Class'].astype(str).unique()))
                roll_no = st.text_input("Roll No")
            
            if "Option 1" in search_method:
                with c2:
                    dob_input = st.text_input("Date of Birth")
            else:
                with c2:
                    aadhaar_input = st.text_input("Aadhaar Number")
            
            submit_btn = st.form_submit_button("🔍 View Result")

        if submit_btn:
            if "Option 1" in search_method:
                filtered_df = df[
                    (df['Class'].astype(str).str.strip().str.lower() == selected_class.strip().lower()) &
                    (df['Roll_No'].astype(str).str.strip() == roll_no.strip()) &
                    (df['DOB'].apply(clean_val) == clean_val(dob_input))
                ]
            else:
                filtered_df = df[
                    (df['Class'].astype(str).str.strip().str.lower() == selected_class.strip().lower()) &
                    (df['Roll_No'].astype(str).str.strip() == roll_no.strip()) &
                    (df['Aadhaar_No'].apply(clean_val) == clean_val(aadhaar_input))
                ]
            
            if filtered_df.empty:
                st.error("❌ Invalid Details! Kripya Roll No, DOB ya Aadhaar Sahi enter karein.")
            else:
                student_info = filtered_df.iloc[0]
                log_parent_search(student_info['Roll_No'], student_info['Student_Name'], student_info['Class'])
                
                pdf_bytes = generate_pdf_scorecard(student_info, filtered_df)

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
                    
                    b_c1, b_c2 = st.columns(2)
                    with b_c1:
                        st.download_button("📥 Download Report Card (PDF)", data=pdf_bytes, file_name=f"Report_{student_info['Roll_No']}.pdf", mime="application/pdf", use_container_width=True)
                    
                    # 🏆 Merit Certificate Download for Top 3 Rankers
                    best_row = filtered_df.sort_values('Class_Rank').iloc[0]
                    if best_row['Class_Rank'] <= 3:
                        with b_c2:
                            cert_pdf = generate_merit_certificate_pdf(student_info, best_row['Exam_Type'], best_row['Percentage'], best_row['Class_Rank'])
                            st.download_button(txt['cert_btn'], data=cert_pdf, file_name=f"Merit_Certificate_{student_info['Roll_No']}.pdf", mime="application/pdf", use_container_width=True)

                st.markdown("---")
                
                # 📊 MODIFICATION 2: Visual Plotly Analytics & Weak Subject Alert
                st.subheader(f"📊 {txt['chart_title']}")
                latest_row = filtered_df.iloc[-1]
                sub_names = [s for s in ALL_SUBJECTS if s in latest_row and pd.notna(latest_row[s])]
                sub_marks = [latest_row[s] for s in sub_names]
                
                # Plotly Chart
                fig = go.Figure(data=[go.Bar(x=sub_names, y=sub_marks, marker_color='#1E88E5', text=sub_marks, textposition='auto')])
                fig.update_layout(title=f"Marks Distribution ({latest_row['Exam_Type']})", xaxis_title="Subjects", yaxis_title="Marks", yaxis=dict(range=[0, 100]))
                st.plotly_chart(fig, use_container_width=True)

                # Weak Subject Red Alert Badge
                weak_subs = [f"{s} ({latest_row[s]} marks)" for s in sub_names if float(latest_row[s]) < 33.0]
                if weak_subs:
                    st.markdown(f"<div class='weak-badge'>{txt['weak_alert']} {', '.join(weak_subs)}</div>", unsafe_allow_html=True)

                st.markdown("---")
                st.subheader("📝 EXAM SCORECARD")
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
    st.header("🎓 CBSE BOARD TOPPERS")
    toppers_data = load_board_toppers()
    if not toppers_data:
        st.info("No toppers uploaded yet.")
    else:
        cards_html = ""
        for t in toppers_data:
            img_b64 = get_base64_image(t.get("photo", ""))
            img_src = f"data:image/png;base64,{img_b64}" if img_b64 else ""
            cards_html += f"""
            <div style="display: inline-block; width: 200px; background: #fff; padding: 12px; margin-right: 15px; border-radius: 10px; border: 2px solid #1E88E5; text-align: center;">
                <img src="{img_src}" style="width: 80px; height: 80px; border-radius: 50%; object-fit: cover; border: 2px solid #1565C0;">
                <div style="font-weight: bold; color: #0D47A1; margin-top: 5px;">{t['name']}</div>
                <div style="font-size: 12px;">{t['class']} ({t['year']})</div>
                <div style="font-size: 15px; font-weight: bold; color: #2E7D32; background: #E8F5E9; margin-top: 4px;">🏆 {t['percentage']}</div>
            </div>
            """
        st.markdown(f'<marquee direction="left" scrollamount="6" onmouseover="this.stop();" onmouseout="this.start();">{cards_html}</marquee>', unsafe_allow_html=True)

# ==============================================================================
# ⚙️ ADMIN PORTAL (🔒 Security & 📲 Bulk WhatsApp Alerts)
# ==============================================================================
elif menu == "⚙️ ADMIN PORTAL":
    st.header("🔒 ADMIN DASHBOARD")
    st.info(f"👁️ **Total Website Visits:** `{total_visits}`")
    
    if not st.session_state["admin_logged_in"]:
        with st.form("login_form"):
            st.subheader("🔐 Admin Login")
            admin_user = st.text_input("Username")
            admin_pass = st.text_input("Password", type="password")
            if st.form_submit_button("Login"):
                if admin_user == "admin" and admin_pass == ADMIN_PASSWORD:
                    st.session_state["admin_logged_in"] = True
                    st.success("✅ Logged In!")
                    st.rerun()
                else:
                    st.error("❌ Invalid Credentials!")
    else:
        st.success("🔓 Admin Logged In")
        if st.button("Logout"):
            st.session_state["admin_logged_in"] = False
            st.rerun()

        st.markdown("---")

        # 📲 MODIFICATION 4: Bulk WhatsApp Notification Links Generator
        with st.expander("📲 1. BULK WHATSAPP NOTIFICATIONS", expanded=False):
            if st.session_state["student_data"] is not None:
                df_notif = st.session_state["student_data"]
                if 'Mobile_No' in df_notif.columns:
                    n_cls = st.selectbox("Select Class for Bulk Dispatch", sorted(df_notif['Class'].astype(str).unique()))
                    filtered_notif = df_notif[df_notif['Class'].astype(str) == str(n_cls)]
                    msg_template = st.text_area("Message Content", "Dear Parent, your child's exam results are live on the portal. Check now!")
                    
                    if st.button("🚀 Generate Bulk WhatsApp Links"):
                        for _, row in filtered_notif.iterrows():
                            mob = re.sub(r'[^0-9]', '', str(row['Mobile_No']))
                            if len(mob) >= 10:
                                mob = "91" + mob[-10:]
                                encoded_msg = urllib.parse.quote(f"Hello {row['Student_Name']},\n{msg_template}")
                                wa_link = f"https://api.whatsapp.com/send?phone={mob}&text={encoded_msg}"
                                st.markdown(f"👉 **{row['Student_Name']}** -> [Click to Send WhatsApp Alert]({wa_link})")
                else:
                    st.error("Mobile_No column missing in data.")

        # Media & Background Settings
        with st.expander("🎨 2. BRANDING & BACKGROUND", expanded=False):
            up_logo = st.file_uploader("Upload Logo", type=["png", "jpg"])
            if st.button("Save Logo") and up_logo:
                Image.open(up_logo).save(LOGO_PATH)
                st.success("Logo Updated!")

        # Board Toppers Management
        with st.expander("🏆 3. BOARD TOPPERS MANAGEMENT", expanded=False):
            b_class = st.selectbox("Class", ["Class 10", "Class 12"])
            b_name = st.text_input("Name")
            b_percent = st.text_input("Percentage (e.g. 98.4%)")
            b_year = st.text_input("Year", value="2025-26")
            b_photo = st.file_uploader("Photo", type=["jpg", "png"])
            if st.button("Add Board Topper") and b_name and b_photo:
                photo_file = f"photos/board/{b_class.replace(' ', '_')}_{clean_val(b_name)}.png"
                Image.open(b_photo).save(photo_file)
                toppers = load_board_toppers()
                toppers.append({"class": b_class, "name": b_name, "percentage": b_percent, "year": b_year, "photo": photo_file})
                with open(BOARD_TOPPERS_FILE, "w") as f:
                    json.dump(toppers, f)
                st.success("Topper Added!")

        # Excel Upload with SQLite Sync
        with st.expander("📤 4. EXCEL DATA UPLOAD & SQLITE SYNC", expanded=False):
            uploaded_file = st.file_uploader("Upload Excel Sheet (.xlsx)", type=["xlsx", "xls"])
            if st.button("Process & Sync Database") and uploaded_file:
                with open(EXCEL_FILE_PATH, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                st.session_state["student_data"] = process_data_excel(EXCEL_FILE_PATH)
                st.success("Excel & SQLite Database Successfully Updated!")

st.markdown("---")
st.markdown("<div style='text-align: center; color: #666;'>© 2026 PM SHRI JNV CHHOTAUDEPUR | All Rights Reserved</div>", unsafe_allow_html=True)