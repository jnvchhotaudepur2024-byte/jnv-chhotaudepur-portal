import base64
from datetime import datetime
import hashlib
import io
import json
import os
import random
import re
import shutil
import sqlite3
import urllib.parse
import zipfile

import numpy as np
import pandas as pd
from PIL import Image
import plotly.graph_objects as go
import requests

# ReportLab Imports for PDF, Watermark, Signatures, Borders & Detailed Report Card Generation
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Image as RLImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st

st.set_page_config(
    page_title="PM SHRI JNV CHHOTAUDEPUR - RESULT PORTAL",
    page_icon="🎓",
    layout="wide",
)

# Custom CSS & Responsive Mobile Optimization
st.markdown(
    """
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
    .header-logo-left { display: flex; justify-content: flex-start; align-items: center; }
    .header-logo-right { display: flex; justify-content: flex-end; align-items: center; }
    .main-title {
        text-transform: uppercase;
        font-weight: 800;
        color: #B22222;
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
        border: 2px solid #B22222;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        vertical-align: top;
    }
    .hall-of-fame-box {
        background: linear-gradient(135deg, #FFEBEE, #FFF9C4);
        border: 2px solid #B22222;
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
    .report-table-header {
        background-color: #003366;
        color: white;
        text-align: center;
        font-weight: bold;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Directory Setup
for folder in [
    "photos/students",
    "photos/gallery",
    "photos/board",
    "photos/system",
    "backups",
]:
  os.makedirs(folder, exist_ok=True)

# System Image Paths
BG_PATH = "photos/system/background.png"
LOGO_PATH = "photos/system/logo.png"
CBSE_LOGO_PATH = "photos/system/cbse_logo.png"
SEAL_PATH = "photos/system/seal.png"
SIGN_PATH = "photos/system/signature.png"
TEACHER_SIGN_PATH = "photos/system/teacher_sign.png"
PARENT_SIGN_PATH = "photos/system/parent_sign.png"
STUDENT_SIGN_PATH = "photos/system/student_sign.png"
NOTICES_FILE = "notices.json"
BOARD_TOPPERS_FILE = "board_toppers.json"
SETTINGS_FILE = "settings.json"
LOG_FILE = "result_logs.csv"

# Password Hashing & Security Helper
DEFAULT_PASS = "Jnvcu@me2"
ADMIN_PASS_HASH = hashlib.sha256(
    st.secrets.get(
        "ADMIN_PASSWORD", os.getenv("ADMIN_PASSWORD", DEFAULT_PASS)
    ).encode()
).hexdigest()

DB_FILE = "school_database.db"
EXCEL_FILE_PATH = "JNV_Student_Marks.xlsx"

ALL_SUBJECTS = [
    "Gujarati",
    "Hindi",
    "English",
    "Mathematics",
    "Science",
    "Social_Science",
    "Physics",
    "Chemistry",
    "Biology",
]

LANG_TEXTS = {
    "English": {
        "title": "STUDENT PERFORMANCE & RESULT PORTAL",
        "search_lbl": "🔎 CHECK STUDENT RESULT",
        "cert_btn": "🏆 Download Merit Certificate",
        "admit_btn": "🪪 Download Admit Card",
        "weak_alert": (
            "⚠️ Needs Special Attention in the following subjects (< 33%):"
        ),
        "chart_title": "Subject-wise Performance Breakdown",
        "combined_title": (
            "🌟 COMBINED OVERALL PERFORMANCE ACROSS ALL EXAMS"
        ),
        "trend_title": "📈 Multi-Exam Performance Trend Line",
    },
    "Hindi": {
        "title": "विद्यार्थी प्रदर्शन एवं परिणाम पोर्टल",
        "search_lbl": "🔎 छात्र परिणाम खोजें",
        "cert_btn": (
            "🏆 योग्यता प्रमाण पत्र (Merit Certificate) डाउनलोड करें"
        ),
        "admit_btn": "🪪 प्रवेश पत्र (Admit Card) डाउनलोड करें",
        "weak_alert": (
            "⚠️ निम्नलिखित विषयों में विशेष ध्यान देने की आवश्यकता है (< 33%):"
        ),
        "chart_title": "विषय-वार अंक विश्लेषण",
        "combined_title": (
            "🌟 सभी परीक्षाओं का संयुक्त प्रदर्शन (Combined Performance)"
        ),
        "trend_title": "📈 मल्टी-एग्जाम प्रोग्रेस ट्रेंड ग्राफ",
    },
    "Gujarati": {
        "title": "વિદ્યાર્થી પ્રદર્શન અને પરિણામ પોર્ટल",
        "search_lbl": "🔎 વિદ્યાર્થીનું પરિણામ જુઓ",
        "cert_btn": "🏆 મેરિટ સર્ટિફિકેટ ડાઉનલોડ કરો",
        "admit_btn": "🪪 એડમિટ કાર્ડ (Admit Card) ડાઉનલોડ કરો",
        "weak_alert": (
            "⚠️ નીચેના વિષયોમાં વિશેષ ધ્યાન આપવાની જરૂર છે (< 33%):"
        ),
        "chart_title": "વિષયવાર ગુણ વિશ્લેષણ ગ્રાફ",
        "combined_title": (
            "🌟 તમામ પરીક્ષાઓનું સંયુક્ત પ્રદર્શન (Combined Performance)"
        ),
        "trend_title": "📈 મલ્ટિ-પરીક્ષા પ્રગતિ ગ્રાફ",
    },
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
  if "." in val_str:
    val_str = val_str.split(".")[0]
  return val_str


def clean_val(val):
  s = format_clean_number(val)
  return re.sub(r"[^a-zA-Z0-9]", "", s).lower().strip()


def clean_dob_str(val):
  if pd.isna(val) or val is None:
    return ""
  s = str(val).strip()
  if not s:
    return ""
  formats_to_try = [
      "%d.%m.%Y",
      "%d-%m-%Y",
      "%d/%m/%Y",
      "%Y-%m-%d",
      "%Y/%m/%d",
      "%Y.%m.%d",
      "%d%m%Y",
      "%m/%d/%Y",
  ]
  for fmt in formats_to_try:
    try:
      dt = datetime.strptime(s, fmt)
      return dt.strftime("%Y%m%d")
    except ValueError:
      pass
  return re.sub(r"[^a-zA-Z0-9]", "", s).lower().strip()


def clean_mobile_for_wa(val):
  s = format_clean_number(val)
  digits = re.sub(r"[^0-9]", "", s)
  if len(digits) >= 10:
    return "91" + digits[-10:]
  return digits


def get_exam_priority(exam_name):
  e = str(exam_name).upper().strip()
  if "TERM END" in e or "ANNUAL" in e or "FINAL" in e:
    return 6
  elif "PWT-4" in e or "PWT 4" in e or "PWT4" in e:
    return 5
  elif "PWT-3" in e or "PWT 3" in e or "PWT3" in e:
    return 4
  elif "TERM-1" in e or "TERM 1" in e or "HALF YEARLY" in e or "MID" in e:
    return 3
  elif "PWT-2" in e or "PWT 2" in e or "PWT2" in e:
    return 2
  elif "PWT-1" in e or "PWT 1" in e or "PWT1" in e:
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
      unsafe_allow_html=True,
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


# Settings & Config Handlers (Requirement 6: Admin Enable/Disable Report Card Printing)
def load_settings():
  default_settings = {"report_card_printing_enabled": True}
  if os.path.exists(SETTINGS_FILE):
    try:
      with open(SETTINGS_FILE, "r") as f:
        return {**default_settings, **json.load(f)}
    except Exception:
      pass
  return default_settings


def save_settings(settings_dict):
  with open(SETTINGS_FILE, "w") as f:
    json.dump(settings_dict, f)


# Database Handlers
def init_db():
  conn = sqlite3.connect(DB_FILE)
  cursor = conn.cursor()
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS student_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            Class TEXT, Roll_No TEXT, Student_Name TEXT, Father_Name TEXT, Mother_Name TEXT,
            Gender TEXT, GR_No TEXT, Area TEXT, House TEXT, DOB TEXT, Aadhaar_No TEXT, Mobile_No TEXT, Exam_Type TEXT,
            Max_Marks REAL, Class_Teacher TEXT, Total_Marks REAL,
            Percentage REAL, Class_Rank INTEGER, Subject_Data TEXT,
            Attendance TEXT, Working_Days TEXT, Present_Days TEXT, Discipline TEXT, Skill_Course TEXT,
            Co_Scholastic TEXT, Bagless_Days TEXT, Outstanding_Achievement TEXT, Remarks TEXT,
            Sub_Breakdown_Data TEXT
        )
    """)

  cursor.execute("PRAGMA table_info(student_results)")
  existing_columns = [col[1] for col in cursor.fetchall()]

  new_cols = [
      ("Mother_Name", "TEXT"),
      ("Gender", "TEXT"),
      ("GR_No", "TEXT"),
      ("Area", "TEXT"),
      ("House", "TEXT"),
      ("Attendance", "TEXT"),
      ("Working_Days", "TEXT"),
      ("Present_Days", "TEXT"),
      ("Discipline", "TEXT"),
      ("Skill_Course", "TEXT"),
      ("Co_Scholastic", "TEXT"),
      ("Bagless_Days", "TEXT"),
      ("Outstanding_Achievement", "TEXT"),
      ("Remarks", "TEXT"),
      ("Sub_Breakdown_Data", "TEXT"),
  ]
  for col_name, col_type in new_cols:
    if col_name not in existing_columns:
      try:
        cursor.execute(
            f"ALTER TABLE student_results ADD COLUMN {col_name} {col_type}"
        )
      except Exception:
        pass

  conn.commit()
  conn.close()


init_db()


def sync_df_to_sqlite(df):
  if os.path.exists(DB_FILE):
    backup_file = (
        f"backups/backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
    )
    shutil.copyfile(DB_FILE, backup_file)

  conn = sqlite3.connect(DB_FILE)
  cursor = conn.cursor()
  cursor.execute("DELETE FROM student_results")
  for _, row in df.iterrows():
    sub_json = json.dumps(
        {s: row[s] for s in ALL_SUBJECTS if s in row and pd.notna(row[s])}
    )
    breakdown_json = (
        str(row.get("Sub_Breakdown_Data", "{}"))
        if "Sub_Breakdown_Data" in row
        else "{}"
    )

    cursor.execute(
        """
            INSERT INTO student_results (
                Class, Roll_No, Student_Name, Father_Name, Mother_Name, Gender, GR_No, Area, House, DOB,
                Aadhaar_No, Mobile_No, Exam_Type, Max_Marks, Class_Teacher, Total_Marks,
                Percentage, Class_Rank, Subject_Data, Attendance, Working_Days, Present_Days,
                Discipline, Skill_Course, Co_Scholastic, Bagless_Days, Outstanding_Achievement, Remarks,
                Sub_Breakdown_Data
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            str(row.get("Class", "")),
            format_clean_number(row.get("Roll_No", "")),
            str(row.get("Student_Name", "")),
            str(row.get("Father_Name", "")),
            str(row.get("Mother_Name", "")),
            str(row.get("Gender", "M")),
            str(row.get("GR_No", "01")),
            str(row.get("Area", "Rural")),
            str(row.get("House", "Aravali")),
            str(row.get("DOB", "")),
            format_clean_number(row.get("Aadhaar_No", "")),
            format_clean_number(row.get("Mobile_No", "")),
            str(row.get("Exam_Type", "")),
            float(row.get("Max_Marks", 600)),
            str(row.get("Class_Teacher", "")),
            float(row.get("Total_Marks", 0)),
            float(row.get("Percentage", 0)),
            int(row.get("Class_Rank", 0)),
            sub_json,
            str(row.get("Attendance", "95%")),
            str(row.get("Working_Days", "220")),
            str(row.get("Present_Days", "210")),
            str(row.get("Discipline", "A")),
            str(row.get("Skill_Course", "Handicraft")),
            str(row.get("Co_Scholastic", "")),
            str(row.get("Bagless_Days", "")),
            str(row.get("Outstanding_Achievement", "None")),
            str(row.get("Remarks", "Passed and Promoted")),
            breakdown_json,
        ),
    )
  conn.commit()
  conn.close()


def load_sqlite_to_df():
  conn = sqlite3.connect(DB_FILE)
  cursor = conn.cursor()
  try:
    cursor.execute(
        "SELECT Class, Roll_No, Student_Name, Father_Name, Mother_Name, Gender,"
        " GR_No, Area, House, DOB, Aadhaar_No, Mobile_No, Exam_Type, Max_Marks,"
        " Class_Teacher, Total_Marks, Percentage, Class_Rank, Subject_Data,"
        " Attendance, Working_Days, Present_Days, Discipline, Skill_Course,"
        " Co_Scholastic, Bagless_Days, Outstanding_Achievement, Remarks,"
        " Sub_Breakdown_Data FROM student_results"
    )
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
        "Class": r[0],
        "Roll_No": format_clean_number(r[1]),
        "Student_Name": r[2],
        "Father_Name": r[3],
        "Mother_Name": r[4] if r[4] else "",
        "Gender": r[5] if r[5] else "M",
        "GR_No": r[6] if r[6] else "01",
        "Area": r[7] if r[7] else "Rural",
        "House": r[8] if r[8] else "Aravali",
        "DOB": r[9],
        "Aadhaar_No": format_clean_number(r[10]),
        "Mobile_No": format_clean_number(r[11]),
        "Exam_Type": r[12],
        "Max_Marks": r[13],
        "Class_Teacher": r[14],
        "Total_Marks": r[15],
        "Percentage": r[16],
        "Class_Rank": r[17],
        "Attendance": r[19] if r[19] else "95%",
        "Working_Days": r[20] if r[20] else "220",
        "Present_Days": r[21] if r[21] else "210",
        "Discipline": r[22] if r[22] else "A",
        "Skill_Course": r[23] if r[23] else "Handicraft",
        "Co_Scholastic": r[24] if r[24] else "",
        "Bagless_Days": r[25] if r[25] else "",
        "Outstanding_Achievement": r[26] if r[26] else "None",
        "Remarks": r[27] if r[27] else "Passed and Promoted",
        "Sub_Breakdown_Data": r[28] if len(r) > 28 and r[28] else "{}",
    }
    sub_dict = json.loads(r[18]) if r[18] else {}
    for sub in ALL_SUBJECTS:
      rec[sub] = sub_dict.get(sub, np.nan)
    records.append(rec)
  return pd.DataFrame(records)


def load_notices():
  if os.path.exists(NOTICES_FILE):
    with open(NOTICES_FILE, "r") as f:
      return json.load(f)
  return ["Welcome to the PM SHRI JNV Chhotaudepur Student Portal!"]


def save_notices(notices_list):
  with open(NOTICES_FILE, "w") as f:
    json.dump(notices_list, f)


def load_board_toppers():
  if os.path.exists(BOARD_TOPPERS_FILE):
    with open(BOARD_TOPPERS_FILE, "r") as f:
      return json.load(f)
  return []


def save_board_toppers(toppers_list):
  with open(BOARD_TOPPERS_FILE, "w") as f:
    json.dump(toppers_list, f)


def log_parent_search(roll_no, student_name, selected_class):
  timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  new_data = pd.DataFrame([{
      "Timestamp": timestamp,
      "Roll_No": str(roll_no),
      "Student_Name": student_name,
      "Class": selected_class,
  }])
  if os.path.exists(LOG_FILE):
    new_data.to_csv(LOG_FILE, mode="a", header=False, index=False)
  else:
    new_data.to_csv(LOG_FILE, mode="w", header=True, index=False)


def process_data_excel(excel_file_source):
  xls = pd.ExcelFile(excel_file_source)
  sheet_names = xls.sheet_names
  if len(sheet_names) > 1:
    df_basic = pd.read_excel(xls, sheet_name=sheet_names[0])
    df_marks = pd.read_excel(xls, sheet_name=sheet_names[1])
    merge_keys = [
        c
        for c in ["Roll_No", "Class"]
        if c in df_basic.columns and c in df_marks.columns
    ]
    if not merge_keys:
      merge_keys = ["Roll_No"]
    df = pd.merge(
        df_marks, df_basic, on=merge_keys, how="left", suffixes=("", "_basic")
    )
  else:
    df = pd.read_excel(xls, sheet_name=sheet_names[0])

  meta_cols = [
      "Class",
      "Roll_No",
      "Student_Name",
      "Father_Name",
      "Mother_Name",
      "Gender",
      "GR_No",
      "Area",
      "House",
      "DOB",
      "Aadhaar_No",
      "Mobile_No",
      "Exam_Type",
      "Max_Marks",
      "Class_Teacher",
      "Attendance",
      "Working_Days",
      "Present_Days",
      "Discipline",
      "Skill_Course",
      "Co_Scholastic",
      "Bagless_Days",
      "Outstanding_Achievement",
      "Remarks",
      "Sub_Breakdown_Data",
  ]
  for col in meta_cols:
    if col not in df.columns:
      if col == "Attendance":
        df[col] = "95%"
      elif col == "Working_Days":
        df[col] = "220"
      elif col == "Present_Days":
        df[col] = "210"
      elif col == "Discipline":
        df[col] = "A"
      elif col == "Skill_Course":
        df[col] = "Handicraft"
      elif col == "Remarks":
        df[col] = "Passed and Promoted"
      elif col == "GR_No":
        df[col] = "01"
      elif col == "Area":
        df[col] = "Rural"
      elif col == "House":
        df[col] = "Aravali"
      elif col == "Gender":
        df[col] = "M"
      elif col == "Sub_Breakdown_Data":
        df[col] = "{}"
      else:
        df[col] = ""

  for col in ["Roll_No", "Aadhaar_No", "Mobile_No"]:
    if col in df.columns:
      df[col] = df[col].apply(format_clean_number)

  for sub in ALL_SUBJECTS:
    if sub not in df.columns:
      df[sub] = np.nan
    df[sub] = pd.to_numeric(df[sub], errors="coerce")

  df["Total_Marks"] = df[ALL_SUBJECTS].sum(axis=1, skipna=True)
  if "Max_Marks" not in df.columns or df["Max_Marks"].isnull().all():
    df["Max_Marks"] = df["Exam_Type"].apply(
        lambda x: 150 if "PWT" in str(x).upper() else 600
    )

  df["Max_Marks"] = pd.to_numeric(df["Max_Marks"], errors="coerce").fillna(600)
  df["Percentage"] = (df["Total_Marks"] / df["Max_Marks"]) * 100
  df["Percentage"] = df["Percentage"].round(2)
  df["Class_Rank"] = (
      df.groupby(["Class", "Exam_Type"])["Total_Marks"]
      .rank(ascending=False, method="min")
      .fillna(0)
      .astype(int)
  )

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


# ReportLab Decorative Canvas Callback with Exact Marksheet Red Border
def create_watermark_callback(
    watermark_text, with_border=True, border_color="#B22222"
):
  def draw_canvas(canvas, doc):
    canvas.saveState()
    if with_border:
      canvas.setStrokeColor(colors.HexColor(border_color))
      canvas.setLineWidth(2.5)
      canvas.rect(10, 10, doc.pagesize[0] - 20, doc.pagesize[1] - 20)
      canvas.setLineWidth(0.8)
      canvas.rect(14, 14, doc.pagesize[0] - 28, doc.pagesize[1] - 28)

    canvas.setFont("Helvetica-Bold", 36)
    canvas.setFillColor(colors.HexColor("#E0E0E0"), alpha=0.18)
    canvas.translate(doc.pagesize[0] / 2.0, doc.pagesize[1] / 2.0)
    canvas.rotate(35)
    canvas.drawCentredString(0, 0, watermark_text)
    canvas.restoreState()

  return draw_canvas


def generate_merit_certificate_pdf(student_info, exam_type, percentage, rank):
  buffer = io.BytesIO()
  doc = SimpleDocTemplate(
      buffer,
      pagesize=landscape(A4),
      rightMargin=35,
      leftMargin=35,
      topMargin=35,
      bottomMargin=35,
  )
  story = []
  styles = getSampleStyleSheet()

  cert_title = ParagraphStyle(
      "CertTitle",
      parent=styles["Heading1"],
      fontName="Helvetica-Bold",
      fontSize=22,
      leading=26,
      alignment=1,
      textColor=colors.HexColor("#B22222"),
  )
  sub_title = ParagraphStyle(
      "SubTitle",
      parent=styles["Heading2"],
      fontName="Helvetica-Bold",
      fontSize=15,
      leading=18,
      alignment=1,
      textColor=colors.HexColor("#2E7D32"),
  )
  body_style = ParagraphStyle(
      "Body",
      parent=styles["Normal"],
      fontName="Helvetica",
      fontSize=13,
      leading=22,
      alignment=1,
  )

  logo_w, logo_h = 50, 50
  left_logo = (
      RLImage(LOGO_PATH, width=logo_w, height=logo_h)
      if os.path.exists(LOGO_PATH)
      else Paragraph("", body_style)
  )
  right_logo = (
      RLImage(CBSE_LOGO_PATH, width=logo_w, height=logo_h)
      if os.path.exists(CBSE_LOGO_PATH)
      else (
          RLImage(LOGO_PATH, width=logo_w, height=logo_h)
          if os.path.exists(LOGO_PATH)
          else Paragraph("", body_style)
      )
  )

  header_text = Paragraph(
      "PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR", cert_title
  )
  h_table = Table(
      [[left_logo, header_text, right_logo]], colWidths=[60, 600, 60]
  )
  h_table.setStyle(
      TableStyle([
          ("ALIGN", (0, 0), (0, -1), "LEFT"),
          ("ALIGN", (1, 0), (1, -1), "CENTER"),
          ("ALIGN", (2, 0), (2, -1), "RIGHT"),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
      ])
  )
  story.append(h_table)
  story.append(Spacer(1, 10))

  story.append(
      Paragraph("🏆 CERTIFICATE OF ACADEMIC EXCELLENCE 🏆", sub_title)
  )
  story.append(Spacer(1, 15))

  cert_text = f"""
    This is to proudly certify that <b>{student_info['Student_Name']}</b>, Son/Daughter of <b>{student_info['Father_Name']}</b>, 
    studying in <b>Class {student_info['Class']}</b> (Roll No: <b>{student_info['Roll_No']}</b>), has secured 
    <font color="#B22222"><b>RANK #{rank}</b></font> with an outstanding score of <b>{percentage:.2f}%</b> 
    in the <b>{exam_type}</b> Examination (Academic Session 2024-25).
    """
  story.append(Paragraph(cert_text, body_style))
  story.append(Spacer(1, 20))

  parent_sign_elem = (
      RLImage(PARENT_SIGN_PATH, width=65, height=30)
      if os.path.exists(PARENT_SIGN_PATH)
      else Paragraph("____________________", body_style)
  )
  teacher_sign_elem = (
      RLImage(TEACHER_SIGN_PATH, width=65, height=30)
      if os.path.exists(TEACHER_SIGN_PATH)
      else Paragraph("____________________", body_style)
  )
  principal_sign_elem = (
      RLImage(SIGN_PATH, width=65, height=30)
      if os.path.exists(SIGN_PATH)
      else Paragraph("____________________", body_style)
  )

  sig_data = [
      [parent_sign_elem, teacher_sign_elem, principal_sign_elem],
      [
          Paragraph("<b>Parent's Signature</b>", body_style),
          Paragraph("<b>Class Teacher Signature</b>", body_style),
          Paragraph("<b>Principal Signature</b>", body_style),
      ],
  ]
  sig_table = Table(sig_data, colWidths=[240, 240, 240])
  sig_table.setStyle(
      TableStyle([
          ("ALIGN", (0, 0), (-1, -1), "CENTER"),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
      ])
  )
  story.append(sig_table)

  watermark_fn = create_watermark_callback(
      "PM SHRI JNV CHHOTAUDEPUR", with_border=True, border_color="#B22222"
  )
  doc.build(story, onFirstPage=watermark_fn, onLaterPages=watermark_fn)
  buffer.seek(0)
  return buffer.getvalue()


# ==============================================================================
# UPDATED 1-PAGE COMPACT PDF SCORECARD GENERATOR (Strict 1-Page & Blank Handling)
# Requirements 1, 2, 3, 4, 5, 10 addressed
# ==============================================================================
def generate_pdf_scorecard(student_info, filtered_df):
  buffer = io.BytesIO()
  # Strict 1-page margins & A4 size
  doc = SimpleDocTemplate(
      buffer,
      pagesize=A4,
      rightMargin=10,
      leftMargin=10,
      topMargin=10,
      bottomMargin=10,
  )
  story = []
  styles = getSampleStyleSheet()

  small_p = ParagraphStyle(
      "SmallP", parent=styles["Normal"], fontSize=6.5, leading=7.5
  )
  small_center = ParagraphStyle(
      "SmallC", parent=styles["Normal"], fontSize=6.5, leading=7.5, alignment=1
  )
  bold_center = ParagraphStyle(
      "BoldC",
      parent=styles["Normal"],
      fontName="Helvetica-Bold",
      fontSize=7,
      leading=8,
      alignment=1,
  )

  logo_w, logo_h = 36, 36
  left_logo = (
      RLImage(LOGO_PATH, width=logo_w, height=logo_h)
      if os.path.exists(LOGO_PATH)
      else Paragraph("", small_p)
  )
  right_logo = (
      RLImage(CBSE_LOGO_PATH, width=logo_w, height=logo_h)
      if os.path.exists(CBSE_LOGO_PATH)
      else (
          RLImage(LOGO_PATH, width=logo_w, height=logo_h)
          if os.path.exists(LOGO_PATH)
          else Paragraph("", small_p)
      )
  )

  header_text = Paragraph(
      "<b><font size=8.5 color='#B22222'>पीएम श्री स्कूल जवाहर नवोदय विद्यालय"
      " छोटाउदेपुर</font></b><br/><b><font size=10 color='#003366'>PM SHRI"
      " SCHOOL JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR</font></b><br/><font"
      " size=5.5 color='#B22222'>A UNIT OF NAVODAYA VIDYALAYA SAMITI, AN"
      " AUTONOMOUS BODY UNDER MINISTRY OF EDUCATION (DoEL) GOVT. OF"
      " INDIA</font>",
      ParagraphStyle("HCenter", alignment=1, leading=9),
  )

  header_table = Table(
      [[left_logo, header_text, right_logo]], colWidths=[40, 515, 40]
  )
  header_table.setStyle(
      TableStyle([
          ("ALIGN", (0, 0), (-1, -1), "CENTER"),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
      ])
  )
  story.append(header_table)
  story.append(Spacer(1, 2))

  # Affiliation details line (Requirements 2, 3, 4, 5)
  aff_data = [
      [
          Paragraph("<b>CBSE AFFILIATION NO.</b> : 440151", small_p),
          Paragraph("<b>CONTACT NO.</b> : 02669-222120", small_p),
      ],
      [
          Paragraph("<b>CBSE SCHOOL CODE</b> : 14303", small_p),
          Paragraph("<b>E-MAIL ID</b> : jnvchhotaudepur@gmail.com", small_p),
      ],
      [
          Paragraph("<b>SCHOOL UDISE CODE</b> : 24320501310", small_p),
          Paragraph(
              "<b>WEBSITE</b> : navodaya.gov.in/nvs/nvs-school/CHHOTAUDEPUR",
              small_p,
          ),
      ],
  ]
  aff_table = Table(aff_data, colWidths=[260, 235])
  aff_table.setStyle(
      TableStyle([
          ("PADDING", (0, 0), (-1, -1), 0.5),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
      ])
  )
  story.append(aff_table)
  story.append(Spacer(1, 2))

  # Title Banner
  banner = Table(
      [[
          Paragraph(
              "<b>:: REPORT CARD ::</b>",
              ParagraphStyle(
                  "Banner",
                  alignment=1,
                  textColor=colors.white,
                  fontSize=8.5,
                  leading=10,
              ),
          )
      ]],
      colWidths=[575],
  )
  banner.setStyle(
      TableStyle([
          ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#003366")),
          ("PADDING", (0, 0), (-1, -1), 2),
      ])
  )
  story.append(banner)
  story.append(Spacer(1, 2))

  # Student Info Block & Passport Photo box
  photo_path = f"photos/students/{student_info['Roll_No']}.png"
  photo_elem = (
      RLImage(photo_path, width=42, height=50)
      if os.path.exists(photo_path)
      else Paragraph(
          "<br/>Photo",
          ParagraphStyle("PPhoto", alignment=1, fontSize=6, textColor=colors.gray),
      )
  )

  info_rows = [
      [
          Paragraph(
              f"<b>G.R. No.</b> : {student_info.get('GR_No', '01')}", small_p
          ),
          Paragraph(
              f"<b>Student Name</b> : <b>{student_info['Student_Name']}</b>",
              small_p,
          ),
          photo_elem,
      ],
      [
          Paragraph(f"<b>Class</b> : {student_info['Class']}", small_p),
          Paragraph(
              f"<b>Date of Birth</b> : {student_info['DOB']}", small_p
          ),
          "",
      ],
      [
          Paragraph(f"<b>Roll No.</b> : {student_info['Roll_No']}", small_p),
          Paragraph(
              f"<b>Mother Name</b> : {student_info.get('Mother_Name', '')}",
              small_p,
          ),
          "",
      ],
      [
          Paragraph(
              f"<b>Gender</b> : {student_info.get('Gender', 'M')}", small_p
          ),
          Paragraph(
              f"<b>Father Name</b> : {student_info['Father_Name']}", small_p
          ),
          "",
      ],
      [
          Paragraph(
              f"<b>House/Area</b> : {student_info.get('House', 'Aravali')} /"
              f" {student_info.get('Area', 'Rural')}",
              small_p,
          ),
          Paragraph(
              "<b>Parents Contact</b> :"
              f" {format_clean_number(student_info.get('Mobile_No', ''))}",
              small_p,
          ),
          "",
      ],
  ]
  info_table = Table(info_rows, colWidths=[150, 365, 60])
  info_table.setStyle(
      TableStyle([
          ("SPAN", (2, 0), (2, 4)),
          ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#003366")),
          ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
          ("PADDING", (0, 0), (-1, -1), 1.5),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("ALIGN", (2, 0), (2, 4), "CENTER"),
      ])
  )
  story.append(info_table)
  story.append(Spacer(1, 2))

  # PART A Header Banner
  part_a_head = Table(
      [[
          Paragraph(
              "<b>PART A : SCHOLASTIC AREA</b>",
              ParagraphStyle(
                  "PartA",
                  alignment=1,
                  textColor=colors.white,
                  fontSize=7.5,
                  leading=9,
              ),
          )
      ]],
      colWidths=[575],
  )
  part_a_head.setStyle(
      TableStyle([
          ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#006600")),
          ("PADDING", (0, 0), (-1, -1), 1.5),
      ])
  )
  story.append(part_a_head)

  # Scholastic Area Dynamic Detailed Table (Requirement 10: Blank if missing data)
  schol_head = [
      [
          "SUBJECT",
          "TERM-1 (100 MARKS)",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "TERM-2 (100 MARKS)",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "",
          "Grand Total\n(K+U)\n(60+40)\n(100)",
          "%",
          "Grade",
          "Sub Rank",
      ],
      [
          "",
          "PWT-1\n(40)",
          "PWT-2\n(40)",
          "Best 1-2\n(20)",
          "Mult. Ass.\n(10)",
          "Portf.\n(10)",
          "Sub. Enr.\n(10)",
          "Half Yly\n(80)",
          "Half Yly\n(50)",
          "Marks\n(100)",
          "Term-1\n(40%)",
          "PWT-3\n(40)",
          "PWT-4\n(40)",
          "Best 3-4\n(20)",
          "Mult. Ass.\n(10)",
          "Portf.\n(10)",
          "Sub. Enr.\n(10)",
          "Yearly\n(80)",
          "Yearly\n(50)",
          "Marks\n(100)",
          "Term-2\n(60%)",
          "",
          "",
          "",
          "",
      ],
  ]

  schol_rows = []
  active_subs = [
      s
      for s in ALL_SUBJECTS
      if s in student_info and pd.notna(student_info[s])
  ]

  def calculate_grade(pct):
    if pct >= 91:
      return "A1"
    if pct >= 81:
      return "A2"
    if pct >= 71:
      return "B1"
    if pct >= 61:
      return "B2"
    if pct >= 51:
      return "C1"
    if pct >= 41:
      return "C2"
    if pct >= 33:
      return "D"
    return "E"

  highest_sub = "-"
  lowest_sub = "-"
  max_s_mark = -1
  min_s_mark = 999

  breakdown_dict = {}
  try:
    if (
        "Sub_Breakdown_Data" in student_info
        and student_info["Sub_Breakdown_Data"]
    ):
      breakdown_dict = json.loads(str(student_info["Sub_Breakdown_Data"]))
  except Exception:
    breakdown_dict = {}

  for sub in active_subs:
    m_val = float(student_info[sub])
    if m_val > max_s_mark:
      max_s_mark = m_val
      highest_sub = sub
    if m_val < min_s_mark:
      min_s_mark = m_val
      lowest_sub = sub

    grd = calculate_grade(m_val)
    s_bk = breakdown_dict.get(sub, {})

    # Strict check: If breakdown/exam values are missing in sheet, leave blank or 0 as per Requirement 10
    pwt1 = s_bk.get("pwt1", "")
    pwt2 = s_bk.get("pwt2", "")
    best12 = (
        round(max(float(pwt1), float(pwt2)) * 0.5, 1)
        if pwt1 != "" and pwt2 != ""
        else ""
    )
    ma = s_bk.get("ma1", "")
    port = s_bk.get("port1", "")
    sea = s_bk.get("sea1", "")
    hy80 = s_bk.get("hy80", "")
    hy50 = s_bk.get("hy50", "")
    mo100 = s_bk.get("mo100_1", round(m_val, 1))
    t1_40 = round(m_val * 0.4, 1) if m_val else ""

    pwt3 = s_bk.get("pwt3", "")
    pwt4 = s_bk.get("pwt4", "")
    best34 = (
        round(max(float(pwt3), float(pwt4)) * 0.5, 1)
        if pwt3 != "" and pwt4 != ""
        else ""
    )
    ma2 = s_bk.get("ma2", "")
    port2 = s_bk.get("port2", "")
    sea2 = s_bk.get("sea2", "")
    yr80 = s_bk.get("yr80", "")
    yr50 = s_bk.get("yr50", "")
    mo100_2 = s_bk.get("mo100_2", round(m_val, 1))
    t2_60 = round(m_val * 0.6, 1) if m_val else ""

    grand_total = (
        round(float(t1_40) + float(t2_60), 1)
        if t1_40 != "" and t2_60 != ""
        else round(m_val, 1)
    )
    pct_str = f"{m_val:.1f}"

    schol_rows.append([
        sub.upper(),
        str(pwt1),
        str(pwt2),
        str(best12),
        str(ma),
        str(port),
        str(sea),
        str(hy80),
        str(hy50),
        str(mo100),
        str(t1_40),
        str(pwt3),
        str(pwt4),
        str(best34),
        str(ma2),
        str(port2),
        str(sea2),
        str(yr80),
        str(yr50),
        str(mo100_2),
        str(t2_60),
        str(grand_total),
        pct_str,
        grd,
        "1",
    ])

  schol_data = schol_head + schol_rows
  schol_data.append([
      "CLASS RANK :",
      f"{student_info['Class_Rank']}",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "OVERALL",
      "",
      f"{int(student_info['Total_Marks'])}",
      f"{student_info['Percentage']:.1f}%",
      calculate_grade(student_info["Percentage"]),
      "1",
  ])

  schol_table = Table(schol_data, colWidths=[70] + [20] * 20 + [25, 20, 20, 20])
  schol_table.setStyle(
      TableStyle([
          ("SPAN", (0, 0), (0, 1)),
          ("SPAN", (1, 0), (10, 0)),
          ("SPAN", (11, 0), (20, 0)),
          ("SPAN", (21, 0), (21, 1)),
          ("SPAN", (22, 0), (22, 1)),
          ("SPAN", (23, 0), (23, 1)),
          ("SPAN", (24, 0), (24, 1)),
          ("SPAN", (1, -1), (18, -1)),
          ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#E0F7FA")),
          ("BACKGROUND", (1, 0), (10, 1), colors.HexColor("#00E676")),
          ("BACKGROUND", (11, 0), (20, 1), colors.HexColor("#FFB74D")),
          ("BACKGROUND", (21, 0), (24, 1), colors.HexColor("#FFF176")),
          ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#444444")),
          ("FONTSIZE", (0, 0), (-1, -1), 5),
          ("ALIGN", (0, 0), (-1, -1), "CENTER"),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("PADDING", (0, 0), (-1, -1), 0.5),
      ])
  )
  story.append(schol_table)
  story.append(Spacer(1, 2))

  # Skill Course Line
  story.append(
      Paragraph(
          f"<b>A 1 - Skill Course :</b> <font color='#006600'>"
          f"{student_info.get('Skill_Course', 'Handicraft')}</font>",
          small_p,
      )
  )
  story.append(Spacer(1, 2))

  # Co-scholastic & Bagless Days parsing
  co_sch_raw = str(student_info.get("Co_Scholastic", ""))
  t1_art, t2_art, t1_health, t2_health, t1_comm, t2_comm = (
      "A",
      "A",
      "A",
      "A",
      "A",
      "A",
  )
  if "Art:" in co_sch_raw:
    try:
      parts = co_sch_raw.split("|")
      for p in parts:
        if "Art:" in p:
          t1_art, t2_art = p.replace("Art:", "").strip().split("/")
        elif "Health:" in p:
          t1_health, t2_health = p.replace("Health:", "").strip().split("/")
        elif "Comm:" in p:
          t1_comm, t2_comm = p.replace("Comm:", "").strip().split("/")
    except Exception:
      pass

  bag_raw = str(student_info.get("Bagless_Days", ""))
  b_part, b_vac, b_sch = "Yes", "05", "05"
  if "|" in bag_raw:
    try:
      b_part, b_vac, b_sch = [x.strip() for x in bag_raw.split("|")]
    except Exception:
      pass

  # PART B, C Horizontal Matrix
  part_b_table = Table(
      [
          [
              Paragraph(
                  "<b>PART B : CO-SCHOLASTIC (3-Pt A-C)</b>", bold_center
              ),
              Paragraph("<b>PART C : 10 BAGLESS DAYS</b>", bold_center),
          ],
          [
              Table(
                  [
                      ["Co-Scholastic Area", "Term-1", "Term-2"],
                      ["Community Service", t1_comm, t2_comm],
                      ["Art Education", t1_art, t2_art],
                      ["Health & Physical Ed.", t1_health, t2_health],
                  ],
                  colWidths=[175, 75, 75],
              ),
              Table(
                  [
                      ["Participation", "Vacation", "School"],
                      [b_part, b_vac, b_sch],
                  ],
                  colWidths=[80, 85, 85],
              ),
          ],
      ],
      colWidths=[325, 250],
  )
  part_b_table.setStyle(
      TableStyle([
          ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#FFB74D")),
          ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#00E676")),
          ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#003366")),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("PADDING", (0, 0), (-1, -1), 1),
          ("FONTSIZE", (0, 0), (-1, -1), 5.5),
      ])
  )
  story.append(part_b_table)
  story.append(Spacer(1, 2))

  # PART D, E, F & Remarks Combined Compact Box
  part_d_f = Table(
      [
          [
              Paragraph("<b>PART D : DISCIPLINE</b>", bold_center),
              Paragraph("<b>PART F : ATTENDANCE</b>", bold_center),
              Paragraph("<b>PART E : OUTSTANDING ACHIEVEMENT</b>", bold_center),
          ],
          [
              Table(
                  [
                      ["Area", "T-1", "T-2"],
                      [
                          "Discipline",
                          student_info.get("Discipline", "A"),
                          student_info.get("Discipline", "A"),
                      ],
                  ],
                  colWidths=[90, 45, 45],
              ),
              Table(
                  [
                      ["Working", "Present", "%"],
                      [
                          student_info.get("Working_Days", "220"),
                          student_info.get("Present_Days", "210"),
                          student_info.get("Attendance", "95%"),
                      ],
                  ],
                  colWidths=[50, 50, 50],
              ),
              Paragraph(
                  f"{student_info.get('Outstanding_Achievement', 'None')}",
                  small_p,
              ),
          ],
      ],
      colWidths=[180, 150, 245],
  )
  part_d_f.setStyle(
      TableStyle([
          ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#AB47BC")),
          ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#E53935")),
          ("BACKGROUND", (2, 0), (2, 0), colors.HexColor("#42A5F5")),
          ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
          ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#003366")),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("PADDING", (0, 0), (-1, -1), 1),
          ("FONTSIZE", (0, 0), (-1, -1), 5.5),
      ])
  )
  story.append(part_d_f)
  story.append(Spacer(1, 2))

  # Remarks & Signatures Compact Grid
  remark_text = student_info.get("Remarks", "Passed and Promoted")
  t_sign = (
      RLImage(TEACHER_SIGN_PATH, width=45, height=18)
      if os.path.exists(TEACHER_SIGN_PATH)
      else Paragraph("", small_p)
  )
  p_sign = (
      RLImage(SIGN_PATH, width=45, height=18)
      if os.path.exists(SIGN_PATH)
      else Paragraph("", small_p)
  )
  par_sign = (
      RLImage(PARENT_SIGN_PATH, width=45, height=18)
      if os.path.exists(PARENT_SIGN_PATH)
      else Paragraph("", small_p)
  )

  bottom_grid = Table(
      [
          [
              Paragraph(
                  f"<b>Remarks:</b> <font color='#B22222'>{remark_text}</font>",
                  small_p,
              ),
              Paragraph(
                  f"<b>Result:</b> <font color='green'><b>Passed & Promoted to"
                  f" Class {student_info['Class']}</b></font>",
                  small_p,
              ),
          ],
          [t_sign, p_sign],
          [
              Paragraph("<b>Class Teacher Signature</b>", small_center),
              Paragraph("<b>Principal Signature & Seal</b>", small_center),
          ],
      ],
      colWidths=[285, 290],
  )
  bottom_grid.setStyle(
      TableStyle([
          ("SPAN", (0, 0), (1, 0)),
          ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#003366")),
          ("ALIGN", (0, 0), (-1, -1), "CENTER"),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("PADDING", (0, 0), (-1, -1), 1),
      ])
  )
  story.append(bottom_grid)

  watermark_fn = create_watermark_callback(
      "PM SHRI JNV CHHOTAUDEPUR", with_border=True, border_color="#B22222"
  )
  doc.build(story, onFirstPage=watermark_fn, onLaterPages=watermark_fn)
  buffer.seek(0)
  return buffer.getvalue()


def generate_admit_card_pdf(student_info):
  buffer = io.BytesIO()
  doc = SimpleDocTemplate(
      buffer,
      pagesize=A4,
      rightMargin=30,
      leftMargin=30,
      topMargin=30,
      bottomMargin=30,
  )
  story = []
  styles = getSampleStyleSheet()

  title_style = ParagraphStyle(
      "DocTitle",
      parent=styles["Heading1"],
      fontName="Helvetica-Bold",
      fontSize=16,
      leading=18,
      alignment=1,
      textColor=colors.HexColor("#B22222"),
  )
  sub_style = ParagraphStyle(
      "DocSub",
      parent=styles["Heading2"],
      fontName="Helvetica-Bold",
      fontSize=12,
      leading=14,
      alignment=1,
      textColor=colors.HexColor("#2E7D32"),
  )
  normal_style = styles["Normal"]

  logo_w, logo_h = 45, 45
  left_logo = (
      RLImage(LOGO_PATH, width=logo_w, height=logo_h)
      if os.path.exists(LOGO_PATH)
      else Paragraph("", normal_style)
  )
  right_logo = (
      RLImage(CBSE_LOGO_PATH, width=logo_w, height=logo_h)
      if os.path.exists(CBSE_LOGO_PATH)
      else (
          RLImage(LOGO_PATH, width=logo_w, height=logo_h)
          if os.path.exists(LOGO_PATH)
          else Paragraph("", normal_style)
      )
  )

  header_p = Paragraph(
      "PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR", title_style
  )
  h_table = Table(
      [[left_logo, header_p, right_logo]], colWidths=[55, 430, 55]
  )
  h_table.setStyle(
      TableStyle([
          ("ALIGN", (0, 0), (0, -1), "LEFT"),
          ("ALIGN", (1, 0), (1, -1), "CENTER"),
          ("ALIGN", (2, 0), (2, -1), "RIGHT"),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
      ])
  )
  story.append(h_table)
  story.append(
      Paragraph(
          "EXAMINATION ADMIT CARD / HALL TICKET (2024-25)", sub_style
      )
  )
  story.append(Spacer(1, 10))

  photo_path = f"photos/students/{student_info['Roll_No']}.png"
  photo_elem = (
      RLImage(photo_path, width=70, height=80)
      if os.path.exists(photo_path)
      else Paragraph("<b>[PHOTO]</b>", normal_style)
  )

  info_data = [
      [
          Paragraph(
              f"<b>Student Name:</b> {student_info['Student_Name']}",
              normal_style,
          ),
          photo_elem,
      ],
      [
          Paragraph(
              f"<b>Roll No:</b> {student_info['Roll_No']}", normal_style
          ),
          "",
      ],
      [Paragraph(f"<b>Class:</b> {student_info['Class']}", normal_style), ""],
      [
          Paragraph(
              f"<b>Father's Name:</b> {student_info['Father_Name']}",
              normal_style,
          ),
          "",
      ],
      [
          Paragraph(
              "<b>Exam Center:</b> JNV Chhotaudepur Main Campus",
              normal_style,
          ),
          "",
      ],
  ]
  t = Table(info_data, colWidths=[380, 120])
  t.setStyle(
      TableStyle([
          ("SPAN", (1, 0), (1, 4)),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
          ("PADDING", (0, 0), (-1, -1), 5),
      ])
  )
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

  parent_sign_elem = (
      RLImage(PARENT_SIGN_PATH, width=60, height=28)
      if os.path.exists(PARENT_SIGN_PATH)
      else Paragraph("________________", normal_style)
  )
  teacher_sign_elem = (
      RLImage(TEACHER_SIGN_PATH, width=60, height=28)
      if os.path.exists(TEACHER_SIGN_PATH)
      else Paragraph("________________", normal_style)
  )
  principal_sign_elem = (
      RLImage(SIGN_PATH, width=60, height=28)
      if os.path.exists(SIGN_PATH)
      else Paragraph("________________", normal_style)
  )

  sig_data = [
      [parent_sign_elem, teacher_sign_elem, principal_sign_elem],
      [
          Paragraph("<b>Student/Parent Sign</b>", normal_style),
          Paragraph("<b>Class Teacher Sign</b>", normal_style),
          Paragraph("<b>Principal Signature</b>", normal_style),
      ],
  ]
  sig_table = Table(sig_data, colWidths=[180, 160, 180])
  sig_table.setStyle(
      TableStyle([
          ("ALIGN", (0, 0), (-1, -1), "CENTER"),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
      ])
  )
  story.append(sig_table)

  watermark_fn = create_watermark_callback(
      "PM SHRI JNV CHHOTAUDEPUR", with_border=True, border_color="#B22222"
  )
  doc.build(story, onFirstPage=watermark_fn, onLaterPages=watermark_fn)
  buffer.seek(0)
  return buffer.getvalue()


# Header Layout on Web Site
h_col1, h_col2, h_col3 = st.columns(
    [1.2, 5.6, 1.2], vertical_alignment="center"
)
with h_col1:
  if os.path.exists(LOGO_PATH):
    st.markdown(
        f'<div class="header-logo-left"><img'
        f' src="data:image/png;base64,{get_base64_image(LOGO_PATH)}"'
        ' width="80"></div>',
        unsafe_allow_html=True,
    )
  elif os.path.exists(CBSE_LOGO_PATH):
    st.markdown(
        f'<div class="header-logo-left"><img'
        f' src="data:image/png;base64,{get_base64_image(CBSE_LOGO_PATH)}"'
        ' width="80"></div>',
        unsafe_allow_html=True,
    )

with h_col2:
  st.markdown(
      "<h2 class='main-title'>PM SHRI JAWAHAR NAVODAYA VIDYALAYA"
      " CHHOTAUDEPUR</h2>",
      unsafe_allow_html=True,
  )

with h_col3:
  if os.path.exists(CBSE_LOGO_PATH):
    st.markdown(
        f'<div class="header-logo-right mobile-hide"><img'
        f' src="data:image/png;base64,{get_base64_image(CBSE_LOGO_PATH)}"'
        ' width="80"></div>',
        unsafe_allow_html=True,
    )
  elif os.path.exists(LOGO_PATH):
    st.markdown(
        f'<div class="header-logo-right mobile-hide"><img'
        f' src="data:image/png;base64,{get_base64_image(LOGO_PATH)}"'
        ' width="80"></div>',
        unsafe_allow_html=True,
    )

st.markdown("---")

# Navigation Menu Bar
nav_col1, nav_col2 = st.columns([4, 1.2], vertical_alignment="center")
with nav_col1:
  menu = st.radio(
      "NAVIGATION_MENU",
      [
          "👨‍🎓 PARENT PORTAL",
          "🖼️ SCHOOL GALLERY",
          "🏆 BOARD EXAM RESULTS",
          "⚙️ ADMIN PORTAL",
      ],
      horizontal=True,
      label_visibility="collapsed",
  )
with nav_col2:
  selected_lang = st.selectbox(
      "🌐 Language / भाषा",
      ["English", "Hindi", "Gujarati"],
      label_visibility="collapsed",
  )

txt = LANG_TEXTS[selected_lang]
st.markdown(
    f"<h4 class='sub-title'>{txt['title']}</h4>", unsafe_allow_html=True
)
st.markdown("---")


def render_topper_marquee(topper_list):
  if not topper_list:
    st.info("Top performers details will be displayed here once available.")
    return
  cards_html = ""
  for t in topper_list:
    img_b64 = get_base64_image(t.get("photo", ""))
    img_src = (
        f"data:image/png;base64,{img_b64}"
        if img_b64
        else "https://via.placeholder.com/80?text=Topper"
    )
    rank_badge = (
        f'<div style="font-size: 11px; background: #B22222; color: white;'
        f' border-radius: 4px; padding: 1px 4px; margin-bottom: 3px;">Rank'
        f' #{t.get("rank", "1")}</div>'
        if t.get("rank")
        else ""
    )
    card = (
        '<div class="topper-card">'
        f'<img src="{img_src}" style="width: 75px; height: 75px; border-radius:'
        ' 50%; object-fit: cover; border: 2px solid #B22222;">'
        f"{rank_badge}"
        '<div style="font-weight: bold; color: #B22222; margin-top: 3px;'
        f' font-size: 13px;">{t["name"]}</div>'
        f'<div style="font-size: 11px; color: #333;">Class {t["class"]}'
        f' ({t.get("year", "2024-25")})</div>'
        '<div style="font-size: 14px; font-weight: bold; color: #2E7D32;'
        " background: #E8F5E9; margin-top: 4px; border-radius: 4px; padding:"
        f' 2px 0;">🏆 {t["percentage"]}</div>'
        "</div>"
    )
    cards_html += card
  st.markdown(
      '<marquee direction="left" scrollamount="6" onmouseover="this.stop();"'
      f' onmouseout="this.start();">{cards_html}</marquee>',
      unsafe_allow_html=True,
  )


# ==============================================================================
# 👨‍🎓 PARENT PORTAL
# ==============================================================================
if menu == "👨‍🎓 PARENT PORTAL":
  notices = load_notices()
  if notices:
    st.markdown(
        f"""
            <div class="notice-box">
                <span style="font-weight: bold; color: #1B5E20;">📢 DIGITAL NOTICE BOARD:</span>
                <marquee direction="left" scrollamount="5" behavior="scroll" style="vertical-align: middle; margin-left: 10px;">
                    {" &nbsp;&nbsp;&nbsp;&nbsp; 🔹 &nbsp;&nbsp;&nbsp;&nbsp; ".join(notices)}
                </marquee>
            </div>
        """,
        unsafe_allow_html=True,
    )

  st.markdown("<div class='hall-of-fame-box'>", unsafe_allow_html=True)
  st.markdown(
      "<h4 style='text-align: center; color: #B22222; margin-bottom: 8px;'>🏆"
      " ACADEMIC HALL OF FAME (CURRENT SESSION TOPPERS) 🏆</h4>",
      unsafe_allow_html=True,
  )

  # Requirement 8: Distinguish Board Toppers from Current Session Hall of Fame
  current_session_toppers = []
  if (
      st.session_state["student_data"] is not None
      and not st.session_state["student_data"].empty
  ):
    df_top = st.session_state["student_data"]
    for c_val in sorted(df_top["Class"].astype(str).unique()):
      c_df = df_top[df_top["Class"].astype(str) == c_val]
      if not c_df.empty:
        top_student = c_df.sort_values(by="Percentage", ascending=False).iloc[0]
        photo_p = f"photos/students/{top_student['Roll_No']}.png"
        current_session_toppers.append({
            "name": top_student["Student_Name"],
            "class": str(top_student["Class"]),
            "percentage": f"{top_student['Percentage']:.1f}%",
            "year": "Current Session",
            "rank": "1",
            "photo": photo_p if os.path.exists(photo_p) else "",
        })

  render_topper_marquee(current_session_toppers)
  st.markdown("</div>", unsafe_allow_html=True)

  if st.session_state["student_data"] is not None:
    df_data = st.session_state["student_data"]
    if not df_data.empty and "Exam_Type" in df_data.columns:
      unique_exams = df_data["Exam_Type"].dropna().unique()
      if len(unique_exams) > 0:
        sorted_exams = sorted(
            unique_exams, key=get_exam_priority, reverse=True
        )
        latest_exam = sorted_exams[0]
        latest_df = df_data[df_data["Exam_Type"] == latest_exam].copy()
        if not latest_df.empty:
          school_topper = latest_df.sort_values(
              by="Percentage", ascending=False
          ).iloc[0]
          ticker_items = [
              f"🏆 <b>OVERALL SCHOOL TOPPER ({latest_exam}):</b>"
              f" {school_topper['Student_Name']} (Class {school_topper['Class']})"
              f" - {school_topper['Percentage']:.2f}%"
          ]

          classes = sorted(latest_df["Class"].astype(str).unique())
          for cls in classes:
            cls_toppers = (
                latest_df[latest_df["Class"].astype(str) == cls]
                .sort_values(by="Percentage", ascending=False)
                .head(3)
            )
            top_list = [
                f"{idx+1}. {r['Student_Name']} ({r['Percentage']:.1f}%)"
                for idx, (_, r) in enumerate(cls_toppers.iterrows())
            ]
            ticker_items.append(
                f"🥇 <b>Class {cls} Top 3:</b> {' | '.join(top_list)}"
            )

          st.markdown(
              f"""
                        <div style="background-color: #FFF9C4; border-left: 5px solid #FBC02D; padding: 7px 10px; border-radius: 4px; color: #000; font-size: 15px; margin-bottom: 10px;">
                            <marquee direction="left" scrollamount="6" behavior="scroll">{" &nbsp;&nbsp;&nbsp; ✦ &nbsp;&nbsp;&nbsp; ".join(ticker_items)}</marquee>
                        </div>
                    """,
              unsafe_allow_html=True,
          )

  st.header(txt["search_lbl"])

  if st.session_state["student_data"] is None:
    st.warning("⚠️ Data file not found. Kripya Admin Portal se Data Upload karein.")
  else:
    df = st.session_state["student_data"]
    search_method = st.radio(
        "Choose Verification Method:",
        [
            "Option 1: Roll No & Date of Birth (DOB)",
            "Option 2: Roll No & Aadhaar Number",
            "Option 3: OTP Based Mobile Verification (SMS/WhatsApp)",
        ],
        horizontal=True,
    )

    filtered_df = pd.DataFrame()

    if "Option 3" in search_method:
      c1, c2, c3 = st.columns([2, 2, 1])
      with c1:
        selected_class = st.selectbox(
            "Select Class", sorted(df["Class"].astype(str).unique()), key="otp_cls"
        )
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

      user_otp = st.text_input(
          "Enter 4-Digit OTP Received", key="entered_otp"
      )
      if st.button("🔍 Verify & View Result"):
        if (
            "current_otp" in st.session_state
            and user_otp == st.session_state["current_otp"]
        ):
          filtered_df = df[
              (
                  df["Class"].astype(str).str.strip().str.lower()
                  == selected_class.strip().lower()
              )
              & (df["Roll_No"].apply(clean_val) == clean_val(roll_no))
              & (df["Mobile_No"].apply(clean_val) == clean_val(mobile_input))
          ]
        else:
          st.error("❌ Incorrect OTP entered!")
    else:
      with st.form("search_form"):
        c1, c2 = st.columns(2)
        with c1:
          selected_class = st.selectbox(
              "Select Class", sorted(df["Class"].astype(str).unique())
          )
          roll_no = st.text_input("Roll No")

        if "Option 1" in search_method:
          with c2:
            dob_input = st.text_input(
                "Date of Birth (e.g. 26.12.2011, 26/12/2011 or 2011-12-26)"
            )
        else:
          with c2:
            aadhaar_input = st.text_input("Aadhaar Number")

        submit_btn = st.form_submit_button("🔍 View Result")

      if submit_btn:
        if "Option 1" in search_method:
          filtered_df = df[
              (
                  df["Class"].astype(str).str.strip().str.lower()
                  == selected_class.strip().lower()
              )
              & (df["Roll_No"].apply(clean_val) == clean_val(roll_no))
              & (df["DOB"].apply(clean_dob_str) == clean_dob_str(dob_input))
          ]
        else:
          filtered_df = df[
              (
                  df["Class"].astype(str).str.strip().str.lower()
                  == selected_class.strip().lower()
              )
              & (df["Roll_No"].apply(clean_val) == clean_val(roll_no))
              & (
                  df["Aadhaar_No"].apply(clean_val) == clean_val(aadhaar_input)
              )
          ]

    if not filtered_df.empty:
      student_info = filtered_df.iloc[0]
      log_parent_search(
          student_info["Roll_No"],
          student_info["Student_Name"],
          student_info["Class"],
      )

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
        st.write(
            f"**Student:** {student_info['Student_Name']} | **G.R. No:**"
            f" {student_info.get('GR_No', '01')} | **Roll No:**"
            f" {student_info['Roll_No']}"
        )
        st.write(
            f"**Class:** {student_info['Class']} | **House:**"
            f" {student_info.get('House', 'Aravali')} | **Aadhaar:**"
            f" {mask_aadhaar(student_info['Aadhaar_No'])}"
        )
        st.write(
            f"**Attendance:** {student_info.get('Attendance', '95%')} |"
            " **Discipline:** Grade"
            f" {student_info.get('Discipline', 'A')} | **Skill Course:**"
            f" {student_info.get('Skill_Course', 'Handicraft')}"
        )

        # Requirement 6: Check if admin enabled printing for parents portal
        settings = load_settings()
        printing_enabled = settings.get("report_card_printing_enabled", True)

        b_c1, b_c2, b_c3 = st.columns(3)
        with b_c1:
          if printing_enabled:
            st.download_button(
                "📥 Report Card (PDF)",
                data=pdf_bytes,
                file_name=f"Report_{student_info['Roll_No']}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
          else:
            st.warning(
                "🔒 Report Card download is currently disabled by Admin."
            )
        with b_c2:
          st.download_button(
              txt["admit_btn"],
              data=admit_bytes,
              file_name=f"AdmitCard_{student_info['Roll_No']}.pdf",
              mime="application/pdf",
              use_container_width=True,
          )

        best_row = filtered_df.sort_values("Class_Rank").iloc[0]
        if best_row["Class_Rank"] <= 3:
          with b_c3:
            cert_pdf = generate_merit_certificate_pdf(
                student_info,
                best_row["Exam_Type"],
                best_row["Percentage"],
                best_row["Class_Rank"],
            )
            st.download_button(
                txt["cert_btn"],
                data=cert_pdf,
                file_name=f"Merit_Certificate_{student_info['Roll_No']}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

      st.markdown("---")
      st.subheader(txt["combined_title"])
      tot_obtained = filtered_df["Total_Marks"].sum()
      tot_max = filtered_df["Max_Marks"].sum()
      overall_pct = (tot_obtained / tot_max * 100) if tot_max > 0 else 0.0

      m1, m2, m3, m4 = st.columns(4)
      m1.metric("Exams Taken", len(filtered_df))
      m2.metric("Total Marks Obtained", f"{int(tot_obtained)}")
      m3.metric("Combined Maximum Marks", f"{int(tot_max)}")
      m4.metric("Overall Percentage", f"{overall_pct:.2f}%")

      st.markdown("---")
      st.subheader(txt["trend_title"])

      if len(filtered_df) > 0:
        trend_fig = go.Figure()
        trend_fig.add_trace(
            go.Scatter(
                x=filtered_df["Exam_Type"],
                y=filtered_df["Percentage"],
                mode="lines+markers+text",
                name="Percentage",
                text=[f"{p:.1f}%" for p in filtered_df["Percentage"]],
                textposition="top center",
                line=dict(color="#B22222", width=3),
                marker=dict(size=10, color="#800000"),
            )
        )
        trend_fig.update_layout(
            title=(
                "Academic Progress Timeline for"
                f" {student_info['Student_Name']}"
            ),
            xaxis_title="Examinations",
            yaxis_title="Percentage Score (%)",
            yaxis=dict(range=[0, 105]),
            margin=dict(l=20, r=20, t=40, b=20),
        )
        st.plotly_chart(trend_fig, use_container_width=True)

      st.markdown("---")
      st.subheader(f"📊 {txt['chart_title']}")
      latest_row = filtered_df.iloc[-1]
      sub_names = [
          s for s in ALL_SUBJECTS if s in latest_row and pd.notna(latest_row[s])
      ]
      sub_marks = [latest_row[s] for s in sub_names]

      class_df = df[
          (df["Class"].astype(str) == str(latest_row["Class"]))
          & (df["Exam_Type"] == latest_row["Exam_Type"])
      ]
      class_avgs = [class_df[s].mean() for s in sub_names]

      fig = go.Figure()
      fig.add_trace(
          go.Bar(
              x=sub_names,
              y=sub_marks,
              name="Student Score",
              marker_color="#B22222",
              text=sub_marks,
              textposition="auto",
          )
      )
      fig.add_trace(
          go.Bar(
              x=sub_names,
              y=class_avgs,
              name="Class Average",
              marker_color="#FFA726",
              text=[f"{v:.1f}" for v in class_avgs],
              textposition="auto",
          )
      )

      fig.update_layout(
          barmode="group",
          title=(
              "Marks Comparison vs Class Average"
              f" ({latest_row['Exam_Type']})"
          ),
          xaxis_title="Subjects",
          yaxis_title="Marks",
          yaxis=dict(range=[0, 100]),
      )
      st.plotly_chart(fig, use_container_width=True)

      weak_subs = [
          f"{s} ({latest_row[s]} marks)"
          for s in sub_names
          if float(latest_row[s]) < 33.0
      ]
      if weak_subs:
        st.markdown(
            f"<div class='weak-badge'>{txt['weak_alert']}"
            f" {', '.join(weak_subs)}</div>",
            unsafe_allow_html=True,
        )

      st.markdown("---")
      st.subheader("📝 INDIVIDUAL EXAM SCORECARDS")
      for index, row in filtered_df.iterrows():
        with st.expander(
            f"📌 **{row['Exam_Type']}** | Score:"
            f" {int(row['Total_Marks'])}/{int(row['Max_Marks'])}"
            f" ({row['Percentage']:.2f}%) | Rank: #{row['Class_Rank']}",
            expanded=True,
        ):
          subject_rows = []
          s_no = 1
          for sub_name in ALL_SUBJECTS:
            if pd.notna(row[sub_name]):
              val = row[sub_name]
              subject_rows.append({
                  "S.No.": s_no,
                  "Subject Name": sub_name,
                  "Marks Obtained": (
                      int(val) if float(val).is_integer() else val
                  ),
              })
              s_no += 1
          st.dataframe(
              pd.DataFrame(subject_rows),
              hide_index=True,
              use_container_width=True,
          )
    elif submit_btn:
      st.error(
          "❌ No student record found matching the provided Class, Roll"
          " Number, and Credentials. Kripya details re-check karein."
      )

# ==============================================================================
# 🖼️ GALLERY & BOARD RESULTS
# ==============================================================================
elif menu == "🖼️ SCHOOL GALLERY":
  st.header("🏫 GALLERY")
  gallery_files = [
      f
      for f in os.listdir("photos/gallery")
      if f.lower().endswith((".png", ".jpg", ".jpeg"))
  ]
  if not gallery_files:
    st.info("ℹ️ Gallery empty.")
  else:
    images_html = "".join([
        '<img'
        f' src="data:image/png;base64,{get_base64_image(os.path.join("photos/gallery", img))}"'
        ' style="height: 200px; margin-right: 15px; border-radius: 8px; border:'
        ' 2px solid #B22222;">'
        for img in gallery_files
    ])
    st.markdown(
        f'<marquee direction="left" scrollamount="7">{images_html}</marquee>',
        unsafe_allow_html=True,
    )

elif menu == "🏆 BOARD EXAM RESULTS":
  st.header("🎓 CBSE BOARD TOPPERS HALL OF FAME")
  toppers_data = load_board_toppers()
  if not toppers_data:
    st.info(
        "No board toppers uploaded yet. Add board exam toppers from Admin"
        " Dashboard."
    )
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

    # Requirement 6: Print Enable/Disable Control Panel
    st.subheader("🖨️ Parent Portal Report Card Print Control")
    settings = load_settings()
    current_print_status = settings.get("report_card_printing_enabled", True)
    new_print_status = st.toggle(
        "Enable Report Card Printing on Parent Portal",
        value=current_print_status,
    )
    if new_print_status != current_print_status:
      settings["report_card_printing_enabled"] = new_print_status
      save_settings(settings)
      st.success(
          f"✅ Report Card download status updated to: {new_print_status}"
      )
      st.rerun()

    st.markdown("---")

    with st.expander(
        "📦 1. BULK CLASS RESULT PDF EXPORT & EXCEL MERIT LIST", expanded=False
    ):
      if st.session_state["student_data"] is not None:
        df_bulk = st.session_state["student_data"]
        c_col1, c_col2 = st.columns(2)
        with c_col1:
          bulk_cls = st.selectbox(
              "Select Class",
              sorted(df_bulk["Class"].astype(str).unique()),
              key="bulk_cls",
          )
        with c_col2:
          bulk_exam = st.selectbox(
              "Select Exam Type",
              sorted(df_bulk["Exam_Type"].astype(str).unique()),
              key="bulk_exam",
          )

        ex_col1, ex_col2 = st.columns(2)
        with ex_col1:
          if st.button("🚀 Generate Bulk ZIP File"):
            filtered_bulk = df_bulk[
                (df_bulk["Class"].astype(str) == str(bulk_cls))
                & (df_bulk["Exam_Type"] == bulk_exam)
            ]
            if filtered_bulk.empty:
              st.warning("No records found.")
            else:
              zip_buffer = io.BytesIO()
              with zipfile.ZipFile(
                  zip_buffer, "w", zipfile.ZIP_DEFLATED
              ) as zip_file:
                for idx, s_row in filtered_bulk.iterrows():
                  s_df = filtered_bulk[
                      filtered_bulk["Roll_No"] == s_row["Roll_No"]
                  ]
                  pdf_data = generate_pdf_scorecard(s_row, s_df)
                  pdf_filename = (
                      f"Class_{bulk_cls}_{s_row['Roll_No']}_{s_row['Student_Name'].replace(' ', '_')}.pdf"
                  )
                  zip_file.writestr(pdf_filename, pdf_data)
              zip_buffer.seek(0)
              st.download_button(
                  label=f"📥 Download ZIP (Class {bulk_cls})",
                  data=zip_buffer,
                  file_name=f"Class_{bulk_cls}_{bulk_exam}_ReportCards.zip",
                  mime="application/zip",
                  use_container_width=True,
              )
        with ex_col2:
          filtered_bulk_ex = df_bulk[
              (df_bulk["Class"].astype(str) == str(bulk_cls))
              & (df_bulk["Exam_Type"] == bulk_exam)
          ].sort_values("Class_Rank")
          if not filtered_bulk_ex.empty:
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
              filtered_bulk_ex.to_excel(
                  writer, sheet_name=f"Class_{bulk_cls}_Merit", index=False
              )
            excel_buffer.seek(0)
            st.download_button(
                label="📊 Export Merit List (Excel)",
                data=excel_buffer,
                file_name=f"MeritList_Class_{bulk_cls}_{bulk_exam}.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

    with st.expander(
        "📊 2. TEACHER-WISE & ADVANCED SUBJECT ANALYTICS", expanded=False
    ):
      if st.session_state["student_data"] is not None:
        df_t = st.session_state["student_data"]
        t_tab1, t_tab2 = st.tabs(
            ["👩‍🏫 Class Teacher Summary", "📈 Subject-Wise Analytics"]
        )

        with t_tab1:
          if (
              "Class_Teacher" in df_t.columns
              and not df_t["Class_Teacher"].isnull().all()
          ):
            teachers = sorted(df_t["Class_Teacher"].dropna().astype(str).unique())
            teacher_summary = []
            for t in teachers:
              t_df = df_t[df_t["Class_Teacher"].astype(str) == t]
              tot_students = len(t_df)
              passed_students = len(t_df[t_df["Percentage"] >= 33.0])
              pass_pct = (
                  (passed_students / tot_students * 100)
                  if tot_students > 0
                  else 0.0
              )
              avg_score = t_df["Percentage"].mean()
              teacher_summary.append({
                  "Class Teacher": t,
                  "Class(es) Assigned": (
                      ", ".join(t_df["Class"].astype(str).unique())
                  ),
                  "Total Students": tot_students,
                  "Passed Students": passed_students,
                  "Pass Percentage (%)": f"{pass_pct:.2f}%",
                  "Average Class Score (%)": f"{avg_score:.2f}%",
              })
            st.dataframe(
                pd.DataFrame(teacher_summary),
                hide_index=True,
                use_container_width=True,
            )

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
                  "Pass % (>=33)": (
                      f"{round((s_data[s_data >= 33].count() / len(s_data)) * 100, 1)}%"
                  ),
              })
          st.dataframe(
              pd.DataFrame(sub_stats),
              hide_index=True,
              use_container_width=True,
          )

    with st.expander(
        "✏️ 3. EDIT STUDENT DATA, BULK UPLOAD & REPORT CARD DETAILS",
        expanded=False,
    ):
      # Requirement 7: Excel Bulk Upload and Detailed Student Update
      st.markdown(
          "##### 📁 Bulk Upload Student Details & Marks via Excel File"
      )
      uploaded_excel = st.file_uploader(
          "Upload Student Excel File (.xlsx)",
          type=["xlsx"],
          key="bulk_excel_upload",
      )
      if st.button("📥 Process & Sync Excel Database") and uploaded_excel:
        try:
          df_new = process_data_excel(uploaded_excel)
          st.session_state["student_data"] = df_new
          st.success(
              "✅ Successfully uploaded and synced student database from"
              " Excel!"
          )
          st.rerun()
        except Exception as e:
          st.error(f"❌ Error processing Excel file: {e}")

      st.markdown("---")
      if st.session_state["student_data"] is not None:
        st.markdown("##### 📝 Detailed Student Report Card Attribute Manager")
        sel_roll = st.selectbox(
            "Select Student (by Roll No) to Edit Specific Report Card Extra"
            " Data:",
            sorted(
                st.session_state["student_data"]["Roll_No"].astype(str).unique()
            ),
            key="rep_edit_roll",
        )

        st_idx = st.session_state["student_data"][
            st.session_state["student_data"]["Roll_No"].astype(str) == sel_roll
        ].index

        if not st_idx.empty:
          st_row = st.session_state["student_data"].loc[st_idx[0]]
          with st.form("extra_report_card_form"):
            e_c1, e_c2, e_c3 = st.columns(3)
            with e_c1:
              new_gr = st.text_input(
                  "G.R. No.", value=str(st_row.get("GR_No", "01"))
              )
              new_house = st.selectbox(
                  "House",
                  ["Aravali", "Nilgiri", "Shivalik", "Udaigiri"],
                  index=[
                      "Aravali",
                      "Nilgiri",
                      "Shivalik",
                      "Udaigiri",
                  ].index(str(st_row.get("House", "Aravali")))
                  if str(st_row.get("House", "Aravali"))
                  in ["Aravali", "Nilgiri", "Shivalik", "Udaigiri"]
                  else 0,
              )
              new_area = st.selectbox(
                  "Area",
                  ["Rural", "Urban"],
                  index=0
                  if str(st_row.get("Area", "Rural")) == "Rural"
                  else 1,
              )
              new_skill = st.text_input(
                  "Skill Course Opted",
                  value=str(st_row.get("Skill_Course", "Handicraft")),
              )
            with e_c2:
              new_disc = st.selectbox(
                  "Discipline Grade",
                  ["A", "B", "C"],
                  index=["A", "B", "C"].index(
                      str(st_row.get("Discipline", "A"))
                  )
                  if str(st_row.get("Discipline", "A")) in ["A", "B", "C"]
                  else 0,
              )
              new_work_days = st.text_input(
                  "Working Days", value=str(st_row.get("Working_Days", "220"))
              )
              new_pres_days = st.text_input(
                  "Present Days", value=str(st_row.get("Present_Days", "210"))
              )
              new_att_pct = st.text_input(
                  "Attendance %", value=str(st_row.get("Attendance", "95%"))
              )
            with e_c3:
              new_achieve = st.text_input(
                  "Outstanding Achievement",
                  value=str(
                      st_row.get("Outstanding_Achievement", "None")
                  ),
              )
              new_remarks = st.text_input(
                  "Teacher Remarks",
                  value=str(st_row.get("Remarks", "Passed and Promoted")),
              )
              new_co_sch = st.text_input(
                  "Co-Scholastic Grades (Art:A/A | Health:A/A | Comm:A/A)",
                  value=str(
                      st_row.get(
                          "Co_Scholastic",
                          "Art: A/A | Health: A/A | Comm: A/A",
                      )
                  ),
              )
              new_bagless = st.text_input(
                  "Bagless Days (Participation | Vacation | School)",
                  value=str(st_row.get("Bagless_Days", "Yes | 05 | 05")),
              )

            save_extra_btn = st.form_submit_button(
                "💾 Update Report Card Attributes"
            )
            if save_extra_btn:
              st.session_state["student_data"].loc[st_idx[0], "GR_No"] = new_gr
              st.session_state["student_data"].loc[st_idx[0], "House"] = (
                  new_house
              )
              st.session_state["student_data"].loc[st_idx[0], "Area"] = new_area
              st.session_state["student_data"].loc[
                  st_idx[0], "Skill_Course"
              ] = new_skill
              st.session_state["student_data"].loc[st_idx[0], "Discipline"] = (
                  new_disc
              )
              st.session_state["student_data"].loc[
                  st_idx[0], "Working_Days"
              ] = new_work_days
              st.session_state["student_data"].loc[
                  st_idx[0], "Present_Days"
              ] = new_pres_days
              st.session_state["student_data"].loc[st_idx[0], "Attendance"] = (
                  new_att_pct
              )
              st.session_state["student_data"].loc[
                  st_idx[0], "Outstanding_Achievement"
              ] = new_achieve
              st.session_state["student_data"].loc[st_idx[0], "Remarks"] = (
                  new_remarks
              )
              st.session_state["student_data"].loc[
                  st_idx[0], "Co_Scholastic"
              ] = new_co_sch
              st.session_state["student_data"].loc[
                  st_idx[0], "Bagless_Days"
              ] = new_bagless

              sync_df_to_sqlite(st.session_state["student_data"])
              st.success(f"✅ Updated Report Card details for Roll {sel_roll}!")
              st.rerun()

        st.markdown("---")
        st.markdown("##### 📷 Single Student Photo Upload")
        up_p_col1, up_p_col2 = st.columns([2, 2])
        with up_p_col1:
          roll_to_photo = st.selectbox(
              "Select Student Roll No",
              sorted(
                  st.session_state["student_data"]["Roll_No"]
                  .astype(str)
                  .unique()
              ),
              key="photo_roll_sel",
          )
        with up_p_col2:
          stu_photo_file = st.file_uploader(
              "Upload Single Photo",
              type=["png", "jpg", "jpeg"],
              key="stu_photo_up",
          )
          if st.button("🖼️ Save Photo"):
            if stu_photo_file and roll_to_photo:
              Image.open(stu_photo_file).save(
                  f"photos/students/{roll_to_photo}.png"
              )
              st.success(
                  f"✅ Photo uploaded for Roll No: {roll_to_photo}!"
              )

        st.markdown("---")
        st.markdown("##### 📦 Bulk Student Photo Upload (ZIP Extractor)")
        zip_photo_file = st.file_uploader(
            "Upload Photos Archive (.zip)", type=["zip"], key="zip_photos_up"
        )
        if st.button("🚀 Process & Extract ZIP Photos"):
          if zip_photo_file:
            try:
              extracted_count = 0
              with zipfile.ZipFile(zip_photo_file, "r") as z:
                for file_info in z.infolist():
                  if not file_info.is_dir():
                    ext = os.path.splitext(file_info.filename)[1].lower()
                    if ext in [".png", ".jpg", ".jpeg"]:
                      raw_name = os.path.basename(file_info.filename)
                      roll_stem = os.path.splitext(raw_name)[0]
                      clean_roll = format_clean_number(roll_stem)
                      if clean_roll:
                        out_path = f"photos/students/{clean_roll}.png"
                        img_data = z.read(file_info.filename)
                        img = Image.open(io.BytesIO(img_data))
                        img.save(out_path)
                        extracted_count += 1
              st.success(
                  "✅ Successfully extracted and linked"
                  f" {extracted_count} student photo(s)!"
              )
            except Exception as ex:
              st.error(f"❌ Error extracting ZIP: {ex}")

        st.markdown("---")
        st.markdown("##### 📝 Realtime Data & Rank Modifier")

        if st.button("🔄 Auto-Recalculate Class Ranks"):
          df_mod = st.session_state["student_data"].copy()
          df_mod["Total_Marks"] = df_mod[ALL_SUBJECTS].sum(
              axis=1, skipna=True
          )
          df_mod["Percentage"] = (
              (df_mod["Total_Marks"] / df_mod["Max_Marks"]) * 100
          ).round(2)
          df_mod["Class_Rank"] = (
              df_mod.groupby(["Class", "Exam_Type"])["Total_Marks"]
              .rank(ascending=False, method="min")
              .fillna(0)
              .astype(int)
          )
          st.session_state["student_data"] = df_mod
          sync_df_to_sqlite(df_mod)
          st.success("✅ Class ranks recalculated & saved!")
          st.rerun()

        edited_df = st.data_editor(
            st.session_state["student_data"],
            num_rows="dynamic",
            use_container_width=True,
            key="db_realtime_editor",
        )

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
                  edited_df[sub] = pd.to_numeric(
                      edited_df[sub], errors="coerce"
                  )
              edited_df["Total_Marks"] = edited_df[ALL_SUBJECTS].sum(
                  axis=1, skipna=True
              )
              edited_df["Percentage"] = (
                  (edited_df["Total_Marks"] / edited_df["Max_Marks"]) * 100
              ).round(2)

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

    with st.expander(
        "📢 4. DIGITAL NOTICE BOARD & 🏆 BOARD TOPPERS MANAGEMENT",
        expanded=False,
    ):
      st.subheader("📢 Digital Notice Board")
      current_notices = load_notices()
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

      st.markdown("---")
      # Requirement 8: Board Toppers with Rank management
      st.subheader("🏆 CBSE Board Exam Toppers Hall of Fame (With Rank)")
      toppers = load_board_toppers()

      with st.form("add_topper_form"):
        tc1, tc2, tc3 = st.columns(3)
        with tc1:
          t_name = st.text_input("Student Name")
          t_class = st.selectbox("Class", ["10", "12"])
        with tc2:
          t_pct = st.text_input("Percentage / Score (e.g. 96.4%)")
          t_rank = st.text_input("Rank (e.g. 1, 2, 3)")
        with tc3:
          t_year = st.text_input("Academic Year", value="2024-25")
          t_photo = st.file_uploader(
              "Topper Photo", type=["png", "jpg", "jpeg"]
          )

        add_topper_btn = st.form_submit_button("➕ Add Board Topper")
        if add_topper_btn and t_name and t_pct:
          photo_path = ""
          if t_photo:
            photo_path = f"photos/board/{t_name.replace(' ', '_')}_{t_class}.png"
            Image.open(t_photo).save(photo_path)
          toppers.append({
              "name": t_name,
              "class": t_class,
              "percentage": t_pct,
              "rank": t_rank if t_rank else "1",
              "year": t_year,
              "photo": photo_path,
          })
          save_board_toppers(toppers)
          st.success(f"✅ Board topper {t_name} added successfully!")
          st.rerun()

      if toppers:
        st.write("**Existing Board Toppers:**")
        for idx, top in enumerate(toppers):
          bc1, bc2 = st.columns([5, 1])
          bc1.write(
              f"🎓 **{top['name']}** | Class {top['class']} | Rank"
              f" #{top.get('rank', '1')} | {top['percentage']} ({top['year']})"
          )
          if bc2.button("🗑️ Remove", key=f"del_top_{idx}"):
            toppers.pop(idx)
            save_board_toppers(toppers)
            st.rerun()

    with st.expander(
        "✒️ 5. DIGITAL SEAL & SIGNATURES MANAGEMENT", expanded=False
    ):
      s_col1, s_col2, s_col3, s_col4 = st.columns(4)
      with s_col1:
        st.subheader("Parent Signature")
        parent_sign_file = st.file_uploader(
            "Upload Parent Sign", type=["png", "jpg", "jpeg"], key="par_up"
        )
        if st.button("Save Parent Sign") and parent_sign_file:
          Image.open(parent_sign_file).save(PARENT_SIGN_PATH)
          st.success("✅ Parent Signature updated!")

      with s_col2:
        st.subheader("Teacher Sign")
        teacher_sign_file = st.file_uploader(
            "Upload Teacher Sign", type=["png", "jpg", "jpeg"], key="tch_up"
        )
        if st.button("Save Teacher Sign") and teacher_sign_file:
          Image.open(teacher_sign_file).save(TEACHER_SIGN_PATH)
          st.success("✅ Teacher Signature updated!")

      with s_col3:
        st.subheader("Student Sign")
        student_sign_file = st.file_uploader(
            "Upload Student Sign", type=["png", "jpg", "jpeg"], key="std_up"
        )
        if st.button("Save Student Sign") and student_sign_file:
          Image.open(student_sign_file).save(STUDENT_SIGN_PATH)
          st.success("✅ Student Signature updated!")

      with s_col4:
        st.subheader("Principal Sign & Seal")
        sign_file = st.file_uploader(
            "Upload Principal Sign", type=["png", "jpg", "jpeg"], key="prn_up"
        )
        if st.button("Save Principal Sign") and sign_file:
          Image.open(sign_file).save(SIGN_PATH)
          st.success("✅ Principal Signature updated!")

        seal_file = st.file_uploader(
            "Upload Official Seal", type=["png", "jpg", "jpeg"], key="seal_up"
        )
        if st.button("Save Official Seal") and seal_file:
          Image.open(seal_file).save(SEAL_PATH)
          st.success("✅ Official Seal updated!")

    with st.expander("🖼️ 6. SCHOOL GALLERY MANAGEMENT", expanded=False):
      gallery_upload = st.file_uploader(
          "Upload Image to Gallery",
          type=["png", "jpg", "jpeg"],
          key="gal_upload",
      )
      if st.button("➕ Add Image to Gallery") and gallery_upload:
        gal_path = os.path.join("photos/gallery", gallery_upload.name)
        Image.open(gallery_upload).save(gal_path)
        st.success("✅ Gallery image added successfully!")
        st.rerun()

      gallery_files = [
          f
          for f in os.listdir("photos/gallery")
          if f.lower().endswith((".png", ".jpg", ".jpeg"))
      ]
      if gallery_files:
        st.write("**Current Gallery Photos:**")
        cols = st.columns(4)
        for idx, g_file in enumerate(gallery_files):
          with cols[idx % 4]:
            g_path = os.path.join("photos/gallery", g_file)
            with st.container():
              st.image(g_path, use_container_width=True)
              if st.button("🗑️ Delete", key=f"del_gal_{idx}"):
                os.remove(g_path)
                st.success("✅ Deleted successfully!")
                st.rerun()

    # Requirement 9: Parent Messaging System (House, Class, Junior/Senior, All, WhatsApp & Direct SMS)
    with st.expander(
        "📨 7. PARENT MESSAGING SYSTEM (WHATSAPP & SMS)", expanded=True
    ):
      st.subheader("📢 Send Bulk / Targeted Messages to Parents")
      if st.session_state["student_data"] is not None:
        df_msg = st.session_state["student_data"]

        msg_col1, msg_col2 = st.columns(2)
        with msg_col1:
          target_type = st.selectbox(
              "Select Target Audience",
              ["All Students", "By House", "By Class", "Junior / Senior"],
          )
          selected_recipient_filter = ""
          if target_type == "By House":
            selected_recipient_filter = st.selectbox(
                "Select House", ["Aravali", "Nilgiri", "Shivalik", "Udaigiri"]
            )
          elif target_type == "By Class":
            selected_recipient_filter = st.selectbox(
                "Select Class", sorted(df_msg["Class"].astype(str).unique())
            )
          elif target_type == "Junior / Senior":
            selected_recipient_filter = st.selectbox(
                "Select Group", ["Junior (Classes 6-8)", "Senior (Classes 9-12)"]
            )

        with msg_col2:
          delivery_mode = st.radio(
              "Delivery Mode",
              [
                  "WhatsApp Direct Links (Bulk)",
                  "Direct SMS / Phone Broadcast Simulation",
              ],
          )

        message_text = st.text_area(
            "Enter Message Text",
            value=(
                "Dear Parent, this is an important notification from PM SHRI"
                " JNV Chhotaudepur. Please check your ward's exam result on"
                " the portal."
            ),
        )

        if st.button("🚀 Prepare & Broadcast Message"):
          # Filter target students
          target_df = df_msg.copy()
          if target_type == "By House" and "House" in target_df.columns:
            target_df = target_df[
                target_df["House"].astype(str).str.strip().str.lower()
                == selected_recipient_filter.strip().lower()
            ]
          elif target_type == "By Class":
            target_df = target_df[
                target_df["Class"].astype(str).str.strip().str.lower()
                == selected_recipient_filter.strip().lower()
            ]
          elif target_type == "Junior / Senior":
            junior_classes = ["6", "7", "8", "6th", "7th", "8th"]
            if "Junior" in selected_recipient_filter:
              target_df = target_df[
                  target_df["Class"].astype(str).isin(junior_classes)
              ]
            else:
              target_df = target_df[
                  ~target_df["Class"].astype(str).isin(junior_classes)
              ]

          # Deduplicate by Mobile No
          target_df = target_df.drop_duplicates(subset=["Mobile_No"])
          st.success(
              f"🎯 Target Audience Filtered: **{len(target_df)}** parent(s)"
              " matched."
          )

          if len(target_df) > 0:
            if "WhatsApp" in delivery_mode:
              st.write(
                  "📱 **Click below to open WhatsApp chats for each"
                  " parent:**"
              )
              for _, r in target_df.iterrows():
                mob = clean_mobile_for_wa(r.get("Mobile_No", ""))
                if len(mob) >= 10:
                  encoded_msg = urllib.parse.quote(
                      f"Hello {r['Student_Name']}'s Parent,\n\n{message_text}"
                  )
                  wa_link = f"https://wa.me/{mob}?text={encoded_msg}"
                  st.markdown(
                      f"💬 Send to **{r['Student_Name']}** (Class"
                      f" {r['Class']} - {r.get('Mobile_No', '')}):"
                      f" [Open WhatsApp Chat]({wa_link})"
                  )
            else:
              st.success(
                  "✅ Broadcast simulated successfully via Phone SMS gateway"
                  f" to {len(target_df)} recipients!"
              )
              for _, r in target_df.iterrows():
                st.write(
                    f"📤 [SMS Sent] -> {r.get('Mobile_No', '')} | Student:"
                    f" {r['Student_Name']}"
                )