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
from datetime import datetime
from PIL import Image

# ReportLab Imports for PDF Generation
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="PM SHRI JNV CHHOTAUDEPUR - RESULT PORTAL", page_icon="🎓", layout="wide")

# Custom CSS for Mobile Responsiveness & Layout Tweaks
st.markdown("""
    <style>
    @media (max-width: 768px) {
        .stApp {
            padding: 5px !important;
        }
        h1 {
            font-size: 1.4rem !important;
        }
        h3 {
            font-size: 1.1rem !important;
        }
        .stButton>button {
            width: 100% !important;
        }
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
    </style>
""", unsafe_allow_html=True)

# Directory Setup
os.makedirs("photos/students", exist_ok=True)
os.makedirs("photos/gallery", exist_ok=True)
os.makedirs("photos/board", exist_ok=True)
os.makedirs("photos/system", exist_ok=True)
os.makedirs("backups", exist_ok=True)

# Master Subject List
ALL_SUBJECTS = [
    'Gujarati', 'Hindi', 'English', 'Mathematics', 
    'Science', 'Social_Science', 'Physics', 'Chemistry', 'Biology'
]

# Helper Functions
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

    for col in ALL_SUBJECTS:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df['Total_Marks'] = df[ALL_SUBJECTS].sum(axis=1, skipna=True)

    if 'Max_Marks' not in df.columns or df['Max_Marks'].isnull().all():
        df['Max_Marks'] = df['Exam_Type'].apply(lambda x: 150 if 'PWT' in str(x).upper() else 600)

    df['Max_Marks'] = pd.to_numeric(df['Max_Marks'], errors='coerce').fillna(600)
    df['Percentage'] = (df['Total_Marks'] / df['Max_Marks']) * 100
    df['Percentage'] = df['Percentage'].round(2)
    df['Class_Rank'] = df.groupby(['Class', 'Exam_Type'])['Total_Marks'].rank(ascending=False, method='min').fillna(0).astype(int)
    
    return df

EXCEL_FILE_PATH = "JNV_Student_Marks.xlsx"

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

    info_data = [
        [Paragraph(f"<b>Student Name:</b> {student_info['Student_Name']}", normal_style), Paragraph(f"<b>Roll No:</b> {student_info['Roll_No']}", normal_style)],
        [Paragraph(f"<b>Class:</b> {student_info['Class']}", normal_style), Paragraph(f"<b>Class Teacher:</b> {student_info['Class_Teacher']}", normal_style)],
        [Paragraph(f"<b>Father's Name:</b> {student_info['Father_Name']}", normal_style), Paragraph(f"<b>Date of Birth:</b> {student_info['DOB']}", normal_style)]
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

# Helper Dialog Confirmation Pop-up (Streamlit >= 1.34)
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

# Sidebar Navigation with 3-Lines Symbol
st.sidebar.title("☰ NAVIGATION")
menu = st.sidebar.radio("SELECT PORTAL / PAGE:", ["👨‍🎓 PARENT PORTAL", "🖼️ SCHOOL GALLERY", "🏆 BOARD EXAM RESULTS", "⚙️ ADMIN PORTAL"])

# Top Header Layout (Logo in front of title and subtitle)
head_col1, head_col2 = st.columns([1, 5], vertical_alignment="center")
with head_col1:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=95)
with head_col2:
    st.markdown("<h1 class='main-title'>🏫 PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR</h1>", unsafe_allow_html=True)
    st.markdown("<h3 class='sub-title'>📊 STUDENT PERFORMANCE & RESULT PORTAL</h3>", unsafe_allow_html=True)

st.markdown("---")


# ==============================================================================
# 👨‍🎓 PARENT PORTAL
# ==============================================================================
if menu == "👨‍🎓 PARENT PORTAL":
    
    # 1. Best Educational Quotes Marquee (Moving Right to Left)
    educational_quotes = [
        "🎓 'Education is the most powerful weapon which you can use to change the world.' – Nelson Mandela",
        "🌟 'Live as if you were to die tomorrow. Learn as if you were to live forever.' – Mahatma Gandhi",
        "💡 'The mind is not a vessel to be filled, but a fire to be kindled.' – Plutarch",
        "🚀 'Education is the passport to the future, for tomorrow belongs to those who prepare for it today.' – Malcolm X",
        "📖 'An investment in knowledge pays the best interest.' – Benjamin Franklin"
    ]
    quotes_ticker_text = " &nbsp;&nbsp;&nbsp;&nbsp; ✦ &nbsp;&nbsp;&nbsp;&nbsp; ".join(educational_quotes)

    st.markdown(
        f"""
        <div style="background: linear-gradient(90deg, #1565C0, #1E88E5); border-radius: 6px; padding: 8px 12px; color: #FFFFFF; font-size: 14px; font-weight: 600; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
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
                ticker_items = [
                    f"🏆 <b>OVERALL SCHOOL TOPPER ({latest_exam}):</b> {school_topper['Student_Name']} (Class {school_topper['Class']}) - {school_topper['Percentage']:.2f}%"
                ]
                
                classes = sorted(latest_df['Class'].astype(str).unique())
                for cls in classes:
                    cls_toppers = latest_df[latest_df['Class'].astype(str) == cls].sort_values(by='Percentage', ascending=False).head(3)
                    top_list = [f"{idx+1}. {r['Student_Name']} ({r['Percentage']:.1f}%)" for idx, (_, r) in enumerate(cls_toppers.iterrows())]
                    ticker_items.append(f"🥇 <b>Class {cls} ({latest_exam}) Top 3:</b> {' | '.join(top_list)}")
                    
                full_ticker_text = " &nbsp;&nbsp;&nbsp; ✦ &nbsp;&nbsp;&nbsp; ".join(ticker_items)
                
                st.markdown(
                    f"""
                    <div style="background-color: #FFF9C4; border-left: 5px solid #FBC02D; padding: 7px 10px; border-radius: 4px; color: #000; font-size: 15px; margin-bottom: 10px;">
                        <marquee direction="left" scrollamount="6" behavior="scroll">{full_ticker_text}</marquee>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            with st.expander(f"🏆 **HALL OF FAME - TOP 3 STUDENTS CLASS-WISE ({latest_exam})**", expanded=False):
                classes = sorted(latest_df['Class'].astype(str).unique())
                for cls in classes:
                    st.markdown(f"#### 🏫 **Class {cls} - Top 3 Performers ({latest_exam})**")
                    cls_toppers = latest_df[latest_df['Class'].astype(str) == cls].sort_values(by='Percentage', ascending=False).head(3)
                    
                    t_cols = st.columns(3)
                    badges = ["🥇 1st Rank", "🥈 2nd Rank", "🥉 3rd Rank"]
                    
                    for idx, (_, topper) in enumerate(cls_toppers.iterrows()):
                        if idx < len(t_cols):
                            with t_cols[idx]:
                                top_photo = f"photos/students/{topper['Roll_No']}.png"
                                if os.path.exists(top_photo):
                                    st.image(top_photo, width=100)
                                else:
                                    st.info("📷 No Photo")
                                st.markdown(f"**{badges[idx]}**")
                                st.write(f"👤 **Name:** {topper['Student_Name']}")
                                st.write(f"📌 **Roll No:** {topper['Roll_No']}")
                                st.write(f"🎯 **Score:** {topper['Percentage']:.2f}%")
                    st.markdown("---")

    st.header("🔎 CHECK STUDENT RESULT")
    
    if st.session_state["student_data"] is None:
        st.warning("⚠️ Data file not found. Kripya Admin Portal se Data Upload karein.")
    else:
        df = st.session_state["student_data"]
        
        search_method = st.radio("Choose Search Verification Method:", ["Option 1: Roll No & Date of Birth (DOB)", "Option 2: Roll No & Aadhaar Number"], horizontal=True)
        
        with st.form("search_form"):
            c1, c2 = st.columns(2)
            with c1:
                selected_class = st.selectbox("Select Class", sorted(df['Class'].astype(str).unique()))
                roll_no = st.text_input("Roll No")
            
            if "Option 1" in search_method:
                with c2:
                    dob_input = st.text_input("Date of Birth (Enter plain numbers or with hyphen)")
            else:
                with c2:
                    aadhaar_input = st.text_input("Aadhaar Number (Enter plain numbers or with hyphen/space)")
            
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
                st.error("❌ Invalid Details! Kripya Roll No, DOB ya Aadhaar Number sahi se enter karein.")
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
                        st.info("📷 Photo Not Uploaded")
                
                with r_col2:
                    st.write(f"**Student Name:** {student_info['Student_Name']} | **Roll No:** {student_info['Roll_No']}")
                    st.write(f"**Class:** {student_info['Class']} | **Class Teacher:** {student_info['Class_Teacher']}")
                    st.write(f"**Father's Name:** {student_info['Father_Name']}")
                    
                    st.download_button(
                        label="📥 Download Official PDF Report Card",
                        data=pdf_bytes,
                        file_name=f"Report_Card_{student_info['Roll_No']}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                
                st.markdown("---")
                
                st.subheader("📈 OVERALL PERFORMANCE SUMMARY")
                avg_pct = filtered_df['Percentage'].mean()
                total_obtained = int(filtered_df['Total_Marks'].sum())
                total_max = int(filtered_df['Max_Marks'].sum())
                
                k1, k2, k3 = st.columns(3)
                k1.metric("Overall Average Score", f"{avg_pct:.2f}%")
                k2.metric("Total Marks Obtained", f"{total_obtained} / {total_max}")
                k3.metric("Overall Status", "PASS / EXCELLENT" if avg_pct >= 33 else "NEEDS IMPROVEMENT")
                
                st.markdown("---")

                st.subheader("📊 CUMULATIVE EXAM-WISE PERFORMANCE TREND")
                cum_summary = []
                for _, r in filtered_df.iterrows():
                    cum_summary.append({
                        "Exam Name": r['Exam_Type'],
                        "Marks Obtained": f"{int(r['Total_Marks'])} / {int(r['Max_Marks'])}",
                        "Percentage (%)": f"{r['Percentage']:.2f}%",
                        "Class Rank": f"#{r['Class_Rank']}",
                        "Status / Feedback Comment": get_rank_comment(r['Class_Rank'])
                    })
                st.dataframe(pd.DataFrame(cum_summary), hide_index=True, use_container_width=True)

                st.markdown("---")
                
                st.subheader("📝 EXAM-WISE DETAILED SUBJECT SCORECARD")
                for index, row in filtered_df.iterrows():
                    with st.expander(f"📌 **{row['Exam_Type']}** | Score: {int(row['Total_Marks'])}/{int(row['Max_Marks'])} ({row['Percentage']:.2f}%) | Rank: #{row['Class_Rank']}", expanded=True):
                        st.info(f"💡 **Teacher's Feedback Comment:** {get_rank_comment(row['Class_Rank'])}")
                        
                        subject_rows = []
                        s_no = 1
                        for sub_name in ALL_SUBJECTS:
                            if pd.notna(row[sub_name]):
                                val = row[sub_name]
                                val_formatted = int(val) if float(val).is_integer() else val
                                subject_rows.append({
                                    'S.No.': s_no,
                                    'Subject Name': sub_name,
                                    'Marks Obtained': val_formatted
                                })
                                s_no += 1
                        
                        m_df = pd.DataFrame(subject_rows)
                        st.dataframe(m_df, hide_index=True, use_container_width=True)


# ==============================================================================
# 🖼️ SCHOOL GALLERY
# ==============================================================================
elif menu == "🖼️ SCHOOL GALLERY":
    st.header("🏫 PM SHRI JNV CHHOTAUDEPUR - GALLERY")
    st.markdown("---")
    
    gallery_files = [f for f in os.listdir("photos/gallery") if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
    
    if len(gallery_files) == 0:
        st.info("ℹ️ Gallery abhi khali hai. Admin Portal se nayi photos upload karein.")
    else:
        images_html = ""
        for img_name in gallery_files:
            img_path = os.path.join("photos/gallery", img_name)
            with open(img_path, "rb") as f:
                enc = base64.b64encode(f.read()).decode()
            images_html += f'<img src="data:image/png;base64,{enc}" style="height: 220px; margin-right: 20px; border-radius: 10px; border: 3px solid #1E88E5; display: inline-block;">'
        
        st.markdown(
            f"""
            <div style="width: 100%; overflow: hidden; background-color: rgba(0,0,0,0.03); padding: 15px; border-radius: 10px;">
                <marquee direction="left" scrollamount="8" behavior="scroll">
                    {images_html}
                </marquee>
            </div>
            """,
            unsafe_allow_html=True
        )


# ==============================================================================
# 🏆 BOARD EXAM RESULTS (Moving Hall of Fame Marquee)
# ==============================================================================
elif menu == "🏆 BOARD EXAM RESULTS":
    st.header("🎓 CBSE BOARD EXAM HALL OF FAME")
    st.markdown("---")
    
    toppers_data = load_board_toppers()
    if not toppers_data:
        st.info("ℹ️ Board Exam Toppers details abhi upload nahi hue hain. Admin Portal se add karein.")
    else:
        cards_html = ""
        for t in toppers_data:
            img_src = ""
            if os.path.exists(t.get("photo", "")):
                with open(t["photo"], "rb") as f:
                    enc = base64.b64encode(f.read()).decode()
                    img_src = f"data:image/png;base64,{enc}"
            else:
                img_src = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="

            cards_html += f"""
            <div style="display: inline-block; width: 210px; background: #ffffff; padding: 14px; margin-right: 18px; border-radius: 12px; border: 2px solid #1E88E5; text-align: center; box-shadow: 0 4px 8px rgba(0,0,0,0.08); vertical-align: top;">
                <img src="{img_src}" style="width: 90px; height: 90px; border-radius: 50%; object-fit: cover; border: 3px solid #1565C0; margin-bottom: 8px;">
                <div style="font-weight: 700; font-size: 15px; color: #0D47A1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{t['name']}</div>
                <div style="font-size: 13px; color: #555; margin-top: 2px;">{t['class']} ({t['year']})</div>
                <div style="font-size: 16px; font-weight: bold; color: #2E7D32; margin-top: 6px; background: #E8F5E9; padding: 4px; border-radius: 6px;">🏆 {t['percentage']}</div>
            </div>
            """

        st.markdown(
            f"""
            <div style="width: 100%; overflow: hidden; background: rgba(30, 136, 229, 0.04); padding: 20px 10px; border-radius: 12px; border: 1px solid #BBDEFB;">
                <marquee direction="left" scrollamount="6" behavior="scroll" onmouseover="this.stop();" onmouseout="this.start();">
                    {cards_html}
                </marquee>
            </div>
            """,
            unsafe_allow_html=True
        )
        st.markdown("<br>", unsafe_allow_html=True)


# ==============================================================================
# ⚙️ ADMIN PORTAL (Accordion Expanders & Pop-up Confirmation)
# ==============================================================================
elif menu == "⚙️ ADMIN PORTAL":
    st.header("🔒 ADMIN DASHBOARD")
    st.info(f"👁️ **Total Website Visits Count:** `{total_visits}`")
    
    if not st.session_state["admin_logged_in"]:
        with st.form("login_form"):
            st.subheader("🔐 Admin Login")
            admin_user = st.text_input("Username")
            admin_pass = st.text_input("Password", type="password")
            login_btn = st.form_submit_button("Login")
            
            if login_btn:
                if admin_user == "admin" and admin_pass == "Jnvcu@me2":
                    st.session_state["admin_logged_in"] = True
                    st.success("✅ Login Successful!")
                    st.rerun()
                else:
                    st.error("❌ Invalid Credentials!")
    else:
        st.success("🔓 Logged in as Admin")
        
        with st.expander("🚪 Logout from Admin Portal", expanded=False):
            if st.button("Confirm Logout"):
                st.session_state["admin_logged_in"] = False
                st.rerun()

        st.markdown("---")

        # Admin Actions Callbacks
        def action_save_logo(up_logo):
            img = Image.open(up_logo)
            img.save(LOGO_PATH)
            st.cache_data.clear()

        def action_remove_logo():
            if os.path.exists(LOGO_PATH):
                os.remove(LOGO_PATH)
                st.cache_data.clear()

        def action_save_bg(up_bg):
            img = Image.open(up_bg)
            img.save(BG_PATH)
            st.cache_data.clear()

        def action_remove_bg():
            if os.path.exists(BG_PATH):
                os.remove(BG_PATH)
                st.cache_data.clear()

        def action_upload_st_photo(st_roll, st_photo):
            img = Image.open(st_photo)
            img.save(f"photos/students/{st_roll.strip()}.png")

        def action_upload_gallery(gal_photo):
            time_stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            img = Image.open(gal_photo)
            img.save(f"photos/gallery/{time_stamp}.png")

        def action_clear_gallery():
            for file in os.listdir("photos/gallery"):
                file_path = os.path.join("photos/gallery", file)
                if os.path.isfile(file_path):
                    os.remove(file_path)

        def action_add_board_topper(b_class, b_name, b_percent, b_year, b_photo):
            photo_filename = f"photos/board/{b_class.replace(' ', '_')}_{clean_val(b_name)}.png"
            img = Image.open(b_photo)
            img.save(photo_filename)
            toppers_list = load_board_toppers()
            toppers_list.append({
                "class": b_class,
                "name": b_name,
                "percentage": b_percent,
                "year": b_year,
                "photo": photo_filename
            })
            with open(BOARD_TOPPERS_FILE, "w") as f:
                json.dump(toppers_list, f)

        def action_remove_board_topper(index):
            toppers_list = load_board_toppers()
            if 0 <= index < len(toppers_list):
                item = toppers_list.pop(index)
                if os.path.exists(item.get("photo", "")):
                    try:
                        os.remove(item["photo"])
                    except:
                        pass
                with open(BOARD_TOPPERS_FILE, "w") as f:
                    json.dump(toppers_list, f)

        def action_process_excel(uploaded_file):
            if os.path.exists(EXCEL_FILE_PATH):
                backup_name = f"backups/Backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                shutil.copy(EXCEL_FILE_PATH, backup_name)
            with open(EXCEL_FILE_PATH, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.session_state["student_data"] = process_data_excel(EXCEL_FILE_PATH)
        
        # 1. Parent Notification System Expander
        with st.expander("📲 1. PARENT NOTIFICATION SYSTEM (WHATSAPP)", expanded=False):
            if st.session_state["student_data"] is not None:
                df_notif = st.session_state["student_data"]
                if 'Mobile_No' in df_notif.columns:
                    n_cls = st.selectbox("Select Class to Notify", sorted(df_notif['Class'].astype(str).unique()))
                    filtered_notif = df_notif[df_notif['Class'].astype(str) == str(n_cls)]
                    
                    st.markdown(f"Found **{len(filtered_notif)}** student record(s) for Class **{n_cls}**.")
                    msg_template = st.text_area("Message Content", "Dear Parent, Your student's exam result has been declared. Please visit PM SHRI JNV Chhotaudepur Result Portal to check the details. Thank you!")
                    
                    if st.button("🚀 Prepare One-Click Parent Messages"):
                        st.success("Click on any link below to dispatch the message directly to the Parent:")
                        for idx, row in filtered_notif.iterrows():
                            mob = re.sub(r'[^0-9]', '', str(row['Mobile_No']))
                            if len(mob) >= 10:
                                if not mob.startswith("91") and len(mob) == 10:
                                    mob = "91" + mob
                                encoded_msg = urllib.parse.quote(f"Hello {row['Student_Name']} ({row['Father_Name']}),\n{msg_template}")
                                wa_link = f"https://api.whatsapp.com/send?phone={mob}&text={encoded_msg}"
                                st.markdown(f"👉 **{row['Student_Name']}** (Roll: {row['Roll_No']}) -> [Send WhatsApp Result Alert]({wa_link})")
                            else:
                                st.caption(f"⚠️ {row['Student_Name']}: Invalid or missing mobile number ({row['Mobile_No']})")
                else:
                    st.error("Excel sheet me 'Mobile_No' column nahi mila. Kripya sheet1 me Mobile_No add karein.")

        # 2. System Branding Settings Expander
        with st.expander("🎨 2. SCHOOL BRANDING & BACKGROUND SETTINGS", expanded=False):
            col_sys1, col_sys2 = st.columns(2)
            with col_sys1:
                st.markdown("##### 🏫 School Logo")
                up_logo = st.file_uploader("Upload School Logo", type=["png", "jpg", "jpeg"], key="up_logo")
                btn_c1, btn_c2 = st.columns(2)
                with btn_c1:
                    if st.button("Save Logo"):
                        if up_logo:
                            confirm_action_dialog("Save new school logo?", action_save_logo, up_logo)
                        else:
                            st.warning("Please choose a file first!")
                with btn_c2:
                    if st.button("❌ Remove Logo"):
                        confirm_action_dialog("Remove school logo?", action_remove_logo)

            with col_sys2:
                st.markdown("##### 🖼️ Site Background")
                up_bg = st.file_uploader("Upload Site Background Image", type=["png", "jpg", "jpeg"], key="up_bg")
                bg_c1, bg_c2 = st.columns(2)
                with bg_c1:
                    if st.button("Save Background"):
                        if up_bg:
                            confirm_action_dialog("Apply new background image?", action_save_bg, up_bg)
                        else:
                            st.warning("Please choose a file first!")
                with bg_c2:
                    if st.button("❌ Remove Background"):
                        confirm_action_dialog("Remove current background image?", action_remove_bg)

        # 3. Upload Student & Gallery Photos Expander
        with st.expander("📸 3. MEDIA UPLOADS (STUDENTS & GALLERY)", expanded=False):
            col_u1, col_u2 = st.columns(2)
            with col_u1:
                st.markdown("##### 👤 Student Photo Upload")
                st_roll = st.text_input("Enter Student Roll No")
                st_photo = st.file_uploader("Student Image", type=["jpg", "jpeg", "png"], key="st_photo")
                if st.button("Upload Student Photo"):
                    if st_roll and st_photo:
                        confirm_action_dialog(f"Upload photo for Roll No {st_roll}?", action_upload_st_photo, st_roll, st_photo)
                    else:
                        st.warning("Please provide both Roll No and Photo!")

            with col_u2:
                st.markdown("##### 🖼️ Gallery Uploads")
                gal_photo = st.file_uploader("School Gallery Image", type=["jpg", "jpeg", "png"], key="gal_photo")
                g_btn1, g_btn2 = st.columns(2)
                with g_btn1:
                    if st.button("Upload to Gallery"):
                        if gal_photo:
                            confirm_action_dialog("Upload photo to school gallery?", action_upload_gallery, gal_photo)
                        else:
                            st.warning("Please choose an image!")
                with g_btn2:
                    if st.button("🗑️ Clear All Gallery"):
                        confirm_action_dialog("Delete ALL gallery photos?", action_clear_gallery)

        # 4. Board Toppers Management Expander (Add & Remove with Photo)
        with st.expander("🏆 4. BOARD EXAM TOPPERS MANAGEMENT (ADD / REMOVE)", expanded=False):
            st.subheader("➕ Add New Board Topper")
            b_class = st.selectbox("Select Class", ["Class 10", "Class 12"], key="add_b_class")
            b_name = st.text_input("Topper Student Name", key="add_b_name")
            b_percent = st.text_input("Percentage (e.g. 98.4%)", key="add_b_percent")
            b_year = st.text_input("Passing Year", value="2025-26", key="add_b_year")
            b_photo = st.file_uploader("Topper Student Photo", type=["jpg", "jpeg", "png"], key="add_b_photo")
            
            if st.button("➕ Save Board Topper"):
                if b_name and b_percent and b_photo:
                    confirm_action_dialog(f"Add '{b_name}' to Board Toppers?", action_add_board_topper, b_class, b_name, b_percent, b_year, b_photo)
                else:
                    st.warning("Please fill all details and upload student photo!")

            st.markdown("---")
            st.subheader("🗑️ Existing Board Toppers List (View / Delete)")
            current_toppers = load_board_toppers()
            if not current_toppers:
                st.info("No board toppers added yet.")
            else:
                for idx, item in enumerate(current_toppers):
                    r_col1, r_col2, r_col3 = st.columns([1, 3, 1], vertical_alignment="center")
                    with r_col1:
                        if os.path.exists(item.get("photo", "")):
                            st.image(item["photo"], width=65)
                        else:
                            st.caption("No photo")
                    with r_col2:
                        st.write(f"**{item['name']}** | {item['class']} ({item['year']})")
                        st.write(f"Score: **{item['percentage']}**")
                    with r_col3:
                        if st.button("🗑️ Delete", key=f"del_top_{idx}"):
                            confirm_action_dialog(f"Delete topper '{item['name']}'?", action_remove_board_topper, idx)
                    st.markdown("<hr style='margin: 6px 0;'>", unsafe_allow_html=True)

        # 5. Result Check Tracker Expander
        with st.expander("📋 5. PARENTS SEARCH LOGS TRACKER", expanded=False):
            if os.path.exists("result_logs.csv"):
                logs_df = pd.read_csv("result_logs.csv")
                st.dataframe(logs_df, use_container_width=True)
            else:
                st.info("No logs recorded yet.")

        # 6. Excel Upload Expander
        with st.expander("📤 6. EXCEL DATA UPLOAD (WITH VLOOKUP & AUTO-BACKUP)", expanded=False):
            st.info("💡 Sheet 1 me basic info (`Roll_No`, `Student_Name`, `Father_Name`, `DOB`, `Aadhaar_No`, `Mobile_No`) aur Sheet 2 me Marks (`Gujarati`, `Hindi`, `English`, `Mathematics`, `Science`, `Social_Science`, `Physics`, `Chemistry`, `Biology`) rakhein.")
            uploaded_file = st.file_uploader("Upload Excel Sheet (.xlsx)", type=["xlsx", "xls"])
            
            if st.button("Process & Update Excel Data"):
                if uploaded_file is not None:
                    confirm_action_dialog("Overwrite existing student database with new Excel?", action_process_excel, uploaded_file)
                else:
                    st.warning("Please select an Excel file first!")

# ==============================================================================
# DEVELOPER CREDIT & COPYRIGHT FOOTER
# ==============================================================================
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: #555555; padding: 15px 0px; font-size: 14px;">
        <p style="margin: 0; font-weight: 600;">
            © 2026 PM SHRI JAWAHAR NAVODAYA VIDYALAYA CHHOTAUDEPUR. All Rights Reserved.
        </p>
        <p style="margin: 5px 0 0 0; font-size: 13px; color: #777777;">
            Designed & Developed for School Academic Management & Result Publication
        </p>
    </div>
    """,
    unsafe_allow_html=True
)