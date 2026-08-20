import streamlit as st
import pandas as pd
import os
import shutil
import re
import json
from datetime import datetime
from PIL import Image

st.set_page_config(page_title="JNV Chhotaudepur - Result Portal", page_icon="🎓", layout="wide")

# Directory Setup
os.makedirs("photos/students", exist_ok=True)
os.makedirs("photos/gallery", exist_ok=True)
os.makedirs("photos/board", exist_ok=True)
os.makedirs("photos/system", exist_ok=True)
os.makedirs("backups", exist_ok=True)

# Helper Function: Clean Alphanumeric String for Flexible Matching
def clean_val(val):
    if pd.isna(val) or val is None:
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(val)).lower().strip()

# Background Image CSS Injection
BG_PATH = "photos/system/background.png"
if os.path.exists(BG_PATH):
    import base64
    with open(BG_PATH, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode()
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: linear-gradient(rgba(255, 255, 255, 0.88), rgba(255, 255, 255, 0.88)), url("data:image/png;base64,{encoded_string}");
            background-size: cover;
            background-attachment: fixed;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

# Continuous Visit Counter
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

# Log Parent Search
def log_parent_search(roll_no, student_name, selected_class):
    log_file = "result_logs.csv"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_data = pd.DataFrame([{"Timestamp": timestamp, "Roll_No": str(roll_no), "Student_Name": student_name, "Class": selected_class}])
    if os.path.exists(log_file):
        new_data.to_csv(log_file, mode='a', header=False, index=False)
    else:
        new_data.to_csv(log_file, mode='w', header=True, index=False)

# Excel Data Processor
def process_data(df):
    meta_cols = ['Class', 'Roll_No', 'Student_Name', 'Father_Name', 'DOB', 'Aadhaar_No', 'Exam_Type', 'Class_Teacher']
    subject_cols = [col for col in df.columns if col not in meta_cols]
    
    df['Total_Marks'] = df[subject_cols].sum(axis=1)
    df['Max_Marks'] = df['Exam_Type'].apply(lambda x: 150 if 'PWT' in str(x).upper() else 600)
    df['Percentage'] = (df['Total_Marks'] / df['Max_Marks']) * 100
    df['Class_Rank'] = df.groupby(['Class', 'Exam_Type'])['Total_Marks'].rank(ascending=False, method='min').astype(int)
    return df

EXCEL_FILE_PATH = "JNV_Student_Marks.xlsx"

if "student_data" not in st.session_state or st.session_state["student_data"] is None:
    if os.path.exists(EXCEL_FILE_PATH):
        try:
            default_df = pd.read_excel(EXCEL_FILE_PATH)
            st.session_state["student_data"] = process_data(default_df)
        except Exception as e:
            st.error(f"Excel read error: {e}")
    else:
        st.session_state["student_data"] = None

if "admin_logged_in" not in st.session_state:
    st.session_state["admin_logged_in"] = False

# Board Toppers JSON Storage
BOARD_TOPPERS_FILE = "board_toppers.json"
def load_board_toppers():
    if os.path.exists(BOARD_TOPPERS_FILE):
        with open(BOARD_TOPPERS_FILE, "r") as f:
            return json.load(f)
    return []

# Header Layout with Top-Right School Logo
head_col1, head_col2 = st.columns([4, 1])
with head_col1:
    st.title("🏫 Jawahar Navodaya Vidyalaya, Chhotaudepur")
    st.subheader("📊 Student Performance & Result Portal")
with head_col2:
    LOGO_PATH = "photos/system/logo.png"
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=120)

st.markdown("---")

# Navigation
menu = st.sidebar.radio("Navigation", ["👨‍🎓 Parent Portal", "🖼️ School Gallery", "🏆 Board Exam Results", "⚙️ Admin Portal"])


# ==============================================================================
# ⚙️ ADMIN PORTAL
# ==============================================================================
if menu == "⚙️ Admin Portal":
    st.header("🔒 Admin Dashboard")
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
        if st.button("🚪 Logout"):
            st.session_state["admin_logged_in"] = False
            st.rerun()
            
        st.markdown("---")
        
        # 1. System Media Management (Logo & Background)
        st.subheader("🎨 School Branding & Site Background")
        col_sys1, col_sys2 = st.columns(2)
        with col_sys1:
            up_logo = st.file_uploader("Upload School Logo (Top-Right)", type=["png", "jpg", "jpeg"], key="up_logo")
            if st.button("Save School Logo"):
                if up_logo:
                    img = Image.open(up_logo)
                    img.save(LOGO_PATH)
                    st.success("✅ School Logo Saved!")
                    st.rerun()
        
        with col_sys2:
            up_bg = st.file_uploader("Upload Site Background Image", type=["png", "jpg", "jpeg"], key="up_bg")
            if st.button("Save Background Image"):
                if up_bg:
                    img = Image.open(up_bg)
                    img.save(BG_PATH)
                    st.success("✅ Background Image Applied!")
                    st.rerun()

        st.markdown("---")
        
        # 2. Upload Student & Gallery Photos
        st.subheader("📸 Media Uploads (Students & Gallery)")
        col_u1, col_u2 = st.columns(2)
        with col_u1:
            st_roll = st.text_input("Enter Student Roll No for Photo")
            st_photo = st.file_uploader("Student Image", type=["jpg", "jpeg", "png"], key="st_photo")
            if st.button("Upload Student Photo"):
                if st_roll and st_photo:
                    img = Image.open(st_photo)
                    img.save(f"photos/students/{st_roll.strip()}.png")
                    st.success(f"✅ Student Photo Saved for Roll No: {st_roll}")
                else:
                    st.warning("Roll No and Image are required.")

        with col_u2:
            gal_photo = st.file_uploader("School Gallery Image", type=["jpg", "jpeg", "png"], key="gal_photo")
            if st.button("Upload to Gallery"):
                if gal_photo:
                    time_stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    img = Image.open(gal_photo)
                    img.save(f"photos/gallery/{time_stamp}.png")
                    st.success("✅ Gallery Image Uploaded!")
                    st.rerun()

        st.markdown("---")
        
        # 3. Board Toppers Management
        st.subheader("🏆 Upload Board Exam Toppers Data")
        with st.form("board_form"):
            b_class = st.selectbox("Select Class", ["Class 10", "Class 12"])
            b_name = st.text_input("Topper Student Name")
            b_percent = st.text_input("Percentage (e.g. 98.4%)")
            b_year = st.text_input("Passing Year", value="2025-26")
            b_photo = st.file_uploader("Topper Student Photo", type=["jpg", "jpeg", "png"])
            b_submit = st.form_submit_button("Add Board Topper")
            
            if b_submit and b_name and b_percent and b_photo:
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
                st.success("✅ Board Topper Added Successfully!")
                st.rerun()

        st.markdown("---")
        
        # 4. Result Check Tracker
        st.subheader("📋 Parents Search Logs Tracker")
        if os.path.exists("result_logs.csv"):
            logs_df = pd.read_csv("result_logs.csv")
            st.dataframe(logs_df, use_container_width=True)

        st.markdown("---")
        
        # 5. Excel Upload with Auto-Backup
        st.subheader("📤 Excel Data Upload (With Auto-Backup)")
        uploaded_file = st.file_uploader("Upload Excel Sheet (.xlsx)", type=["xlsx", "xls"])
        
        if uploaded_file is not None:
            if os.path.exists(EXCEL_FILE_PATH):
                backup_name = f"backups/Backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                shutil.copy(EXCEL_FILE_PATH, backup_name)
                st.info(f"🛡️ **Auto-Backup Saved:** `{backup_name}`")
            
            with open(EXCEL_FILE_PATH, "wb") as f:
                f.write(uploaded_file.getbuffer())
                
            df = pd.read_excel(EXCEL_FILE_PATH)
            st.session_state["student_data"] = process_data(df)
            st.success("✅ Excel Updated & Instantly Live!")
            st.rerun()


# ==============================================================================
# 👨‍🎓 PARENT PORTAL
# ==============================================================================
elif menu == "👨‍🎓 Parent Portal":
    
    # Floating / Scrolling Educational Quotes Ticker
    quotes_text = "💡 <i>'Education is the most powerful weapon which you can use to change the world.'</i> | 📖 <i>'An investment in knowledge pays the best interest.'</i> | 🌟 <i>'Learning gives creativity, creativity leads to thinking, thinking provides knowledge, knowledge makes you great.'</i>"
    st.markdown(
        f"""
        <div style="background-color: #2196F3; padding: 8px; border-radius: 5px; color: #fff; font-size: 15px; margin-bottom: 10px;">
            <marquee scrollamount="5">{quotes_text}</marquee>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Flash News Ticker
    if st.session_state["student_data"] is not None:
        df_data = st.session_state["student_data"]
        student_summary = df_data.groupby(['Class', 'Student_Name', 'Roll_No']).agg(Overall_Percentage=('Percentage', 'mean')).reset_index()
        school_topper = student_summary.sort_values(by='Overall_Percentage', ascending=False).iloc[0]
        
        news_text = f"🏆 <b>SCHOOL TOPPER:</b> {school_topper['Student_Name']} (Class {school_topper['Class']}) - {school_topper['Overall_Percentage']:.2f}% | "
        st.markdown(
            f"""
            <div style="background-color: #ffeb3b; padding: 8px; border-radius: 5px; color: #000; font-size: 16px;">
                <marquee scrollamount="6">{news_text}</marquee>
            </div>
            """,
            unsafe_allow_html=True
        )
        st.write("")

    # Toppers Results With Photo
    if st.session_state["student_data"] is not None:
        with st.expander("🏆 **SCHOOL TOPPERS & CLASS TOPPERS (HALL OF FAME)**", expanded=False):
            df_data = st.session_state["student_data"]
            student_summary = df_data.groupby(['Class', 'Student_Name', 'Roll_No']).agg(Overall_Percentage=('Percentage', 'mean')).reset_index()
            school_topper = student_summary.sort_values(by='Overall_Percentage', ascending=False).iloc[0]
            
            st.markdown("#### 🥇 Overall School Topper")
            t_col1, t_col2 = st.columns([1, 4])
            with t_col1:
                top_photo_path = f"photos/students/{school_topper['Roll_No']}.png"
                if os.path.exists(top_photo_path):
                    st.image(top_photo_path, width=120)
                else:
                    st.info("📷 No Photo")
            with t_col2:
                st.write(f"🌟 **Name:** {school_topper['Student_Name']}")
                st.write(f"📌 **Class:** {school_topper['Class']} | **Roll No:** {school_topper['Roll_No']}")
                st.write(f"🎯 **Overall Percentage:** {school_topper['Overall_Percentage']:.2f}%")

    st.markdown("---")
    st.header("🔎 Check Student Result")
    
    if st.session_state["student_data"] is None:
        st.warning("⚠️ Data file not found. Kripya Admin Portal se Data Upload karein.")
    else:
        df = st.session_state["student_data"]
        
        search_method = st.radio("Choose Search Verification Method:", ["Option 1: Roll No & Date of Birth (DOB)", "Option 2: Roll No & Aadhaar Number"], horizontal=True)
        
        with st.form("search_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                selected_class = st.selectbox("Select Class", sorted(df['Class'].astype(str).unique()))
                roll_no = st.text_input("Roll No")
            
            if "Option 1" in search_method:
                with c2:
                    dob_input = st.text_input("Date of Birth (Enter plain numbers or with hyphen)")
                with c3:
                    st.write("")
            else:
                with c2:
                    aadhaar_input = st.text_input("Aadhaar Number (Enter plain numbers or with hyphen/space)")
                with c3:
                    st.write("")
            
            submit_btn = st.form_submit_button("🔍 View Result")

        if submit_btn:
            # Flexible Match Logic (Strips hyphen, spaces, slashes)
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
                
                st.success(f"🎓 Result Found for: **{student_info['Student_Name']}**")
                
                r_col1, r_col2 = st.columns([1, 4])
                with r_col1:
                    photo_file = f"photos/students/{student_info['Roll_No']}.png"
                    if os.path.exists(photo_file):
                        st.image(photo_file, width=130)
                    else:
                        st.info("📷 Student Photo Not Uploaded")
                
                with r_col2:
                    st.write(f"**Student Name:** {student_info['Student_Name']} | **Roll No:** {student_info['Roll_No']}")
                    st.write(f"**Class:** {student_info['Class']} | **Class Teacher:** {student_info['Class_Teacher']}")
                    st.write(f"**Father's Name:** {student_info['Father_Name']}")
                
                st.markdown("---")
                
                st.subheader("📈 Overall Performance Summary")
                avg_pct = filtered_df['Percentage'].mean()
                total_obtained = filtered_df['Total_Marks'].sum()
                total_max = filtered_df['Max_Marks'].sum()
                
                k1, k2, k3 = st.columns(3)
                k1.metric("Overall Percentage", f"{avg_pct:.2f}%")
                k2.metric("Total Marks Obtained", f"{total_obtained} / {total_max}")
                k3.metric("Status", "PASS / EXCELLENT" if avg_pct >= 60 else "NEEDS IMPROVEMENT")
                
                st.markdown("---")
                
                st.subheader("📝 Exam-Wise Detailed Scorecard")
                meta_cols = ['Class', 'Roll_No', 'Student_Name', 'Father_Name', 'DOB', 'Aadhaar_No', 'Exam_Type', 'Class_Teacher', 'Total_Marks', 'Max_Marks', 'Percentage', 'Class_Rank']
                subject_cols = [col for col in df.columns if col not in meta_cols]
                
                for index, row in filtered_df.iterrows():
                    with st.expander(f"📌 **{row['Exam_Type']}** | Score: {row['Total_Marks']}/{row['Max_Marks']} ({row['Percentage']:.2f}%) | Rank: #{row['Class_Rank']}", expanded=True):
                        m_df = pd.DataFrame({
                            'Subject': subject_cols,
                            'Marks Obtained': [row[sub] for sub in subject_cols]
                        })
                        st.dataframe(m_df.T, use_container_width=True)


# ==============================================================================
# 🖼️ SCHOOL GALLERY (RIGHT TO LEFT MOVING SCROLL)
# ==============================================================================
elif menu == "🖼️ School Gallery":
    st.header("🏫 Jawahar Navodaya Vidyalaya - Walking School Gallery")
    st.markdown("---")
    
    gallery_files = os.listdir("photos/gallery")
    if len(gallery_files) == 0:
        st.info("Gallery me abhi koi photo upload nahi hui hai. Admin Portal se Upload karein.")
    else:
        # Build HTML Marquee for Walking Images Right-to-Left
        import base64
        images_html = ""
        for img_name in gallery_files:
            img_path = os.path.join("photos/gallery", img_name)
            with open(img_path, "rb") as f:
                enc = base64.b64encode(f.read()).decode()
            images_html += f'<img src="data:image/png;base64,{enc}" style="height: 220px; margin-right: 20px; border-radius: 10px; border: 3px solid #FF9800; display: inline-block;">'
        
        st.markdown(
            f"""
            <div style="width: 100%; overflow: hidden; background-color: rgba(0,0,0,0.05); padding: 15px; border-radius: 10px;">
                <marquee direction="left" scrollamount="8" behavior="scroll">
                    {images_html}
                </marquee>
            </div>
            """,
            unsafe_allow_html=True
        )


# ==============================================================================
# 🏆 BOARD EXAM RESULTS
# ==============================================================================
elif menu == "🏆 Board Exam Results":
    st.header("🎓 CBSE Board Exam Hall of Fame (Class 10 & 12 Toppers)")
    st.markdown("---")
    
    toppers_data = load_board_toppers()
    if not toppers_data:
        st.info("Board Exam Toppers details abhi upload nahi hue hain.")
    else:
        b_col1, b_col2 = st.columns(2)
        
        with b_col1:
            st.subheader("🥇 Class 12 CBSE Board Toppers")
            c12_list = [t for t in toppers_data if "12" in t["class"]]
            for t in c12_list:
                tc1, tc2 = st.columns([1, 3])
                with tc1:
                    if os.path.exists(t["photo"]):
                        st.image(t["photo"], width=100)
                with tc2:
                    st.write(f"🌟 **{t['name']}**")
                    st.write(f"🎯 Score: **{t['percentage']}** ({t['year']})")
                st.write("---")

        with b_col2:
            st.subheader("🥇 Class 10 CBSE Board Toppers")
            c10_list = [t for t in toppers_data if "10" in t["class"]]
            for t in c10_list:
                tc1, tc2 = st.columns([1, 3])
                with tc1:
                    if os.path.exists(t["photo"]):
                        st.image(t["photo"], width=100)
                with tc2:
                    st.write(f"🌟 **{t['name']}**")
                    st.write(f"🎯 Score: **{t['percentage']}** ({t['year']})")
                st.write("---")


# ==============================================================================
# DEVELOPER CREDIT & COPYRIGHT FOOTER
# ==============================================================================
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; padding: 10px;'>
        <p style='font-family: "Brush Script MT", cursive, sans-serif; font-size: 32px; color: #E91E63; margin-bottom: 5px;'>
            Developer: ANIL CHAUDHARY
        </p>
        <p style='font-family: Arial, sans-serif; font-size: 14px; font-weight: bold; color: #333; letter-spacing: 1px;'>
            © COPYRIGHTS JNV CHHOTAUDEPUR ALL RIGHTS RESERVED
        </p>
    </div>
    """,
    unsafe_allow_html=True
)