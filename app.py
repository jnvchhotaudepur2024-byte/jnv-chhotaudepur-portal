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

# Helper Function: Clean Alphanumeric String
def clean_val(val):
    if pd.isna(val) or val is None:
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(val)).lower().strip()

# Helper Function: Rank-Based Feedback / Status Comments
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

# Cached Background Image CSS Injection
@st.cache_data
def get_base64_image(image_path):
    import base64
    if os.path.exists(image_path):
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode()
    return None

BG_PATH = "photos/system/background.png"
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
    meta_cols = ['Class', 'Roll_No', 'Student_Name', 'Father_Name', 'DOB', 'Aadhaar_No', 'Exam_Type', 'Max_Marks', 'Class_Teacher']
    
    if 'Max_Marks' not in df.columns:
        df['Max_Marks'] = df['Exam_Type'].apply(lambda x: 150 if 'PWT' in str(x).upper() else 600)
    
    subject_cols = [col for col in df.columns if col not in meta_cols]
    
    df['Total_Marks'] = df[subject_cols].sum(axis=1)
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

# Sidebar Navigation
st.sidebar.title("📌 Navigation")
menu = st.sidebar.radio("Select Portal / Page:", ["👨‍🎓 Parent Portal", "🖼️ School Gallery", "🏆 Board Exam Results", "⚙️ Admin Portal"])

# Top Header Layout
head_col1, head_col2 = st.columns([4, 1])
with head_col1:
    st.title("🏫 Jawahar Navodaya Vidyalaya, Chhotaudepur")
    st.subheader("📊 Student Performance & Result Portal")
with head_col2:
    LOGO_PATH = "photos/system/logo.png"
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=110)

st.markdown("---")


# ==============================================================================
# 👨‍🎓 PARENT PORTAL
# ==============================================================================
if menu == "👨‍🎓 Parent Portal":
    
    # 1. Continuous Right-to-Left Ticker: LAST UPLOADED EXAM ONLY
    if st.session_state["student_data"] is not None:
        df_data = st.session_state["student_data"]
        
        if not df_data.empty and 'Exam_Type' in df_data.columns:
            # Identify the latest uploaded exam type dynamically
            latest_exam = df_data['Exam_Type'].dropna().iloc[-1]
            latest_df = df_data[df_data['Exam_Type'] == latest_exam].copy()
            
            if not latest_df.empty:
                # Overall School Topper for Latest Exam
                school_topper = latest_df.sort_values(by='Percentage', ascending=False).iloc[0]
                ticker_items = [
                    f"🏆 <b>OVERALL SCHOOL TOPPER ({latest_exam}):</b> {school_topper['Student_Name']} (Class {school_topper['Class']}) - {school_topper['Percentage']:.2f}%"
                ]
                
                # Top 3 Toppers for EVERY Class in Latest Exam
                classes = sorted(latest_df['Class'].astype(str).unique())
                for cls in classes:
                    cls_toppers = latest_df[latest_df['Class'].astype(str) == cls].sort_values(by='Percentage', ascending=False).head(3)
                    top_list = [f"{idx+1}. {r['Student_Name']} ({r['Percentage']:.1f}%)" for idx, (_, r) in enumerate(cls_toppers.iterrows())]
                    ticker_items.append(f"🥇 <b>Class {cls} ({latest_exam}) Top 3:</b> {' | '.join(top_list)}")
                    
                full_ticker_text = " &nbsp;&nbsp;&nbsp; ✦ &nbsp;&nbsp;&nbsp; ".join(ticker_items)
                
                st.markdown(
                    f"""
                    <div style="background-color: #FFF9C4; border-left: 5px solid #FBC02D; padding: 7px 10px; border-radius: 4px; color: #000; font-size: 15px;">
                        <marquee direction="left" scrollamount="6" behavior="scroll">{full_ticker_text}</marquee>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                st.write("")

    # 2. Class-wise Top 3 Hall of Fame Expander: LAST UPLOADED EXAM ONLY
    if st.session_state["student_data"] is not None:
        df_data = st.session_state["student_data"]
        if not df_data.empty and 'Exam_Type' in df_data.columns:
            latest_exam = df_data['Exam_Type'].dropna().iloc[-1]
            latest_df = df_data[df_data['Exam_Type'] == latest_exam].copy()
            
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

    st.header("🔎 Check Student Result")
    
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
                
                st.markdown("---")
                
                st.subheader("📈 Overall Performance Summary")
                avg_pct = filtered_df['Percentage'].mean()
                total_obtained = filtered_df['Total_Marks'].sum()
                total_max = filtered_df['Max_Marks'].sum()
                
                k1, k2, k3 = st.columns(3)
                k1.metric("Overall Average Score", f"{avg_pct:.2f}%")
                k2.metric("Total Marks Obtained", f"{total_obtained} / {total_max}")
                k3.metric("Overall Status", "PASS / EXCELLENT" if avg_pct >= 60 else "NEEDS IMPROVEMENT")
                
                st.markdown("---")

                # Cumulative Exam Progression Table
                st.subheader("📊 Cumulative Exam-Wise Performance Trend")
                cum_summary = []
                for _, r in filtered_df.iterrows():
                    cum_summary.append({
                        "Exam Name": r['Exam_Type'],
                        "Marks Obtained": f"{r['Total_Marks']} / {r['Max_Marks']}",
                        "Percentage (%)": f"{r['Percentage']:.2f}%",
                        "Class Rank": f"#{r['Class_Rank']}",
                        "Status / Feedback Comment": get_rank_comment(r['Class_Rank'])
                    })
                st.dataframe(pd.DataFrame(cum_summary), hide_index=True, use_container_width=True)

                st.markdown("---")
                
                # Detailed Subject Scorecard
                st.subheader("📝 Exam-Wise Detailed Subject Scorecard")
                meta_cols = ['Class', 'Roll_No', 'Student_Name', 'Father_Name', 'DOB', 'Aadhaar_No', 'Exam_Type', 'Max_Marks', 'Class_Teacher', 'Total_Marks', 'Percentage', 'Class_Rank']
                subject_cols = [col for col in df.columns if col not in meta_cols]
                
                for index, row in filtered_df.iterrows():
                    with st.expander(f"📌 **{row['Exam_Type']}** | Score: {row['Total_Marks']}/{row['Max_Marks']} ({row['Percentage']:.2f}%) | Rank: #{row['Class_Rank']}", expanded=True):
                        st.info(f"💡 **Teacher's Feedback Comment:** {get_rank_comment(row['Class_Rank'])}")
                        
                        subject_rows = []
                        for s_no, sub_name in enumerate(subject_cols, start=1):
                            subject_rows.append({
                                'S.No.': s_no,
                                'Subject Name': sub_name,
                                'Marks Obtained': row[sub_name]
                            })
                        
                        m_df = pd.DataFrame(subject_rows)
                        st.dataframe(m_df, hide_index=True, use_container_width=True)


# ==============================================================================
# 🖼️ SCHOOL GALLERY
# ==============================================================================
elif menu == "🖼️ School Gallery":
    st.header("🏫 Jawahar Navodaya Vidyalaya - Walking School Gallery")
    st.markdown("---")
    
    gallery_files = [f for f in os.listdir("photos/gallery") if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
    
    if len(gallery_files) == 0:
        st.info("ℹ️ Gallery abhi khali hai. Admin Portal se nayi photos upload karein.")
    else:
        import base64
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
# 🏆 BOARD EXAM RESULTS
# ==============================================================================
elif menu == "🏆 Board Exam Results":
    st.header("🎓 CBSE Board Exam Hall of Fame (Top 3 Toppers)")
    st.markdown("---")
    
    toppers_data = load_board_toppers()
    if not toppers_data:
        st.info("ℹ️ Board Exam Toppers details abhi upload nahi hue hain. Admin Portal se add karein.")
    else:
        b_col1, b_col2 = st.columns(2)
        
        # Class 12 Top 3
        with b_col1:
            st.subheader("🥇 Class 12 CBSE Board - Top 3 Toppers")
            c12_list = [t for t in toppers_data if "12" in t["class"]][:3]
            if not c12_list:
                st.write("Class 12 toppers data available nahi hai.")
            for rank_idx, t in enumerate(c12_list, start=1):
                tc1, tc2 = st.columns([1, 3])
                with tc1:
                    if os.path.exists(t["photo"]):
                        st.image(t["photo"], width=110)
                    else:
                        st.info("📷 No Photo")
                with tc2:
                    st.write(f"🏆 **Rank #{rank_idx}: {t['name']}**")
                    st.write(f"🎯 Score: **{t['percentage']}** ({t['year']})")
                st.write("---")

        # Class 10 Top 3
        with b_col2:
            st.subheader("🥇 Class 10 CBSE Board - Top 3 Toppers")
            c10_list = [t for t in toppers_data if "10" in t["class"]][:3]
            if not c10_list:
                st.write("Class 10 toppers data available nahi hai.")
            for rank_idx, t in enumerate(c10_list, start=1):
                tc1, tc2 = st.columns([1, 3])
                with tc1:
                    if os.path.exists(t["photo"]):
                        st.image(t["photo"], width=110)
                    else:
                        st.info("📷 No Photo")
                with tc2:
                    st.write(f"🏆 **Rank #{rank_idx}: {t['name']}**")
                    st.write(f"🎯 Score: **{t['percentage']}** ({t['year']})")
                st.write("---")


# ==============================================================================
# ⚙️ ADMIN PORTAL
# ==============================================================================
elif menu == "⚙️ Admin Portal":
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
        
        # 1. System Branding
        st.subheader("🎨 School Branding & Background Settings")
        col_sys1, col_sys2 = st.columns(2)
        with col_sys1:
            up_logo = st.file_uploader("Upload School Logo (Top-Right)", type=["png", "jpg", "jpeg"], key="up_logo")
            btn_c1, btn_c2 = st.columns(2)
            with btn_c1:
                if st.button("Save School Logo"):
                    if up_logo:
                        img = Image.open(up_logo)
                        img.save(LOGO_PATH)
                        st.cache_data.clear()
                        st.success("✅ School Logo Saved!")
                        st.rerun()
            with btn_c2:
                if st.button("❌ Remove Logo"):
                    if os.path.exists(LOGO_PATH):
                        os.remove(LOGO_PATH)
                        st.cache_data.clear()
                        st.success("✅ Logo Removed!")
                        st.rerun()

        with col_sys2:
            up_bg = st.file_uploader("Upload Site Background Image", type=["png", "jpg", "jpeg"], key="up_bg")
            bg_c1, bg_c2 = st.columns(2)
            with bg_c1:
                if st.button("Save Background Image"):
                    if up_bg:
                        img = Image.open(up_bg)
                        img.save(BG_PATH)
                        st.cache_data.clear()
                        st.success("✅ Background Image Applied!")
                        st.rerun()
            with bg_c2:
                if st.button("❌ Remove Background"):
                    if os.path.exists(BG_PATH):
                        os.remove(BG_PATH)
                        st.cache_data.clear()
                        st.success("✅ Background Removed!")
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
            g_btn1, g_btn2 = st.columns(2)
            with g_btn1:
                if st.button("Upload to Gallery"):
                    if gal_photo:
                        time_stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        img = Image.open(gal_photo)
                        img.save(f"photos/gallery/{time_stamp}.png")
                        st.success("✅ Gallery Image Uploaded Successfully!")
                        st.rerun()
            with g_btn2:
                if st.button("🗑️ Clear All Gallery Photos"):
                    for file in os.listdir("photos/gallery"):
                        file_path = os.path.join("photos/gallery", file)
                        if os.path.isfile(file_path):
                            os.remove(file_path)
                    st.success("✅ Sabhi Gallery Photos Delete Ho Gayi Hain!")
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
# DEVELOPER CREDIT & COPYRIGHT FOOTER
# ==============================================================================
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; padding: 12px 0 5px 0;'>
        <p style='font-family: "Trebuchet MS", "Segoe UI", sans-serif; font-size: 22px; font-weight: bold; color: #1E88E5; margin-bottom: 3px;'>
            Developer: ANIL CHAUDHARY
        </p>
        <p style='font-family: Arial, sans-serif; font-size: 12px; font-weight: 600; color: #444444; letter-spacing: 0.8px;'>
            © COPYRIGHTS JNV CHHOTAUDEPUR ALL RIGHTS RESERVED
        </p>
    </div>
    """,
    unsafe_allow_html=True
)