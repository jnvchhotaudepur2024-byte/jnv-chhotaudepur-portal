import streamlit as st
import pandas as pd
import os
import shutil
from datetime import datetime
from PIL import Image

st.set_page_config(page_title="JNV Chhotaudepur - Result Portal", page_icon="🎓", layout="wide")

# Directory Creation for Photos & Backups
os.makedirs("photos/students", exist_ok=True)
os.makedirs("photos/gallery", exist_ok=True)
os.makedirs("photos/board", exist_ok=True)
os.makedirs("backups", exist_ok=True)

# 1. Continuous Visit Counter System
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

# 2. Result Check Log System
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

# Navigation
menu = st.sidebar.radio("Navigation", ["👨‍🎓 Parent Portal", "🖼️ School Gallery", "🏆 Board Exam Results", "⚙️ Admin Portal"])


# ==============================================================================
# ⚙️ ADMIN PORTAL
# ==============================================================================
if menu == "⚙️ Admin Portal":
    st.header("🔒 Admin Dashboard")
    
    # 1. Continuous Visit Counter Display
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
        
        # 2. Track Parents Who Checked Results
        st.subheader("📋 Parents Search Logs (Result Status Tracker)")
        if os.path.exists("result_logs.csv"):
            logs_df = pd.read_csv("result_logs.csv")
            st.dataframe(logs_df, use_container_width=True)
        else:
            st.write("Abhi tak kisi parent ne result search nahi kiya hai.")
            
        st.markdown("---")
        
        # 3. Student & Gallery Photo Upload Manager
        st.subheader("📸 Upload Student & School Photos")
        col_u1, col_u2 = st.columns(2)
        
        with col_u1:
            st.write("--- **Student Photo Upload** ---")
            st_roll = st.text_input("Enter Student Roll No for Photo")
            st_photo = st.file_uploader("Choose Student Image", type=["jpg", "jpeg", "png"], key="st_photo")
            if st.button("Upload Student Photo"):
                if st_roll and st_photo:
                    img_path = f"photos/students/{st_roll.strip()}.png"
                    img = Image.open(st_photo)
                    img.save(img_path)
                    st.success(f"✅ Student Photo Saved for Roll No: {st_roll}")
                else:
                    st.warning("Roll No aur Photo dono select karein.")

        with col_u2:
            st.write("--- **School Gallery Photo Upload** ---")
            gal_title = st.text_input("Gallery Image Title/Event Name")
            gal_photo = st.file_uploader("Choose Gallery Image", type=["jpg", "jpeg", "png"], key="gal_photo")
            if st.button("Upload to Gallery"):
                if gal_photo:
                    time_stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    img_path = f"photos/gallery/{time_stamp}.png"
                    img = Image.open(gal_photo)
                    img.save(img_path)
                    st.success("✅ Photo Uploaded to School Gallery!")
                else:
                    st.warning("Photo select karein.")

        st.markdown("---")
        
        # 5. File Upload with Auto-Backup
        st.subheader("📤 Excel Data Upload (With Auto-Backup)")
        uploaded_file = st.file_uploader("Upload New Excel Sheet (.xlsx)", type=["xlsx", "xls"])
        
        if uploaded_file is not None:
            # Create Auto-Backup of Old File
            if os.path.exists(EXCEL_FILE_PATH):
                backup_name = f"backups/Backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                shutil.copy(EXCEL_FILE_PATH, backup_name)
                st.info(f"🛡️ **Auto-Backup Saved:** `{backup_name}`")
            
            # Save New File
            with open(EXCEL_FILE_PATH, "wb") as f:
                f.write(uploaded_file.getbuffer())
                
            df = pd.read_excel(EXCEL_FILE_PATH)
            st.session_state["student_data"] = process_data(df)
            st.success("✅ New Excel File Saved & Live Updated!")


# ==============================================================================
# 👨‍🎓 PARENT PORTAL
# ==============================================================================
elif menu == "👨‍🎓 Parent Portal":
    
    st.markdown("### 🏫 Jawahar Navodaya Vidyalaya, Chhotaudepur")
    
    # Flash News Ticker
    if st.session_state["student_data"] is not None:
        df_data = st.session_state["student_data"]
        student_summary = df_data.groupby(['Class', 'Student_Name', 'Roll_No']).agg(Overall_Percentage=('Percentage', 'mean')).reset_index()
        school_topper = student_summary.sort_values(by='Overall_Percentage', ascending=False).iloc[0]
        
        news_text = f"🏆 <b>SCHOOL TOPPER:</b> {school_topper['Student_Name']} (Class {school_topper['Class']}) - {school_topper['Overall_Percentage']:.2f}% | "
        
        st.markdown(
            f"""
            <div style="background-color: #ffeb3b; padding: 10px; border-radius: 5px; color: #000; font-size: 16px;">
                <marquee scrollamount="6">{news_text}</marquee>
            </div>
            """,
            unsafe_allow_html=True
        )
        st.write("")

    # 4. Toppers Results With Photo
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
        
        # 8. Dual Search Option (Roll No + DOB OR Roll No + Aadhaar)
        search_method = st.radio("Choose Verification Method to View Result:", ["Option 1: Roll No & Date of Birth (DOB)", "Option 2: Roll No & Aadhaar Card Number"], horizontal=True)
        
        with st.form("search_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                selected_class = st.selectbox("Select Class", sorted(df['Class'].astype(str).unique()))
                roll_no = st.text_input("Roll No")
            
            if "Option 1" in search_method:
                with c2:
                    dob = st.text_input("Date of Birth (e.g., 15-08-2008)")
                with c3:
                    st.write("")
            else:
                with c2:
                    aadhaar = st.text_input("Aadhaar Card Number")
                with c3:
                    st.write("")
            
            submit_btn = st.form_submit_button("🔍 View Result")

        if submit_btn:
            # Search Filter Logic
            if "Option 1" in search_method:
                filtered_df = df[
                    (df['Class'].astype(str).str.strip().str.lower() == selected_class.strip().lower()) &
                    (df['Roll_No'].astype(str).str.strip() == roll_no.strip()) &
                    (df['DOB'].astype(str).str.strip() == dob.strip())
                ]
            else:
                filtered_df = df[
                    (df['Class'].astype(str).str.strip().str.lower() == selected_class.strip().lower()) &
                    (df['Roll_No'].astype(str).str.strip() == roll_no.strip()) &
                    (df['Aadhaar_No'].astype(str).str.strip() == aadhaar.strip())
                ]
            
            if filtered_df.empty:
                st.error("❌ Invalid Details! Kripya Roll No, DOB ya Aadhaar Number sahi se enter karein.")
            else:
                student_info = filtered_df.iloc[0]
                
                # Log Parent Search in Admin Tracker
                log_parent_search(student_info['Roll_No'], student_info['Student_Name'], student_info['Class'])
                
                st.success(f"🎓 Result Found for: **{student_info['Student_Name']}**")
                
                # Student Result Header with Student Photo
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
                
                # Cumulative Performance Summary
                st.subheader("📈 Overall Performance Summary")
                avg_pct = filtered_df['Percentage'].mean()
                total_obtained = filtered_df['Total_Marks'].sum()
                total_max = filtered_df['Max_Marks'].sum()
                
                k1, k2, k3 = st.columns(3)
                k1.metric("Overall Percentage", f"{avg_pct:.2f}%")
                k2.metric("Total Marks Obtained", f"{total_obtained} / {total_max}")
                k3.metric("Status", "PASS / EXCELLENT" if avg_pct >= 60 else "NEEDS IMPROVEMENT")
                
                st.markdown("---")
                
                # Scorecards
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
# 🖼️ SCHOOL GALLERY
# ==============================================================================
elif menu == "🖼️ School Gallery":
    st.header("🏫 Jawahar Navodaya Vidyalaya - Photo Gallery")
    st.markdown("---")
    
    gallery_files = os.listdir("photos/gallery")
    if len(gallery_files) == 0:
        st.info("Gallery me abhi koi photo upload nahi hui hai.")
    else:
        cols = st.columns(3)
        for idx, img_name in enumerate(gallery_files):
            img_path = os.path.join("photos/gallery", img_name)
            with cols[idx % 3]:
                st.image(img_path, use_container_width=True)


# ==============================================================================
# 🏆 BOARD EXAM RESULTS (CLASS 10 & 12)
# ==============================================================================
elif menu == "🏆 Board Exam Results":
    st.header("🎓 CBSE Board Exam Hall of Fame (Class 10 & 12 Last Year Toppers)")
    st.markdown("---")
    
    b_col1, b_col2 = st.columns(2)
    
    with b_col1:
        st.subheader("🥇 Class 12 CBSE Board Toppers")
        st.write("• **Ananya Sharma:** 98.4%")
        st.write("• **Karan Rathwa:** 96.2%")
        st.write("• **Priya Patel:** 95.8%")

    with b_col2:
        st.subheader("🥇 Class 10 CBSE Board Toppers")
        st.write("• **Rahul Tadvi:** 97.8%")
        st.write("• **Neha Verma:** 96.5%")
        st.write("• **Sanjay Parmar:** 95.0%")


# ==============================================================================
# 6 & 7. STYLISH DEVELOPER CREDIT & COPYRIGHT FOOTER
# ==============================================================================
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; padding: 10px;'>
        <p style='font-family: "Brush Script MT", cursive, sans-serif; font-size: 30px; color: #E91E63; margin-bottom: 5px;'>
            Developer: ANIL CHAUDHARY
        </p>
        <p style='font-family: Arial, sans-serif; font-size: 14px; font-weight: bold; color: #333; letter-spacing: 1px;'>
            © COPYRIGHTS JNV CHHOTAUDEPUR ALL RIGHTS RESERVED
        </p>
    </div>
    """,
    unsafe_allow_html=True
)