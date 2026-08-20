import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="JNV Chhotaudepur - Result Portal", page_icon="🎓", layout="wide")

st.title("🏫 Jawahar Navodaya Vidyalaya, Chhotaudepur")
st.subheader("📊 Student Performance, Result & Toppers Portal")
st.markdown("---")

# Data Processor Function
def process_data(df):
    meta_cols = ['Class', 'Roll_No', 'Student_Name', 'Father_Name', 'DOB', 'Exam_Type', 'Class_Teacher']
    subject_cols = [col for col in df.columns if col not in meta_cols]
    
    df['Total_Marks'] = df[subject_cols].sum(axis=1)
    df['Max_Marks'] = df['Exam_Type'].apply(lambda x: 150 if 'PWT' in str(x).upper() else 600)
    df['Percentage'] = (df['Total_Marks'] / df['Max_Marks']) * 100
    df['Class_Rank'] = df.groupby(['Class', 'Exam_Type'])['Total_Marks'].rank(ascending=False, method='min').astype(int)
    return df

# Permanent File Auto-Load Logic
EXCEL_FILE_PATH = "JNV_Student_Marks.xlsx"

if "student_data" not in st.session_state or st.session_state["student_data"] is None:
    if os.path.exists(EXCEL_FILE_PATH):
        try:
            default_df = pd.read_excel(EXCEL_FILE_PATH)
            st.session_state["student_data"] = process_data(default_df)
        except Exception as e:
            st.error(f"Error reading permanent Excel file: {e}")
    else:
        st.session_state["student_data"] = None

if "admin_logged_in" not in st.session_state:
    st.session_state["admin_logged_in"] = False

menu = st.sidebar.radio("Navigation", ["👨‍🎓 Parent Portal (Check Result)", "⚙️ Admin Portal (Protected)"])

def get_toppers_info(df):
    if df is None or df.empty:
        return None, None
    
    student_summary = df.groupby(['Class', 'Student_Name', 'Roll_No']).agg(
        Overall_Percentage=('Percentage', 'mean')
    ).reset_index()
    
    school_topper = student_summary.sort_values(by='Overall_Percentage', ascending=False).iloc[0]
    
    class_toppers = {}
    for cls in student_summary['Class'].unique():
        top_3 = student_summary[student_summary['Class'] == cls].sort_values(
            by='Overall_Percentage', ascending=False
        ).head(3)
        class_toppers[cls] = top_3
        
    return school_topper, class_toppers


# ================= ADMIN PORTAL =================
if menu == "⚙️ Admin Portal (Protected)":
    st.header("🔒 Admin Login & Dashboard")
    
    if not st.session_state["admin_logged_in"]:
        with st.form("login_form"):
            st.subheader("🔐 Enter Credentials")
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
        st.success("🔓 Authenticated as Admin")
        if st.button("🚪 Logout"):
            st.session_state["admin_logged_in"] = False
            st.rerun()
            
        st.markdown("---")
        st.subheader("📤 Temporary Data Upload (Session Overwrite)")
        
        uploaded_file = st.file_uploader("Upload New Excel File (.xlsx)", type=["xlsx", "xls"])
        
        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
            st.session_state["student_data"] = process_data(df)
            st.success("✅ Excel Sheet Successfully Loaded for Current Session!")
            st.dataframe(df)

        if os.path.exists(EXCEL_FILE_PATH):
            st.info(f"ℹ️ Permanent File Active: `{EXCEL_FILE_PATH}` is auto-loaded from repository.")


# ================= PARENT PORTAL =================
elif menu == "👨‍🎓 Parent Portal (Check Result)":
    
    # Flash News Ticker
    if st.session_state["student_data"] is not None:
        school_topper, class_toppers = get_toppers_info(st.session_state["student_data"])
        
        if school_topper is not None:
            news_text = f"🏆 <b>SCHOOL TOPPER:</b> {school_topper['Student_Name']} (Class {school_topper['Class']}) - {school_topper['Overall_Percentage']:.2f}% | "
            for cls, top_df in class_toppers.items():
                first_place = top_df.iloc[0]
                news_text += f"🥇 <b>Class {cls} Topper:</b> {first_place['Student_Name']} ({first_place['Overall_Percentage']:.2f}%) | "
            
            st.markdown(
                f"""
                <div style="background-color: #ffeb3b; padding: 10px; border-radius: 5px; color: #000; font-size: 16px;">
                    <marquee scrollamount="6">{news_text}</marquee>
                </div>
                """,
                unsafe_allow_html=True
            )
            st.write("")

    # Toppers Hall of Fame
    if st.session_state["student_data"] is not None:
        with st.expander("🏆 **VIEW CLASS TOPPERS & SCHOOL TOPPER (HALL OF FAME)**", expanded=False):
            school_topper, class_toppers = get_toppers_info(st.session_state["student_data"])
            
            st.markdown("### 🥇 School Overall Topper")
            st.info(f"🌟 **{school_topper['Student_Name']}** | Class: **{school_topper['Class']}** | Roll No: **{school_topper['Roll_No']}** | Percentage: **{school_topper['Overall_Percentage']:.2f}%**")
            
            st.markdown("---")
            st.markdown("### 🎖️ Class-Wise Top 3 Rankers")
            
            cols = st.columns(len(class_toppers))
            for i, (cls, top_df) in enumerate(class_toppers.items()):
                with cols[i]:
                    st.write(f"#### 📌 Class {cls}")
                    for rank, (_, row) in enumerate(top_df.iterrows(), start=1):
                        medal = "🥇" if rank == 1 else ("🥈" if rank == 2 else "🥉")
                        st.write(f"{medal} **Rank {rank}:** {row['Student_Name']}")
                        st.caption(f"Score: {row['Overall_Percentage']:.2f}% | Roll No: {row['Roll_No']}")

    st.markdown("---")
    st.header("🔎 Check Student Result")
    
    if st.session_state["student_data"] is None:
        st.warning("⚠️ Koi Data nahi mila. GitHub Repository me `JNV_Student_Marks.xlsx` file upload karein.")
    else:
        df = st.session_state["student_data"]
        
        with st.form("search_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                selected_class = st.selectbox("Select Class", sorted(df['Class'].astype(str).unique()))
                roll_no = st.text_input("Roll No")
            with c2:
                student_name = st.text_input("Student Name")
                father_name = st.text_input("Father's Name")
            with c3:
                dob = st.text_input("DOB (e.g., 15-08-2008)")
            
            submit_btn = st.form_submit_button("🔍 Check Result")

        if submit_btn:
            filtered_df = df[
                (df['Class'].astype(str).str.strip().str.lower() == selected_class.strip().lower()) &
                (df['Roll_No'].astype(str).str.strip() == roll_no.strip()) &
                (df['Student_Name'].astype(str).str.strip().str.lower() == student_name.strip().lower()) &
                (df['Father_Name'].astype(str).str.strip().str.lower() == father_name.strip().lower()) &
                (df['DOB'].astype(str).str.strip() == dob.strip())
            ]
            
            if filtered_df.empty:
                st.error("❌ Record Nahi Mila! Details dobara check karein.")
            else:
                student_info = filtered_df.iloc[0]
                st.success(f"🎓 Result Found: **{student_info['Student_Name']}** (Roll No: {student_info['Roll_No']})")
                
                b1, b2, b3 = st.columns(3)
                b1.write(f"**Class:** {student_info['Class']}")
                b2.write(f"**Father's Name:** {student_info['Father_Name']}")
                b3.write(f"**Class Teacher:** {student_info['Class_Teacher']}")
                st.markdown("---")
                
                st.subheader("📈 Cumulative Performance Summary")
                avg_pct = filtered_df['Percentage'].mean()
                total_obtained = filtered_df['Total_Marks'].sum()
                total_max = filtered_df['Max_Marks'].sum()
                exams_given = len(filtered_df)
                
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("Overall Average %", f"{avg_pct:.2f}%")
                k2.metric("Total Score", f"{total_obtained} / {total_max}")
                k3.metric("Exams Recorded", f"{exams_given}")
                k4.metric("Status", "PASS / EXCELLENT" if avg_pct >= 60 else "NEEDS IMPROVEMENT")
                
                st.markdown("---")
                
                st.subheader("📝 Exam Scorecards")
                meta_cols = ['Class', 'Roll_No', 'Student_Name', 'Father_Name', 'DOB', 'Exam_Type', 'Class_Teacher', 'Total_Marks', 'Max_Marks', 'Percentage', 'Class_Rank']
                subject_cols = [col for col in df.columns if col not in meta_cols]
                
                for index, row in filtered_df.iterrows():
                    with st.expander(f"📌 **{row['Exam_Type']}** | Score: {row['Total_Marks']}/{row['Max_Marks']} ({row['Percentage']:.2f}%) | Rank: #{row['Class_Rank']}", expanded=True):
                        m_df = pd.DataFrame({
                            'Subject': subject_cols,
                            'Marks Obtained': [row[sub] for sub in subject_cols]
                        })
                        st.dataframe(m_df.T, use_container_width=True)