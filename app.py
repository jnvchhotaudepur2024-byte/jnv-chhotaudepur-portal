import os
import shutil
import datetime
import pandas as pd

# ==========================================
# 1. AUTO BACKUP SYSTEM (Desktop/jnv app backup)
# ==========================================
def perform_backup():
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    backup_folder = os.path.join(desktop_path, "jnv app backup")
    
    if not os.path.exists(backup_folder):
        os.makedirs(backup_folder)
        
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
    # Python Code Ka Backup
    current_script = __file__
    backup_script_name = f"code_backup_{timestamp}.txt"
    shutil.copy(current_script, os.path.join(backup_folder, backup_script_name))
    
    # Excel Data Ka Backup (Agar File Exists Karti Hai)
    if os.path.exists("jnv_app_data.xlsx"):
        backup_excel_name = f"data_backup_{timestamp}.xlsx"
        shutil.copy("jnv_app_data.xlsx", os.path.join(backup_folder, backup_excel_name))
        
    print(f"[OK] Backup successfully saved in: {backup_folder}")

# ==========================================
# 2. DEFAULT QUOTES & DATA GENERATOR
# ==========================================
DEFAULT_QUOTES = [
    "Education is the most powerful weapon which you can use to change the world. - Nelson Mandela",
    "Learning gives creativity, creativity leads to thinking, thinking provides knowledge. - A.P.J. Abdul Kalam",
    "Arise, awake, and stop not till the goal is reached. - Swami Vivekananda",
    "Live as if you were to die tomorrow. Learn as if you were to live forever. - Mahatma Gandhi",
    "The mind is not a vessel to be filled, but a fire to be kindled. - Plutarch",
    "Education is not preparation for life; education is life itself. - John Dewey",
    "Knowledge is power. Information is liberating. - Kofi Annan",
    "The beautiful thing about learning is that no one can take it away from you. - B.B. King"
]

# ==========================================
# 3. MARKSHEET GENERATOR (WATERMARK + CUMULATIVE)
# ==========================================
def generate_parent_portal():
    if not os.path.exists("jnv_app_data.xlsx"):
        print("[ERROR] jnv_app_data.xlsx file nahi mili!")
        return

    # Load Sheets
    df_students = pd.read_excel("jnv_app_data.xlsx", sheet_name="Students")
    df_marks = pd.read_excel("jnv_app_data.xlsx", sheet_name="Marks_Entry")
    
    try:
        df_fame = pd.read_excel("jnv_app_data.xlsx", sheet_name="Hall_Of_Fame")
        fame_items = [f"🏆 {row['Student_Name']} ({row['Class']}) - {row['Achievement']}" for _, row in df_fame.iterrows()]
    except Exception:
        fame_items = ["🏆 Amit Kumar (Class 12) - 98.4%", "🏆 Neha Singh (Class 10) - Science Olympiad Gold"]

    try:
        df_q = pd.read_excel("jnv_app_data.xlsx", sheet_name="Quotes")
        quotes_list = [f"“{row['Quote']}” - {row['Author']}" for _, row in df_q.iterrows()]
    except Exception:
        quotes_list = DEFAULT_QUOTES

    # Portal Output Directory
    portal_dir = "parents_portal"
    if not os.path.exists(portal_dir):
        os.makedirs(portal_dir)

    # Process Each Student
    for _, student in df_students.iterrows():
        roll = student['Roll_No']
        name = student['Student_Name']
        s_class = student['Class']
        section = student['Section']
        father = student['Father_Name']

        # Get all exam records for student
        s_marks = df_marks[df_marks['Roll_No'] == roll]
        
        if s_marks.empty:
            continue

        # Subject columns
        subjects = ['Hindi', 'English', 'Math', 'Science', 'SST']
        
        # Build Exam Rows HTML
        exam_rows_html = ""
        for _, m_row in s_marks.iterrows():
            total = sum(m_row[sub] for sub in subjects)
            exam_rows_html += f"""
            <tr>
                <td><b>{m_row['Exam_Type']}</b></td>
                <td>{m_row['Hindi']}</td>
                <td>{m_row['English']}</td>
                <td>{m_row['Math']}</td>
                <td>{m_row['Science']}</td>
                <td>{m_row['SST']}</td>
                <td><b>{total} / 500</b></td>
            </tr>
            """

        # Calculate Cumulative Averages
        cum_means = s_marks[subjects].mean().round(1)
        cum_total_avg = cum_means.sum().round(1)
        cum_percentage = round((cum_total_avg / 500) * 100, 2)

        cum_row_html = f"""
        <tr style="background-color: #e8f4f8; font-weight: bold;">
            <td>CUMULATIVE OVERALL</td>
            <td>{cum_means['Hindi']}</td>
            <td>{cum_means['English']}</td>
            <td>{cum_means['Math']}</td>
            <td>{cum_means['Science']}</td>
            <td>{cum_means['SST']}</td>
            <td>{cum_total_avg} / 500 ({cum_percentage}%)</td>
        </tr>
        """

        # Ticker Contents
        quotes_marquee = " &nbsp;&nbsp;&nbsp; | &nbsp;&nbsp;&nbsp; ".join(quotes_list)
        fame_marquee = " &nbsp;&nbsp;&nbsp; ★ &nbsp;&nbsp;&nbsp; ".join(fame_items)

        # HTML Template with Logo Background & Marquees
        html_content = f"""<!DOCTYPE html>
<html>
<head>
    <title>Parents Portal - Marksheet ({name})</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f4f6f9;
        }}
        .ticker-header {{
            background: #1e3c72;
            color: #fff;
            padding: 8px 0;
            font-size: 14px;
            font-weight: 500;
        }}
        .ticker-fame {{
            background: #d4af37;
            color: #000;
            padding: 6px 0;
            font-size: 14px;
            font-weight: bold;
        }}
        .container {{
            max-width: 900px;
            margin: 20px auto;
            background: #ffffff;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            position: relative;
            overflow: hidden;
        }}
        /* Watermark shri jnv logo background */
        .container::before {{
            content: "";
            background-image: url('jnv_logo.png'); /* Put jnv_logo.png in project folder */
            background-repeat: no-repeat;
            background-position: center;
            background-size: 350px;
            opacity: 0.08;
            top: 0; left: 0; bottom: 0; right: 0;
            position: absolute;
            z-index: 0;
        }}
        .content {{
            position: relative;
            z-index: 1;
        }}
        .header-title {{
            text-align: center;
            color: #0f2027;
            border-bottom: 2px solid #1e3c72;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }}
        .info-table, .marks-table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }}
        .info-table td {{
            padding: 8px;
            font-size: 15px;
        }}
        .marks-table th, .marks-table td {{
            border: 1px solid #cccccc;
            padding: 10px;
            text-align: center;
        }}
        .marks-table th {{
            background-color: #1e3c72;
            color: white;
        }}
    </style>
</head>
<body>

    <!-- TOP FLOATING QUOTES TICKER -->
    <div class="ticker-header">
        <marquee behavior="scroll" direction="left" scrollamount="5">
            🎓 {quotes_marquee}
        </marquee>
    </div>

    <!-- HALL OF FAME TICKER (RIGHT TO LEFT) -->
    <div class="ticker-fame">
        <marquee behavior="scroll" direction="left" scrollamount="6">
            {fame_marquee}
        </marquee>
    </div>

    <div class="container">
        <div class="content">
            <div class="header-title">
                <h2>PM SHRI JAWAHAR NAVODAYA VIDYALAYA</h2>
                <h3>STUDENT CUMULATIVE PERFORMANCE REPORT</h3>
            </div>

            <table class="info-table">
                <tr>
                    <td><b>Roll No:</b> {roll}</td>
                    <td><b>Student Name:</b> {name}</td>
                </tr>
                <tr>
                    <td><b>Class & Section:</b> {s_class} - {section}</td>
                    <td><b>Father's Name:</b> {father}</td>
                </tr>
            </table>

            <h4>Exam-wise & Cumulative Marks Breakup:</h4>
            <table class="marks-table">
                <thead>
                    <tr>
                        <th>Exam Name</th>
                        <th>Hindi</th>
                        <th>English</th>
                        <th>Math</th>
                        <th>Science</th>
                        <th>SST</th>
                        <th>Total</th>
                    </tr>
                </thead>
                <tbody>
                    {exam_rows_html}
                    {cum_row_html}
                </tbody>
            </table>
            
            <p style="text-align: right; margin-top: 40px;"><b>Principal Signature</b></p>
        </div>
    </div>

</body>
</html>
"""
        # Save / Overwrite existing report card automatically
        filename = os.path.join(portal_dir, f"marksheet_roll_{roll}.html")
        with open(filename, "w", encoding="utf-8") as f:
            f.write(html_content)

    print(f"[OK] Dynamic marksheets automatically updated in '{portal_dir}' folder!")

# ==========================================
# MAIN EXECUTION
# ==========================================
if __name__ == "__main__":
    # 1. Automatic Code & Data Backup
    perform_backup()
    
    # 2. Generate / Replace Parent Portal Marksheets
    generate_parent_portal()