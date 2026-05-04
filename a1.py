import streamlit as st
import pandas as pd
import os
from datetime import datetime
import qrcode
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import inch

# ================= CONFIG =================
st.set_page_config(page_title="Campus ERP", layout="wide")

FILE = "campus_flow_template.xlsx"

SCHEMA = {
    "students": ["student_id","name","gender","course","year","section","admission_date","attendance_percentage","status","email"],
    "faculty": ["faculty_id","name","subject"],
    "rooms": ["room_id","capacity","type"],
    "schedule": ["class_id","faculty_id","room_id","time_slot","subject"],
    "attendance": ["student_id","class_id","date","status"],
    "users": ["username","password","role","student_id","faculty_id"],
    "results": ["student_id","subject","marks","grade"]
}

# ================= UTIL FUNCTIONS =================
def clean(x):
    return str(x).strip().replace(".0","")

def grade_calc(m):
    try:
        m = int(m)
        return "A+" if m>=90 else "A" if m>=75 else "B" if m>=60 else "C" if m>=40 else "F"
    except:
        return "N/A"

def load():
    if not os.path.exists(FILE):
        with pd.ExcelWriter(FILE, engine="openpyxl") as w:
            for s, cols in SCHEMA.items():
                pd.DataFrame(columns=cols).to_excel(w, sheet_name=s, index=False)
    
    db = {}
    xls = pd.ExcelFile(FILE)
    for s in SCHEMA:
        df = pd.read_excel(FILE, sheet_name=s) if s in xls.sheet_names else pd.DataFrame(columns=SCHEMA[s])
        df = df.astype(str).applymap(clean)
        db[s] = df
    return db

def save(db):
    with pd.ExcelWriter(FILE, engine="openpyxl") as w:
        for s, df in db.items():
            df.to_excel(w, sheet_name=s, index=False)
    st.session_state.db = load()

# ================= PDF GENERATORS =================
def make_qr(data, file):
    img = qrcode.make(data)
    img.save(file)
    return file

def report_pdf(sid, name, course, df):
    file = f"{sid}_report.pdf"
    doc = SimpleDocTemplate(file, pagesize=A4)
    styles = getSampleStyleSheet()
    content = []
    
    title_style = ParagraphStyle(name="t", fontSize=24, alignment=TA_CENTER, textColor=colors.HexColor("#2563eb"))
    content.append(Paragraph("ACADEMIC REPORT CARD", title_style))
    content.append(Spacer(1, 20))

    qr_path = make_qr(f"ID:{sid}|Name:{name}", f"{sid}_qr.png")
    
    info_table = Table([
        [Image(qr_path, 1.2*inch, 1.2*inch), 
         [Paragraph(f"<b>Name:</b> {name}", styles["Normal"]),
          Paragraph(f"<b>Student ID:</b> {sid}", styles["Normal"]),
          Paragraph(f"<b>Course:</b> {course}", styles["Normal"])]]
    ], colWidths=[1.5*inch, 4*inch])
    
    content.append(info_table)
    content.append(Spacer(1, 20))

    data = [["Subject", "Marks", "Grade"]]
    for _, r in df.iterrows():
        data.append([r["subject"], r["marks"], r["grade"]])

    table = Table(data, colWidths=[2.5*inch, 1.5*inch, 1.5*inch])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2563eb")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.grey)
    ]))
    content.append(table)
    doc.build(content)
    return file

# ================= APP LOGIC =================
if "db" not in st.session_state:
    st.session_state.db = load()
if "login" not in st.session_state:
    st.session_state.login = False

if not st.session_state.login:
    st.title("🎓 Campus ERP Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    r = st.selectbox("Role", ["Admin", "Faculty", "Student"])

    if st.button("Login"):
        users = st.session_state.db["users"]
        if u == "admin" and p == "admin123":
            st.session_state.login, st.session_state.role = True, "Admin"
            st.rerun()
        
        match = users[(users.username == u) & (users.password == p) & (users.role == r)]
        if not match.empty:
            st.session_state.login = True
            st.session_state.role = r
            st.session_state.link = clean(match.iloc[0]["student_id"]) if r == "Student" else clean(match.iloc[0]["faculty_id"])
            st.rerun()
        else:
            st.error("Invalid Credentials")

else:
    db = st.session_state.db
    role = st.session_state.role
    st.sidebar.title(f"👤 {role}")
    
    if role == "Admin":
        menu = ["Dashboard", "Manage Students", "Manage Faculty", "Logout"]
    elif role == "Faculty":
        menu = ["My Schedule", "Mark Attendance", "Enter Marks", "Logout"]
    else:
        menu = ["My Results", "My Attendance", "Logout"]

    ch = st.sidebar.radio("Navigation", menu)

    # --- FACULTY SECTION ---
    if role == "Faculty":
        fid = st.session_state.link
        
        if ch == "My Schedule":
            st.header("📅 My Class Schedule")
            sched = db["schedule"]
            my_sched = sched[sched["faculty_id"] == fid]
            if my_sched.empty:
                st.info("No classes assigned to you.")
            else:
                st.table(my_sched)

        elif ch == "Enter Marks":
            st.header("📝 Grade Students")
            sid = st.selectbox("Select Student", db["students"]["student_id"])
            sub = st.text_input("Subject")
            m = st.number_input("Marks", 0, 100)
            if st.button("Submit Marks"):
                new_row = pd.DataFrame([[sid, sub, m, grade_calc(m)]], columns=SCHEMA["results"])
                db["results"] = pd.concat([db["results"], new_row], ignore_index=True)
                save(db)
                st.success("Marks Saved!")

    # --- STUDENT SECTION ---
    elif role == "Student":
        sid = st.session_state.link
        
        if ch == "My Results":
            st.header("📊 My Academic Performance")
            res = db["results"]
            my_res = res[res["student_id"] == sid]
            
            if my_res.empty:
                st.warning("Results haven't been uploaded yet.")
            else:
                st.dataframe(my_res, use_container_width=True)
                
                # Report Card Generation
                info = db["students"][db["students"]["student_id"] == sid]
                name = info.iloc[0]["name"] if not info.empty else "Student"
                course = info.iloc[0]["course"] if not info.empty else "General"
                
                if st.button("Generate & Download PDF Report Card"):
                    pdf_path = report_pdf(sid, name, course, my_res)
                    with open(pdf_path, "rb") as f:
                        st.download_button("Click here to Download", f, file_name=pdf_path)

        elif ch == "My Attendance":
            st.header("✔️ My Attendance Record")
            att = db["attendance"]
            st.dataframe(att[att["student_id"] == sid])

    if ch == "Logout":
        st.session_state.clear()
        st.rerun()
