import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import hashlib
import smtplib

# PDF
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import inch

# ================= CONFIG =================
st.set_page_config(page_title="Campus ERP", layout="wide")

FILE_NAME = "campus_flow.xlsx"

SHEETS = ["students","faculty","rooms","schedule","attendance","logs","users","results"]

SCHEMA = {
    "students": ["student_id","name","gender","course","year","section","admission_date","attendance_percentage","status","email"],
    "faculty": ["faculty_id","name","subject"],
    "rooms": ["room_id","capacity","type"],
    "schedule": ["class_id","faculty_id","room_id","time_slot","subject"],
    "attendance": ["student_id","class_id","date","status"],
    "logs": ["log_id","student_id","entry_time","exit_time"],
    "users": ["username","password","role","student_id","faculty_id"],
    "results": ["student_id","subject","marks","grade"]
}

# ================= UTIL =================
def hash_pass(p):
    return hashlib.sha256(p.encode()).hexdigest()

def clean_id(x):
    return str(x).strip().replace(".0","")

# ================= FILE =================
def create_file():
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
        for s, cols in SCHEMA.items():
            pd.DataFrame(columns=cols).to_excel(writer, sheet_name=s, index=False)

def load_data():
    if not os.path.exists(FILE_NAME):
        create_file()
    db = {}
    xls = pd.ExcelFile(FILE_NAME)
    for s in SHEETS:
        if s in xls.sheet_names:
            db[s] = pd.read_excel(FILE_NAME, sheet_name=s)
        else:
            db[s] = pd.DataFrame(columns=SCHEMA[s])
    return db

def save_data(db):
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="w") as writer:
        for s, df in db.items():
            df.to_excel(writer, sheet_name=s, index=False)
    st.session_state.db = load_data()

def get_grade(m):
    if m >= 90: return "A+"
    elif m >= 75: return "A"
    elif m >= 60: return "B"
    elif m >= 40: return "C"
    return "F"

# ================= EMAIL =================
def send_email(to_email, subject, body):
    sender = "your_email@gmail.com"
    password = "your_app_password"

    msg = f"Subject: {subject}\n\n{body}"

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender, password)
    server.sendmail(sender, to_email, msg)
    server.quit()

# ================= CERTIFICATE =================
def generate_certificate(name, sid, course):
    file = f"{sid}_certificate.pdf"
    doc = SimpleDocTemplate(file, pagesize=A4)

    styles = getSampleStyleSheet()

    title = ParagraphStyle(name="t", fontSize=28, alignment=TA_CENTER, textColor=colors.HexColor("#C9A227"))
    normal = ParagraphStyle(name="n", alignment=TA_CENTER)

    content = []

    try:
        content.append(Image("assets/logo.png", width=100, height=100))
    except: pass

    content.append(Paragraph("CERTIFICATE OF ACHIEVEMENT", title))
    content.append(Spacer(1,20))

    content.append(Paragraph(f"<b>{name}</b>", title))
    content.append(Paragraph(f"ID: {sid}", normal))
    content.append(Spacer(1,20))

    content.append(Paragraph(f"Completed <b>{course}</b>", normal))

    try:
        content.append(Image("assets/signature.png", width=120, height=50))
    except: pass

    def border(c, d):
        c.setStrokeColor(colors.HexColor("#C9A227"))
        c.rect(20,20,555,800)

    doc.build(content, onFirstPage=border)
    return file

# ================= REPORT =================
def generate_report_card(sid, name, course, df):
    file = f"{sid}_report.pdf"
    doc = SimpleDocTemplate(file, pagesize=A4)
    styles = getSampleStyleSheet()

    content = []

    content.append(Paragraph("<b>CAMPUS ERP</b>", styles["Title"]))
    content.append(Spacer(1,10))

    content.append(Paragraph(f"{name} | {sid} | {course}", styles["Normal"]))
    content.append(Spacer(1,20))

    data = [["Subject","Marks","Grade"]]
    for _, r in df.iterrows():
        data.append([r["subject"], r["marks"], r["grade"]])

    table = Table(data)
    table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.blue),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("GRID",(0,0),(-1,-1),1,colors.black)
    ]))

    content.append(table)

    total = df["marks"].sum()
    avg = df["marks"].mean()

    content.append(Spacer(1,20))
    content.append(Paragraph(f"Total: {int(total)} | Avg: {round(avg,2)}", styles["Normal"]))

    doc.build(content)
    return file

# ================= SESSION =================
if "db" not in st.session_state:
    st.session_state.db = load_data()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# ================= LOGIN =================
if not st.session_state.logged_in:
    st.title("Campus ERP Login")

    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    r = st.selectbox("Role", ["Admin","Faculty","Student"])

    if st.button("Login"):
        users = st.session_state.db["users"]

        if u == "admin" and p == "admin123":
            st.session_state.logged_in = True
            st.session_state.role = "Admin"
            st.rerun()

        match = users[
            (users["username"] == u) &
            (users["password"] == p) &
            (users["role"] == r)
        ]

        if not match.empty:
            st.session_state.logged_in = True
            st.session_state.role = r

            if r == "Student":
                st.session_state.link_id = clean_id(match.iloc[0]["student_id"])

            st.rerun()
        else:
            st.error("Invalid login")

# ================= APP =================
else:
    db = st.session_state.db
    role = st.session_state.role

    menu = ["Dashboard","Students","Analytics","Timetable","Logout"] if role=="Admin" else \
           ["Marks Entry","Attendance","Logout"] if role=="Faculty" else \
           ["My Results","Logout"]

    choice = st.sidebar.radio("Menu", menu)

    # ===== STUDENTS =====
    if choice == "Students":
        st.dataframe(db["students"])

        search = st.text_input("Search")
        if search:
            st.dataframe(db["students"][db["students"]["name"].str.contains(search,case=False,na=False)])

    # ===== MARKS =====
    elif choice == "Marks Entry":
        sid = st.text_input("Student ID")
        sub = st.text_input("Subject")
        m = st.number_input("Marks",0,100)

        if st.button("Save"):
            exists = db["results"][
                (db["results"]["student_id"]==sid) &
                (db["results"]["subject"]==sub)
            ]

            if not exists.empty:
                st.warning("Already exists")
            else:
                grade = get_grade(m)
                new = pd.DataFrame([[sid,sub,m,grade]], columns=SCHEMA["results"])
                db["results"] = pd.concat([db["results"],new])
                save_data(db)
                st.success("Saved")

    # ===== STUDENT DASHBOARD =====
    elif choice == "My Results":
        sid = st.session_state.link_id
        df = db["results"]
        student_df = df[df["student_id"].astype(str)==sid]

        if student_df.empty:
            st.error("No data")
        else:
            st.dataframe(student_df)

            if st.button("Download Certificate"):
                file = generate_certificate("Student", sid, "Course")
                with open(file,"rb") as f:
                    st.download_button("Download", f)

            if st.button("Download Report"):
                file = generate_report_card(sid,"Student","Course",student_df)
                with open(file,"rb") as f:
                    st.download_button("Download", f)

    # ===== LOGOUT =====
    elif choice == "Logout":
        st.session_state.clear()
        st.rerun()
