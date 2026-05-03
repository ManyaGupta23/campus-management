# ================= IMPORT =================
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
    "students": ["student_id","name","course","email"],
    "faculty": ["faculty_id","name","subject"],
    "schedule": ["class_id","faculty_id","room_id","time_slot","subject"],
    "attendance": ["student_id","class_id","date","status"],
    "users": ["username","password","role","student_id","faculty_id"],
    "results": ["student_id","subject","marks","grade"]
}

# ================= HELPERS =================
def clean(x):
    return str(x).strip().replace(".0","")

def grade(m):
    m = int(m)
    return "A+" if m>=90 else "A" if m>=75 else "B" if m>=60 else "C" if m>=40 else "F"

# ================= LOAD / SAVE =================
def load():
    if not os.path.exists(FILE):
        with pd.ExcelWriter(FILE, engine="openpyxl") as w:
            for s, cols in SCHEMA.items():
                pd.DataFrame(columns=cols).to_excel(w, sheet_name=s, index=False)

    db = {}
    xls = pd.ExcelFile(FILE)

    for s in SCHEMA:
        df = pd.read_excel(FILE, sheet_name=s) if s in xls.sheet_names else pd.DataFrame(columns=SCHEMA[s])
        df = df.astype(str)
        for c in df.columns:
            df[c] = df[c].apply(clean)
        db[s] = df

    return db

def save(db):
    with pd.ExcelWriter(FILE, engine="openpyxl", mode="w") as w:
        for s, df in db.items():
            df.to_excel(w, sheet_name=s, index=False)
    st.session_state.db = load()

# ================= REPORT CARD =================
def report_pdf(sid, name, course, df):

    file = f"{sid}_report.pdf"
    doc = SimpleDocTemplate(file, pagesize=A4)
    styles = getSampleStyleSheet()
    content = []

    title = ParagraphStyle(
        name="title",
        fontSize=22,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#C9A227")
    )

    content.append(Paragraph("ACADEMIC REPORT CARD", title))
    content.append(Spacer(1, 10))

    content.append(Paragraph(f"<b>Name:</b> {name}", styles["Normal"]))
    content.append(Paragraph(f"<b>ID:</b> {sid}", styles["Normal"]))
    content.append(Paragraph(f"<b>Course:</b> {course}", styles["Normal"]))
    content.append(Spacer(1, 10))

    data = [["Subject","Marks","Grade"]]

    for _, r in df.iterrows():
        data.append([r["subject"], r["marks"], r["grade"]])

    table = Table(data)

    table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#C9A227")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("GRID",(0,0),(-1,-1),0.5,colors.black),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("PADDING",(0,0),(-1,-1),6),
    ]))

    content.append(table)

    doc.build(content)
    return file

# ================= SESSION =================
if "db" not in st.session_state:
    st.session_state.db = load()

if "login" not in st.session_state:
    st.session_state.login = False

# ================= LOGIN =================
if not st.session_state.login:

    st.title("🎓 Campus ERP Login")

    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    r = st.selectbox("Role", ["Admin","Faculty","Student"])

    if st.button("Login"):

        users = st.session_state.db["users"]

        if u == "admin" and p == "admin123":
            st.session_state.login = True
            st.session_state.role = "Admin"
            st.rerun()

        m = users[(users.username == u) & (users.password == p) & (users.role == r)]

        if not m.empty:
            st.session_state.login = True
            st.session_state.role = r

            st.session_state.link = (
                m.iloc[0]["faculty_id"] if r == "Faculty"
                else m.iloc[0]["student_id"]
            )

            st.rerun()

        else:
            st.error("Invalid Login")

# ================= MAIN =================
else:

    db = st.session_state.db
    role = st.session_state.role

    st.sidebar.title(role)

    if role == "Admin":
        menu = ["Dashboard","Manage Students","Manage Faculty","Manage Schedule","Logout"]

    elif role == "Faculty":
        menu = ["My Schedule","Students","Marks","Attendance","Report Card","View Data","Logout"]

    else:
        menu = ["My Results","My Attendance","Certificate","Logout"]

    ch = st.sidebar.radio("Menu", menu)

    # ================= ADMIN =================
    if role == "Admin":

        if ch == "Dashboard":
            st.metric("Students", len(db["students"]))
            st.metric("Faculty", len(db["faculty"]))
            st.metric("Schedule", len(db["schedule"]))

        elif ch == "Manage Students":
            st.dataframe(db["students"])

        elif ch == "Manage Faculty":
            st.dataframe(db["faculty"])

        elif ch == "Manage Schedule":
            st.dataframe(db["schedule"])

    # ================= FACULTY =================
    elif role == "Faculty":

        faculty_id = str(st.session_state.link)

        if ch == "My Schedule":
            st.subheader("📅 My Schedule")

            df = db["schedule"].copy()
            df["faculty_id"] = df["faculty_id"].astype(str)

            my_schedule = df[df["faculty_id"] == faculty_id]

            if my_schedule.empty:
                st.warning("No schedule assigned")
            else:
                st.dataframe(my_schedule)

        elif ch == "Students":
            st.dataframe(db["students"])

        elif ch == "Marks":
            sid = st.text_input("Student ID")
            sub = st.text_input("Subject")
            m = st.number_input("Marks", 0, 100)

            if st.button("Save"):
                db["results"] = pd.concat([
                    db["results"],
                    pd.DataFrame([[sid, sub, m, grade(m)]],
                    columns=SCHEMA["results"])
                ])
                save(db)

        elif ch == "Attendance":
            sid = st.text_input("Student ID")
            cid = st.text_input("Class ID")
            status = st.selectbox("Status", ["Present","Absent"])

            if st.button("Save"):
                db["attendance"] = pd.concat([
                    db["attendance"],
                    pd.DataFrame([[sid, cid, str(datetime.today().date()), status]],
                    columns=SCHEMA["attendance"])
                ])
                save(db)

        elif ch == "Report Card":

            sid = st.text_input("Student ID")

            if st.button("Generate"):

                student = db["students"][db["students"]["student_id"] == sid]
                marks = db["results"][db["results"]["student_id"] == sid]

                if student.empty:
                    st.error("Student not found")
                else:
                    name = student.iloc[0]["name"]
                    course = student.iloc[0]["course"]

                    file = report_pdf(sid, name, course, marks)

                    st.success("Report Generated")

                    with open(file, "rb") as f:
                        st.download_button("Download Report", f, file_name=file)

        elif ch == "View Data":
            st.write("Marks")
            st.dataframe(db["results"])

            st.write("Attendance")
            st.dataframe(db["attendance"])

    # ================= STUDENT =================
    elif role == "Student":

        sid = str(st.session_state.link)

        if ch == "My Results":
            df = db["results"].copy()
            df["student_id"] = df["student_id"].astype(str)
            st.dataframe(df[df["student_id"] == sid])

        elif ch == "My Attendance":
            df = db["attendance"].copy()
            df["student_id"] = df["student_id"].astype(str)
            st.dataframe(df[df["student_id"] == sid])

        elif ch == "Certificate":
            st.success("Certificate feature already exists in your code")
