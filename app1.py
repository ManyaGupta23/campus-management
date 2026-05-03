import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px

# PDF
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import inch

# ================= CONFIG =================
st.set_page_config(page_title="Campus ERP", layout="wide")

FILE_NAME = "campus_flow_template.xlsx"

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
def clean(x):
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
            df = pd.read_excel(FILE_NAME, sheet_name=s)
        else:
            df = pd.DataFrame(columns=SCHEMA[s])

        for col in df.columns:
            df[col] = df[col].apply(clean)

        db[s] = df

    return db

def save_data(db):
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="w") as writer:
        for s, df in db.items():
            df.to_excel(writer, sheet_name=s, index=False)

    st.session_state.db = load_data()

# ================= GRADE =================
def get_grade(m):
    m = int(m)
    if m >= 90: return "A+"
    elif m >= 75: return "A"
    elif m >= 60: return "B"
    elif m >= 40: return "C"
    return "F"

# ================= PDF =================
def generate_university_report(sid, name, course, df):

    file = f"{sid}_report.pdf"
    doc = SimpleDocTemplate(file, pagesize=A4)
    styles = getSampleStyleSheet()

    title = ParagraphStyle(name="t", fontSize=28, alignment=TA_CENTER,
                           textColor=colors.HexColor("#C9A227"))

    subtitle = ParagraphStyle(name="s", fontSize=16, alignment=TA_CENTER)

    content = []

    # LOGO
    if os.path.exists("logo.png"):
        content.append(Image("logo.png", width=1.2*inch, height=1.2*inch, hAlign='CENTER'))

    content.append(Paragraph("CAMPUS ERP UNIVERSITY", title))
    content.append(Paragraph("Official Academic Transcript", subtitle))
    content.append(Spacer(1, 20))

    # STUDENT INFO
    info = [
        ["Name", name],
        ["Student ID", sid],
        ["Course", course],
        ["Date", str(datetime.today().date())]
    ]

    table = Table(info, colWidths=[120, 300])
    table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.5,colors.grey),
        ("BACKGROUND",(0,0),(0,-1),colors.HexColor("#f5f5f5"))
    ]))

    content.append(table)
    content.append(Spacer(1,20))

    # MARKS TABLE
    data = [["Subject","Marks","Grade"]]
    for _, r in df.iterrows():
        data.append([r["subject"], int(r["marks"]), r["grade"]])

    table2 = Table(data)
    table2.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#C9A227")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("GRID",(0,0),(-1,-1),1,colors.black)
    ]))

    content.append(table2)

    total = df["marks"].astype(int).sum()
    avg = df["marks"].astype(int).mean()

    content.append(Spacer(1,20))
    content.append(Paragraph(f"Total: {total}", styles["Normal"]))
    content.append(Paragraph(f"Average: {round(avg,2)}", styles["Normal"]))

    # RESULT
    if avg >= 75:
        res = "DISTINCTION"
    elif avg >= 60:
        res = "FIRST DIVISION"
    elif avg >= 40:
        res = "SECOND DIVISION"
    else:
        res = "FAIL"

    content.append(Spacer(1,10))
    content.append(Paragraph(f"Result: {res}", styles["Normal"]))

    # SIGNATURE
    if os.path.exists("signature.png"):
        content.append(Spacer(1,40))
        content.append(Paragraph("Authorized Signature", styles["Normal"]))
        content.append(Image("signature.png", width=2*inch, height=0.8*inch))

    # BORDER
    def border(c, d):
        c.setStrokeColor(colors.HexColor("#C9A227"))
        c.setLineWidth(3)
        c.rect(20,20,555,800)

    doc.build(content, onFirstPage=border)

    return file

# ================= SESSION =================
if "db" not in st.session_state:
    st.session_state.db = load_data()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.link_id = None

# ================= LOGIN =================
if not st.session_state.logged_in:
    st.title("🏫 Login")

    u = clean(st.text_input("Username"))
    p = clean(st.text_input("Password", type="password"))
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
            (users["role"].str.lower() == r.lower())
        ]

        if not match.empty:
            st.session_state.logged_in = True
            st.session_state.role = r

            row = match.iloc[0]
            if r.lower() == "student":
                st.session_state.link_id = row["student_id"]

            st.rerun()
        else:
            st.error("Invalid login")

# ================= MAIN =================
else:
    db = st.session_state.db
    role = st.session_state.role

    menu = ["Dashboard","Students","Analytics","Logout"] if role=="Admin" else \
           ["Marks Entry","Logout"] if role=="Faculty" else \
           ["My Results","Logout"]

    choice = st.sidebar.radio("Menu", menu)

    if choice == "Dashboard":
        st.metric("Students", len(db["students"]))

    elif choice == "Students":
        st.dataframe(db["students"])

    elif choice == "Analytics":
        res = db["results"]
        if not res.empty:
            st.plotly_chart(px.histogram(res, x="marks"))

    elif choice == "Marks Entry":
        sid = clean(st.text_input("Student ID"))
        sub = st.text_input("Subject")
        m = st.number_input("Marks",0,100)

        if st.button("Save"):
            grade = get_grade(m)
            new = pd.DataFrame([[sid,sub,m,grade]], columns=SCHEMA["results"])
            db["results"] = pd.concat([db["results"],new])
            save_data(db)
            st.success("Saved")

    elif choice == "My Results":
        sid = clean(st.session_state.link_id)

        df = db["results"]
        student_df = df[df["student_id"] == sid]

        info = db["students"][db["students"]["student_id"] == sid]
        name = info.iloc[0]["name"] if not info.empty else "Student"
        course = info.iloc[0]["course"] if not info.empty else "Course"

        if student_df.empty:
            st.error("No result")
        else:
            st.table(student_df)

            if st.button("🎓 Download Report Card"):
                file = generate_university_report(sid, name, course, student_df)
                with open(file, "rb") as f:
                    st.download_button("Download PDF", f)

    elif choice == "Logout":
        st.session_state.clear()
        st.rerun()
