# ================= CAMPUS ERP (FINAL DEPLOY VERSION) =================

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

# ================= UTIL =================
def clean(x):
    return str(x).strip().replace(".0","")

def grade(m):
    m = int(m)
    return "A+" if m>=90 else "A" if m>=75 else "B" if m>=60 else "C" if m>=40 else "F"

# ================= FILE HANDLING (FIXED) =================
def load():
    if not os.path.exists(FILE):
        with pd.ExcelWriter(FILE, engine="openpyxl") as w:
            for s, cols in SCHEMA.items():
                pd.DataFrame(columns=cols).to_excel(w, sheet_name=s, index=False)

    db = {}
    xls = pd.ExcelFile(FILE)

    for s in SCHEMA:
        if s in xls.sheet_names:
            df = pd.read_excel(FILE, sheet_name=s)
        else:
            df = pd.DataFrame(columns=SCHEMA[s])

        df = df.astype(str)

        # ✅ FIXED (no applymap error)
        for col in df.columns:
            df[col] = df[col].apply(clean)

        db[s] = df

    return db


def save(db):
    with pd.ExcelWriter(FILE, engine="openpyxl", mode="w") as w:
        for s, df in db.items():
            df.to_excel(w, sheet_name=s, index=False)
    st.session_state.db = load()

# ================= QR =================
def make_qr(data, file):
    img = qrcode.make(data)
    img.save(file)
    return file

# ================= REPORT PDF =================
def report_pdf(sid,name,course,df):
    file = f"{sid}_report.pdf"
    doc = SimpleDocTemplate(file, pagesize=A4)
    styles = getSampleStyleSheet()
    content = []

    title = ParagraphStyle(name="t", fontSize=26, alignment=TA_CENTER)
    content.append(Paragraph("ACADEMIC REPORT CARD", title))
    content.append(Spacer(1,20))

    qr = make_qr(f"{sid}-{name}", f"{sid}_qr.png")
    content.append(Image(qr,2*inch,2*inch))
    content.append(Paragraph(f"Name: {name}", styles["Normal"]))
    content.append(Paragraph(f"Course: {course}", styles["Normal"]))

    data = [["Subject","Marks","Grade"]]
    for _,r in df.iterrows():
        data.append([r["subject"],r["marks"],r["grade"]])

    table = Table(data)
    table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.grey),
        ("GRID",(0,0),(-1,-1),1,colors.black)
    ]))

    content.append(Spacer(1,15))
    content.append(table)

    doc.build(content)
    return file

# ================= CERTIFICATE =================
def certificate_pdf(sid,name,course):
    file = f"{sid}_certificate.pdf"
    doc = SimpleDocTemplate(file, pagesize=A4)
    styles = getSampleStyleSheet()
    content = []

    title = ParagraphStyle(name="t", fontSize=28, alignment=TA_CENTER)
    content.append(Paragraph("CERTIFICATE OF COMPLETION", title))
    content.append(Spacer(1,30))

    para = ParagraphStyle(name="p", alignment=TA_CENTER, fontSize=16)
    content.append(Paragraph(f"This is to certify that <b>{name}</b>", para))
    content.append(Paragraph(f"has successfully completed <b>{course}</b>", para))

    qr = make_qr(f"CERT-{sid}", f"{sid}_certqr.png")
    content.append(Spacer(1,20))
    content.append(Image(qr,2*inch,2*inch,hAlign='CENTER'))

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

        if u=="admin" and p=="admin123":
            st.session_state.login=True
            st.session_state.role="Admin"
            st.rerun()

        m = users[(users.username==u)&(users.password==p)&(users.role==r)]

        if not m.empty:
            st.session_state.login=True
            st.session_state.role=r
            st.session_state.link = m.iloc[0]["student_id"] if r=="Student" else m.iloc[0]["faculty_id"]
            st.rerun()
        else:
            st.error("Invalid login")

# ================= MAIN =================
else:
    db = st.session_state.db
    role = st.session_state.role

    st.sidebar.title(f"👤 {role}")

    if role=="Admin":
        menu=["Dashboard","Manage Students","Manage Faculty","Logout"]
    elif role=="Faculty":
        menu=["My Schedule","Marks","Attendance","Logout"]
    else:
        menu=["My Results","My Attendance","Certificate","Report Card","Logout"]

    ch = st.sidebar.radio("Menu", menu)

    # ================= ADMIN =================
    if role=="Admin":

        if ch=="Dashboard":
            st.metric("Students", len(db["students"]))
            st.metric("Faculty", len(db["faculty"]))

        elif ch=="Manage Students":
            st.dataframe(db["students"])

            sid=st.text_input("Student ID")
            name=st.text_input("Name")

            if st.button("Add Student"):
                if sid in db["students"]["student_id"].values:
                    st.error("Student already exists")
                else:
                    db["students"] = pd.concat([
                        db["students"],
                        pd.DataFrame([[sid,name,"","","","","","","",""]],
                                     columns=SCHEMA["students"])
                    ])
                    save(db)

        elif ch=="Manage Faculty":
            st.dataframe(db["faculty"])

    # ================= FACULTY =================
    elif role=="Faculty":

        fid = str(st.session_state.link).strip()

        if ch=="My Schedule":
            df = db["schedule"].copy()
            df["faculty_id"] = df["faculty_id"].astype(str).str.strip()
            my = df[df["faculty_id"] == fid]

            if my.empty:
                st.warning("No schedule assigned")
            else:
                st.dataframe(my)

        elif ch=="Marks":
            sid=st.text_input("Student ID")
            sub=st.text_input("Subject")
            m=st.number_input("Marks",0,100)

            if st.button("Save Marks"):
                db["results"] = pd.concat([
                    db["results"],
                    pd.DataFrame([[sid,sub,m,grade(m)]], columns=SCHEMA["results"])
                ])
                save(db)

        elif ch=="Attendance":
            sid=st.text_input("Student ID")
            cid=st.text_input("Class ID")

            if st.button("Save Attendance"):
                db["attendance"] = pd.concat([
                    db["attendance"],
                    pd.DataFrame([[sid,cid,str(datetime.today().date()),"Present"]],
                                 columns=SCHEMA["attendance"])
                ])
                save(db)

    # ================= STUDENT =================
    elif role=="Student":

        sid = str(st.session_state.link).strip()

        if ch=="My Results":
            df = db["results"].copy()
            df["student_id"] = df["student_id"].astype(str).str.strip()
            my = df[df["student_id"]==sid]

            st.dataframe(my)

            if not my.empty:
                info=db["students"][db["students"]["student_id"].astype(str).str.strip()==sid]
                name = info.iloc[0]["name"] if not info.empty else "Student"
                course = info.iloc[0]["course"] if not info.empty else "Course"

                if st.button("📥 Download Report"):
                    f = report_pdf(sid,name,course,my)
                    with open(f,"rb") as file:
                        st.download_button("Download Report",file.read(),file_name=f)

        elif ch=="My Attendance":
            df=db["attendance"].copy()
            df["student_id"] = df["student_id"].astype(str).str.strip()
            st.dataframe(df[df["student_id"]==sid])

        elif ch=="Certificate":
            info=db["students"][db["students"]["student_id"].astype(str).str.strip()==sid]
            name=info.iloc[0]["name"] if not info.empty else "Student"
            course=info.iloc[0]["course"] if not info.empty else "Course"

            if st.button("Download Certificate"):
                f=certificate_pdf(sid,name,course)
                with open(f,"rb") as file:
                    st.download_button("Download",file.read(),file_name=f)

        elif ch=="Report Card":
            df=db["results"].copy()
            df["student_id"] = df["student_id"].astype(str).str.strip()
            my=df[df["student_id"]==sid]

            if my.empty:
                st.warning("No results found")
            else:
                info=db["students"][db["students"]["student_id"].astype(str).str.strip()==sid]
                name=info.iloc[0]["name"] if not info.empty else "Student"
                course=info.iloc[0]["course"] if not info.empty else "Course"

                if st.button("Download Report Card"):
                    f=report_pdf(sid,name,course,my)
                    with open(f,"rb") as file:
                        st.download_button("Download",file.read(),file_name=f)

    if ch=="Logout":
        st.session_state.clear()
        st.rerun()
