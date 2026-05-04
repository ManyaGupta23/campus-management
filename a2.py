# ================= CAMPUS ERP (FINAL BUG-FREE VERSION) =================

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

st.set_page_config(page_title="Campus ERP", layout="wide")

FILE = "campus_flow_template.xlsx"

SCHEMA = {
    "students": ["student_id","name","course"],
    "faculty": ["faculty_id","name"],
    "schedule": ["class_id","faculty_id","subject"],
    "attendance": ["student_id","class_id","date","status"],
    "users": ["username","password","role","student_id","faculty_id"],
    "results": ["student_id","subject","marks","grade"]
}

# ================= UTIL =================
def clean(x):
    return str(x).strip()

def grade(m):
    m = int(m)
    return "A+" if m>=90 else "A" if m>=75 else "B" if m>=60 else "C" if m>=40 else "F"

# ================= FILE =================
def load():
    if not os.path.exists(FILE):
        with pd.ExcelWriter(FILE) as w:
            for s in SCHEMA:
                pd.DataFrame(columns=SCHEMA[s]).to_excel(w, sheet_name=s, index=False)

    db = {}
    xls = pd.ExcelFile(FILE)

    for s in SCHEMA:
        df = pd.read_excel(FILE, sheet_name=s) if s in xls.sheet_names else pd.DataFrame(columns=SCHEMA[s])
        df = df.astype(str)

        for col in df.columns:
            df[col] = df[col].apply(clean)

        db[s] = df

    return db

def save(db):
    with pd.ExcelWriter(FILE, mode="w") as w:
        for s in db:
            db[s].to_excel(w, sheet_name=s, index=False)
    st.session_state.db = load()

# ================= QR =================
def make_qr(data, file):
    img = qrcode.make(data)
    img.save(file)
    return file

# ================= REPORT =================
def report_pdf(sid, name, course, df):
    file = f"{sid}_report.pdf"
    doc = SimpleDocTemplate(file, pagesize=A4)
    styles = getSampleStyleSheet()
    content = []

    content.append(Paragraph("REPORT CARD", styles["Title"]))
    content.append(Spacer(1, 20))

    qr = make_qr(sid, f"{sid}.png")

    info = [[Image(qr,2*inch,2*inch),
            [Paragraph(f"Name: {name}", styles["Normal"]),
             Paragraph(f"ID: {sid}", styles["Normal"]),
             Paragraph(f"Course: {course}", styles["Normal"])]]]

    content.append(Table(info))

    data = [["Subject","Marks","Grade"]]
    for _,r in df.iterrows():
        data.append([r["subject"],r["marks"],r["grade"]])

    content.append(Table(data))
    doc.build(content)
    return file

# ================= CERT =================
def certificate_pdf(sid, name, course):
    file = f"{sid}_cert.pdf"
    doc = SimpleDocTemplate(file, pagesize=A4)
    styles = getSampleStyleSheet()

    content = []
    content.append(Paragraph("CERTIFICATE", styles["Title"]))
    content.append(Spacer(1,20))
    content.append(Paragraph(f"{name} completed {course}", styles["Normal"]))

    qr = make_qr(sid, f"{sid}_cert.png")
    content.append(Image(qr,2*inch,2*inch))

    doc.build(content)
    return file

# ================= SESSION =================
if "db" not in st.session_state:
    st.session_state.db = load()

if "login" not in st.session_state:
    st.session_state.login = False

# ================= LOGIN =================
if not st.session_state.login:

    col1,col2 = st.columns([1.5,1])

    with col1:
        st.image("https://images.unsplash.com/photo-1523050854058-8df90110c9f1", use_container_width=True)

    with col2:
        st.title("Login")

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
                st.session_state.link = clean(m.iloc[0]["student_id"] if r=="Student" else m.iloc[0]["faculty_id"])
                st.rerun()
            else:
                st.error("Invalid login")

# ================= MAIN =================
else:

    db = st.session_state.db
    role = st.session_state.role

    st.sidebar.title(role)

    if role=="Admin":
        menu=["Students","Faculty","Schedule","Logout"]
    elif role=="Faculty":
        menu=["My Schedule","Marks","Attendance","Logout"]
    else:
        menu=["My Results","Attendance","Certificate","Logout"]

    ch = st.sidebar.radio("Menu", menu)

# ================= ADMIN =================
    if role=="Admin":

        if ch=="Students":
            st.dataframe(db["students"])

            sid=st.text_input("Student ID")
            name=st.text_input("Name")
            course=st.text_input("Course")

            if st.button("Add Student"):
                db["students"]=pd.concat([db["students"],pd.DataFrame([[sid,name,course]],columns=SCHEMA["students"])])
                db["users"]=pd.concat([db["users"],pd.DataFrame([[sid,"1234","Student",sid,""]],columns=SCHEMA["users"])])
                save(db)

            del_id=st.text_input("Delete Student ID")
            if st.button("Delete Student"):
                db["students"]=db["students"][db["students"]["student_id"]!=del_id]
                save(db)

        elif ch=="Faculty":
            st.dataframe(db["faculty"])

            fid=st.text_input("Faculty ID")
            name=st.text_input("Name")

            if st.button("Add Faculty"):
                db["faculty"]=pd.concat([db["faculty"],pd.DataFrame([[fid,name]],columns=SCHEMA["faculty"])])
                db["users"]=pd.concat([db["users"],pd.DataFrame([[fid,"1234","Faculty","",fid]],columns=SCHEMA["users"])])
                save(db)

        elif ch=="Schedule":
            st.dataframe(db["schedule"])

            cid=st.text_input("Class ID")
            fid=st.text_input("Faculty ID")
            sub=st.text_input("Subject")

            if st.button("Add Schedule"):
                db["schedule"]=pd.concat([db["schedule"],pd.DataFrame([[cid,fid,sub]],columns=SCHEMA["schedule"])])
                save(db)

# ================= FACULTY =================
    elif role=="Faculty":

        fid = clean(st.session_state.link)

        if ch=="My Schedule":
            df=db["schedule"].copy()
            df["faculty_id"]=df["faculty_id"].apply(clean)
            my=df[df["faculty_id"]==fid]

            if my.empty:
                st.warning("No schedule")
            else:
                st.dataframe(my)

        elif ch=="Marks":
            sid=st.text_input("Student ID")
            sub=st.text_input("Subject")
            m=st.number_input("Marks",0,100)

            if st.button("Save"):
                db["results"]=pd.concat([db["results"],pd.DataFrame([[sid,sub,m,grade(m)]],columns=SCHEMA["results"])])
                save(db)

        elif ch=="Attendance":
            sid=st.text_input("Student ID")
            cid=st.text_input("Class ID")
            status=st.selectbox("Status",["Present","Absent"])

            if st.button("Save Attendance"):
                db["attendance"]=pd.concat([db["attendance"],pd.DataFrame([[sid,cid,str(datetime.today().date()),status]],columns=SCHEMA["attendance"])])
                save(db)

# ================= STUDENT =================
    elif role=="Student":

        sid = clean(st.session_state.link)

        if ch=="My Results":
            df=db["results"].copy()
            df["student_id"]=df["student_id"].apply(clean)

            my=df[df["student_id"]==sid]
            st.dataframe(my)

            if my.empty:
                st.warning("No results")
            else:
                if st.button("Download Report Card"):
                    f=report_pdf(sid,"Student","Course",my)
                    with open(f,"rb") as file:
                        st.download_button("Download",file.read(),file_name=f)

        elif ch=="Attendance":
            df=db["attendance"].copy()
            df["student_id"]=df["student_id"].apply(clean)
            st.dataframe(df[df["student_id"]==sid])

        elif ch=="Certificate":
            if st.button("Download Certificate"):
                f=certificate_pdf(sid,"Student","Course")
                with open(f,"rb") as file:
                    st.download_button("Download",file.read(),file_name=f)

    if ch=="Logout":
        st.session_state.clear()
        st.session_state.login=False
        st.rerun()
