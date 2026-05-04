# ================= CAMPUS ERP (FINAL PREMIUM VERSION) =================

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

# ================= FILE HANDLING =================
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

# ================= PREMIUM REPORT =================
def report_pdf(sid,name,course,df):
    file = f"{sid}_report.pdf"
    doc = SimpleDocTemplate(file, pagesize=A4)
    styles = getSampleStyleSheet()
    content = []

    if os.path.exists("logo.png"):
        content.append(Image("logo.png",3*inch,1.5*inch,hAlign='CENTER'))

    title = ParagraphStyle(name="t", fontSize=26, alignment=TA_CENTER, textColor=colors.HexColor("#1e40af"))
    content.append(Paragraph("ACADEMIC REPORT CARD", title))
    content.append(Spacer(1,20))

    qr = make_qr(f"{sid}-{name}", f"{sid}_qr.png")

    info = [
        [Paragraph(f"<b>Name:</b> {name}", styles["Normal"])],
        [Paragraph(f"<b>ID:</b> {sid}", styles["Normal"])],
        [Paragraph(f"<b>Course:</b> {course}", styles["Normal"])],
    ]

    content.append(Table([[Image(qr,2*inch,2*inch), info]]))
    content.append(Spacer(1,20))

    data = [["Subject","Marks","Grade"]]
    for _,r in df.iterrows():
        data.append([r["subject"],r["marks"],r["grade"]])

    table = Table(data)
    table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#2563eb")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("GRID",(0,0),(-1,-1),1,colors.black),
    ]))

    content.append(table)

    avg = df["marks"].astype(float).mean() if not df.empty else 0
    remark = "Excellent" if avg>=80 else "Good" if avg>=60 else "Needs Improvement"

    content.append(Spacer(1,15))
    content.append(Paragraph(f"<b>Average:</b> {round(avg,2)}", styles["Normal"]))
    content.append(Paragraph(f"<b>Remark:</b> {remark}", styles["Normal"]))

    if os.path.exists("signature.png"):
        content.append(Spacer(1,30))
        content.append(Image("signature.png",2*inch,1*inch,hAlign='RIGHT'))

    doc.build(content)
    return file

# ================= PREMIUM CERTIFICATE =================
def certificate_pdf(sid,name,course):
    file = f"{sid}_certificate.pdf"
    doc = SimpleDocTemplate(file, pagesize=A4)
    styles = getSampleStyleSheet()
    content = []

    if os.path.exists("logo.png"):
        content.append(Image("logo.png",3*inch,1.5*inch,hAlign='CENTER'))

    title = ParagraphStyle(name="t", fontSize=30, alignment=TA_CENTER, textColor=colors.HexColor("#2563eb"))
    content.append(Paragraph("CERTIFICATE OF COMPLETION", title))
    content.append(Spacer(1,30))

    para = ParagraphStyle(name="p", alignment=TA_CENTER, fontSize=16, leading=24)

    content.append(Paragraph("This is to certify that", para))
    content.append(Spacer(1,10))
    content.append(Paragraph(f"<b style='font-size:22px'>{name}</b>", para))
    content.append(Spacer(1,10))
    content.append(Paragraph("has successfully completed", para))
    content.append(Paragraph(f"<b>{course}</b>", para))

    content.append(Spacer(1,25))

    qr = make_qr(f"CERT-{sid}", f"{sid}_cert.png")

    content.append(Table([[Image(qr,2*inch,2*inch), Paragraph(f"Student ID: {sid}", styles["Normal"])]]))

    if os.path.exists("signature.png"):
        content.append(Spacer(1,40))
        content.append(Image("signature.png",2.5*inch,1.2*inch,hAlign='CENTER'))

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
        if os.path.exists("campus.jpg"):
            st.image("campus.jpg", use_container_width=True)

    with col2:
        st.title("🎓 Login")

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
                st.session_state.link = str(
                    m.iloc[0]["student_id"] if r=="Student" else m.iloc[0]["faculty_id"]
                ).strip()
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
        menu=["My Results","My Attendance","Certificate","Logout"]

    ch = st.sidebar.radio("Menu", menu)

# ================= ADMIN =================
    if role=="Admin":

        if ch=="Students":
            st.dataframe(db["students"])

            sid=st.text_input("Student ID")
            name=st.text_input("Name")

            if st.button("Add Student"):
                db["students"]=pd.concat([db["students"],pd.DataFrame([[sid,name,"","","","","","","",""]],columns=SCHEMA["students"])])
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
                db["faculty"]=pd.concat([db["faculty"],pd.DataFrame([[fid,name,""]],columns=SCHEMA["faculty"])])
                db["users"]=pd.concat([db["users"],pd.DataFrame([[fid,"1234","Faculty","",fid]],columns=SCHEMA["users"])])
                save(db)

            del_id=st.text_input("Delete Faculty ID")
            if st.button("Delete Faculty"):
                db["faculty"]=db["faculty"][db["faculty"]["faculty_id"]!=del_id]
                save(db)

        elif ch=="Schedule":
            st.dataframe(db["schedule"])

            cid=st.text_input("Class ID")
            fid=st.text_input("Faculty ID")
            sub=st.text_input("Subject")

            if st.button("Add Schedule"):
                db["schedule"]=pd.concat([db["schedule"],pd.DataFrame([[cid,fid,"","",sub]],columns=SCHEMA["schedule"])])
                save(db)

# ================= FACULTY =================
    elif role=="Faculty":

        fid=str(st.session_state.link)

        if ch=="My Schedule":
            df=db["schedule"]
            st.dataframe(df[df["faculty_id"]==fid])

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

        sid=str(st.session_state.link)

        if ch=="My Results":
            df=db["results"]
            my=df[df["student_id"]==sid]
            st.dataframe(my)

            if not my.empty:
                if st.button("Download Report Card"):
                    f=report_pdf(sid,"Student","Course",my)
                    st.download_button("Download",open(f,"rb").read(),file_name=f)

        elif ch=="My Attendance":
            st.dataframe(db["attendance"][db["attendance"]["student_id"]==sid])

        elif ch=="Certificate":
            if st.button("Download Certificate"):
                f=certificate_pdf(sid,"Student","Course")
                st.download_button("Download",open(f,"rb").read(),file_name=f)

    if ch=="Logout":
        st.session_state.clear()
        st.session_state.login = False
        st.rerun()
