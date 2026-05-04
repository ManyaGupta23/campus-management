# ================= CAMPUS ERP FINAL VERSION =================

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
    "faculty": ["faculty_id","name","subject"],
    "schedule": ["class_id","faculty_id","subject","time"],
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
            for s,c in SCHEMA.items():
                pd.DataFrame(columns=c).to_excel(w,s,index=False)

    db={}
    xls=pd.ExcelFile(FILE)

    for s in SCHEMA:
        df = pd.read_excel(FILE, sheet_name=s) if s in xls.sheet_names else pd.DataFrame(columns=SCHEMA[s])
        df = df.astype(str)
        for col in df.columns:
            df[col] = df[col].apply(clean)
        db[s]=df

    return db

def save(db):
    with pd.ExcelWriter(FILE) as w:
        for s,df in db.items():
            df.to_excel(w,s,index=False)
    st.session_state.db = load()

# ================= QR =================
def make_qr(data,file):
    img=qrcode.make(data)
    img.save(file)
    return file

# ================= PDF =================
def report_pdf(sid,name,course,df):
    file=f"{sid}_report.pdf"
    doc=SimpleDocTemplate(file,pagesize=A4)
    styles=getSampleStyleSheet()
    content=[]

    title=ParagraphStyle(name="t",fontSize=26,alignment=TA_CENTER,textColor=colors.darkblue)

    content.append(Paragraph("ACADEMIC REPORT CARD",title))
    content.append(Spacer(1,20))

    qr = make_qr(f"{sid}-{name}",f"{sid}.png")

    content.append(Table([[Image(qr,2*inch,2*inch),
                          Paragraph(f"<b>{name}</b><br/>ID: {sid}<br/>Course: {course}",styles["Normal"])]]))

    data=[["Subject","Marks","Grade"]]
    for _,r in df.iterrows():
        data.append([r["subject"],r["marks"],r["grade"]])

    table=Table(data)
    table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.blue),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("GRID",(0,0),(-1,-1),1,colors.black)
    ]))

    content.append(Spacer(1,20))
    content.append(table)

    doc.build(content)
    return file

def certificate_pdf(sid,name,course):
    file=f"{sid}_cert.pdf"
    doc=SimpleDocTemplate(file,pagesize=A4)
    styles=getSampleStyleSheet()
    content=[]

    title=ParagraphStyle(name="t",fontSize=28,alignment=TA_CENTER,textColor=colors.green)

    content.append(Paragraph("CERTIFICATE OF COMPLETION",title))
    content.append(Spacer(1,30))

    content.append(Paragraph(f"This certifies <b>{name}</b>",styles["Normal"]))
    content.append(Paragraph(f"Completed course: <b>{course}</b>",styles["Normal"]))

    qr = make_qr(f"CERT-{sid}",f"{sid}_c.png")

    content.append(Spacer(1,20))
    content.append(Image(qr,2*inch,2*inch))

    doc.build(content)
    return file

# ================= SESSION =================
if "db" not in st.session_state:
    st.session_state.db = load()

if "login" not in st.session_state:
    st.session_state.login=False

# ================= LOGIN =================
if not st.session_state.login:

    col1,col2=st.columns([1.5,1])

    with col1:
        if os.path.exists("campus.jpg"):
            st.image("campus.jpg",use_container_width=True)
        else:
            st.image("https://images.unsplash.com/photo-1523050854058-8df90110c9f1",use_container_width=True)

    with col2:
        st.title("Campus ERP Login")

        u=st.text_input("Username")
        p=st.text_input("Password",type="password")
        r=st.selectbox("Role",["Admin","Faculty","Student"])

        if st.button("Login"):
            users=st.session_state.db["users"]

            if u=="admin" and p=="admin123":
                st.session_state.login=True
                st.session_state.role="Admin"
                st.rerun()

            m=users[(users.username==u)&(users.password==p)&(users.role==r)]

            if not m.empty:
                st.session_state.login=True
                st.session_state.role=r
                st.session_state.link=m.iloc[0]["student_id"] if r=="Student" else m.iloc[0]["faculty_id"]
                st.rerun()
            else:
                st.error("Invalid Login")

# ================= MAIN =================
else:
    db=st.session_state.db
    role=st.session_state.role

    st.sidebar.title(role)

    if role=="Admin":
        menu=["Students","Faculty","Schedule","Logout"]
    elif role=="Faculty":
        menu=["My Schedule","Marks","Attendance","Logout"]
    else:
        menu=["My Results","Report Card","Certificate","Logout"]

    ch=st.sidebar.radio("Menu",menu)

    # ADMIN
    if role=="Admin":

        if ch=="Students":
            st.dataframe(db["students"])

            sid=st.text_input("ID")
            name=st.text_input("Name")
            course=st.text_input("Course")

            if st.button("Add"):
                db["students"]=pd.concat([db["students"],pd.DataFrame([[sid,name,course]],columns=SCHEMA["students"])])
                db["users"]=pd.concat([db["users"],pd.DataFrame([[sid,"1234","Student",sid,""]],columns=SCHEMA["users"])])
                save(db)

        elif ch=="Faculty":
            st.dataframe(db["faculty"])

            fid=st.text_input("FID")
            name=st.text_input("Name")
            sub=st.text_input("Subject")

            if st.button("Add Faculty"):
                db["faculty"]=pd.concat([db["faculty"],pd.DataFrame([[fid,name,sub]],columns=SCHEMA["faculty"])])
                db["users"]=pd.concat([db["users"],pd.DataFrame([[fid,"1234","Faculty","",fid]],columns=SCHEMA["users"])])
                save(db)

        elif ch=="Schedule":
            st.dataframe(db["schedule"])

            cid=st.text_input("Class ID")
            fid=st.text_input("Faculty ID")
            sub=st.text_input("Subject")
            time=st.text_input("Time")

            if st.button("Add Schedule"):
                db["schedule"]=pd.concat([db["schedule"],pd.DataFrame([[cid,fid,sub,time]],columns=SCHEMA["schedule"])])
                save(db)

    # FACULTY
    elif role=="Faculty":

        fid=str(st.session_state.link).strip()

        if ch=="My Schedule":
            df=db["schedule"]
            df["faculty_id"]=df["faculty_id"].astype(str).str.strip()
            my=df[df["faculty_id"]==fid]

            if my.empty:
                st.warning("No Schedule Assigned")
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

            if st.button("Save"):
                db["attendance"]=pd.concat([db["attendance"],pd.DataFrame([[sid,cid,str(datetime.today().date()),status]],columns=SCHEMA["attendance"])])
                save(db)

    # STUDENT
    elif role=="Student":

        sid=str(st.session_state.link).strip()

        if ch=="My Results":
            df=db["results"]
            df["student_id"]=df["student_id"].astype(str).str.strip()
            my=df[df["student_id"]==sid]

            st.dataframe(my)

        elif ch=="Report Card":
            df=db["results"]
            df["student_id"]=df["student_id"].astype(str).str.strip()
            my=df[df["student_id"]==sid]

            if my.empty:
                st.warning("No Result Found")
            else:
                info=db["students"][db["students"]["student_id"]==sid]
                name=info.iloc[0]["name"]
                course=info.iloc[0]["course"]

                if st.button("Download Report"):
                    f=report_pdf(sid,name,course,my)
                    with open(f,"rb") as file:
                        st.download_button("Download",file.read(),file_name=f)

        elif ch=="Certificate":
            info=db["students"][db["students"]["student_id"]==sid]

            if not info.empty:
                name=info.iloc[0]["name"]
                course=info.iloc[0]["course"]

                if st.button("Download Certificate"):
                    f=certificate_pdf(sid,name,course)
                    with open(f,"rb") as file:
                        st.download_button("Download",file.read(),file_name=f)

    if ch=="Logout":
        st.session_state.clear()
        st.rerun()
