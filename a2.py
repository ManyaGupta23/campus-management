# ================= IMPORT =================
import streamlit as st
import pandas as pd
import os
from datetime import datetime
import qrcode

# PDF
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
    "rooms": ["room_id","capacity"],
    "schedule": ["class_id","faculty_id","room_id","time_slot","subject"],
    "attendance": ["student_id","class_id","date","status"],
    "users": ["username","password","role","student_id","faculty_id"],
    "results": ["student_id","subject","marks","grade"]
}

# ================= UI =================
st.markdown("""
<style>
.main {background:#f4f6fb;}
section[data-testid="stSidebar"] {background:#1e293b;}
section[data-testid="stSidebar"] * {color:white;}
.stButton>button {width:100%;height:45px;border-radius:10px;background:#2563eb;color:white;}
.card {background:white;padding:20px;border-radius:12px;box-shadow:0px 4px 10px rgba(0,0,0,0.1);}
</style>
""", unsafe_allow_html=True)

# ================= UTIL =================
def clean(x): return str(x).strip().replace(".0","")

def grade(m):
    m=int(m)
    return "A+" if m>=90 else "A" if m>=75 else "B" if m>=60 else "C" if m>=40 else "F"

# ================= FILE =================
def load():
    if not os.path.exists(FILE):
        with pd.ExcelWriter(FILE, engine="openpyxl") as w:
            for s, cols in SCHEMA.items():
                pd.DataFrame(columns=cols).to_excel(w, sheet_name=s, index=False)

    db={}
    xls=pd.ExcelFile(FILE)

    for s in SCHEMA:
        df=pd.read_excel(FILE, sheet_name=s) if s in xls.sheet_names else pd.DataFrame(columns=SCHEMA[s])
        df=df.astype(str)
        for col in df.columns:
            df[col]=df[col].apply(clean)
        db[s]=df

    return db

def save(db):
    with pd.ExcelWriter(FILE, engine="openpyxl", mode="w") as w:
        for s,df in db.items():
            df.to_excel(w, sheet_name=s, index=False)
    st.session_state.db = load()

# ================= QR =================
def qr(data,file):
    img=qrcode.make(data)
    img.save(file)
    return file

# ================= PDF REPORT =================
def report_pdf(sid,name,course,df):
    file=f"{sid}_report.pdf"
    doc=SimpleDocTemplate(file,pagesize=A4)
    styles=getSampleStyleSheet()
    content=[]

    if os.path.exists("logo.png"):
        content.append(Image("logo.png",1.3*inch,1.3*inch,hAlign='CENTER'))

    title=ParagraphStyle(name="t",fontSize=22,alignment=TA_CENTER,textColor=colors.HexColor("#C9A227"))
    content.append(Paragraph("ACADEMIC REPORT CARD",title))
    content.append(Spacer(1,20))

    content.append(Paragraph(f"Name: {name}",styles["Normal"]))
    content.append(Paragraph(f"Student ID: {sid}",styles["Normal"]))
    content.append(Paragraph(f"Course: {course}",styles["Normal"]))

    data=[["Subject","Marks","Grade"]]
    for _,r in df.iterrows():
        data.append([r["subject"],r["marks"],r["grade"]])

    t=Table(data)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#C9A227")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("GRID",(0,0),(-1,-1),1,colors.black)
    ]))
    content.append(t)

    content.append(Spacer(1,20))
    q=qr(f"{sid}-{name}",f"{sid}_qr.png")
    content.append(Image(q,1.5*inch,1.5*inch))

    if os.path.exists("signature.png"):
        content.append(Image("signature.png",2*inch,1*inch))

    doc.build(content)
    return file

# ================= CERTIFICATE =================
def certificate_pdf(sid,name,course):
    file=f"{sid}_certificate.pdf"
    doc=SimpleDocTemplate(file,pagesize=A4)
    styles=getSampleStyleSheet()
    content=[]

    if os.path.exists("logo.png"):
        content.append(Image("logo.png",1.5*inch,1.5*inch,hAlign='CENTER'))

    title=ParagraphStyle(name="t",fontSize=28,alignment=TA_CENTER,textColor=colors.HexColor("#C9A227"))
    content.append(Paragraph("CERTIFICATE OF COMPLETION",title))
    content.append(Spacer(1,30))

    content.append(Paragraph("This is to certify that",styles["Normal"]))
    content.append(Spacer(1,10))

    content.append(Paragraph(f"<b>{name}</b>",ParagraphStyle(name="n",fontSize=24,alignment=TA_CENTER)))
    content.append(Spacer(1,20))

    content.append(Paragraph(f"Student ID: {sid}",styles["Normal"]))
    content.append(Paragraph(f"Completed course {course}",styles["Normal"]))

    content.append(Spacer(1,30))
    content.append(Paragraph(f"Date: {datetime.today().date()}",styles["Normal"]))

    q=qr(f"CERT-{sid}",f"{sid}_certqr.png")
    content.append(Image(q,1.6*inch,1.6*inch,hAlign='CENTER'))

    if os.path.exists("signature.png"):
        content.append(Spacer(1,30))
        content.append(Image("signature.png",2*inch,1*inch,hAlign='CENTER'))

    doc.build(content)
    return file

# ================= SESSION =================
if "db" not in st.session_state:
    st.session_state.db = load()

if "login" not in st.session_state:
    st.session_state.login=False

# ================= LOGIN =================
if not st.session_state.login:

    c1,c2=st.columns([1.2,1])

    with c1:
        if os.path.exists("campus.jpg"):
            st.image("campus.jpg",use_container_width=True)

    with c2:
        st.title("🎓 Campus ERP Login")

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
                st.session_state.link = m.iloc[0]["student_id"] if r=="Student" else m.iloc[0]["faculty_id"]
                st.rerun()
            else:
                st.error("Invalid login")

# ================= MAIN =================
else:
    db=st.session_state.db
    role=st.session_state.role

    st.sidebar.title(role)

    if role=="Admin":
        menu=["Dashboard","Students","Faculty","Schedule","Logout"]
    elif role=="Faculty":
        menu=["My Schedule","Marks","Attendance","Logout"]
    else:
        menu=["My Results","Certificate","Logout"]

    ch=st.sidebar.radio("Menu",menu)

    # ADMIN
    if role=="Admin":

        if ch=="Dashboard":
            st.metric("Students",len(db["students"]))
            st.metric("Faculty",len(db["faculty"]))
            st.metric("Rooms",len(db["rooms"]))

        elif ch=="Students":
            st.dataframe(db["students"])
        elif ch=="Faculty":
            st.dataframe(db["faculty"])
        elif ch=="Schedule":
            st.dataframe(db["schedule"])

    # FACULTY
    elif role=="Faculty":

        if ch=="My Schedule":
            st.dataframe(db["schedule"][db["schedule"]["faculty_id"]==st.session_state.link])

        elif ch=="Marks":
            sid=st.text_input("Student ID")
            sub=st.text_input("Subject")
            m=st.number_input("Marks",0,100)

            if st.button("Save"):
                db["results"]=pd.concat([db["results"],
                pd.DataFrame([[sid,sub,m,grade(m)]],columns=SCHEMA["results"])])
                save(db)

        elif ch=="Attendance":
            sid=st.text_input("Student ID")
            cid=st.text_input("Class ID")
            status=st.selectbox("Status",["Present","Absent"])

            if st.button("Save"):
                db["attendance"]=pd.concat([db["attendance"],
                pd.DataFrame([[sid,cid,str(datetime.today().date()),status]],columns=SCHEMA["attendance"])])
                save(db)

    # STUDENT
    elif role=="Student":

        sid=str(st.session_state.link).replace(".0","")
        df=db["results"].copy()
        df["student_id"]=df["student_id"].astype(str).str.replace(".0","",regex=False)

        student_df=df[df["student_id"]==sid]
        st.dataframe(student_df)

        info=db["students"][db["students"]["student_id"]==sid]
        name=info.iloc[0]["name"] if not info.empty else "Student"
        course=info.iloc[0]["course"] if not info.empty else "Course"

        if ch=="My Results":
            if st.button("Download Report"):
                f=report_pdf(sid,name,course,student_df)
                with open(f,"rb") as file:
                    st.download_button("Download",file.read(),file_name=f"{sid}_report.pdf",mime="application/pdf")

        elif ch=="Certificate":
            if st.button("Download Certificate"):
                f=certificate_pdf(sid,name,course)
                with open(f,"rb") as file:
                    st.download_button("Download",file.read(),file_name=f"{sid}_certificate.pdf",mime="application/pdf")

    if ch=="Logout":
        st.session_state.clear()
        st.rerun()
