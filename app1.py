# ================= IMPORT =================
import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import smtplib
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

# ================= UTIL =================
def clean(x):
    return str(x).strip().replace(".0","")

# ================= FILE =================
def load():
    if not os.path.exists(FILE):
        with pd.ExcelWriter(FILE, engine="openpyxl") as w:
            for s, cols in SCHEMA.items():
                pd.DataFrame(columns=cols).to_excel(w, sheet_name=s, index=False)

    db={}
    xls = pd.ExcelFile(FILE)

    for s in SCHEMA:
        df = pd.read_excel(FILE, sheet_name=s) if s in xls.sheet_names else pd.DataFrame(columns=SCHEMA[s])
        df = df.astype(str).applymap(clean)
        db[s]=df

    return db

def save(db):
    with pd.ExcelWriter(FILE, engine="openpyxl", mode="w") as w:
        for s,df in db.items():
            df.to_excel(w, sheet_name=s, index=False)
    st.session_state.db = load()

# ================= GRADE =================
def grade(m):
    m=int(m)
    return "A+" if m>=90 else "A" if m>=75 else "B" if m>=60 else "C" if m>=40 else "F"

# ================= EMAIL =================
def send_email(to, sub, body):
    try:
        s = smtplib.SMTP("smtp.gmail.com",587)
        s.starttls()
        s.login("your_email@gmail.com","app_password")
        s.sendmail("your_email@gmail.com",to,f"Subject:{sub}\n\n{body}")
        s.quit()
    except:
        pass

# ================= QR =================
def generate_qr(data, file):
    img = qrcode.make(data)
    img.save(file)
    return file

# ================= PDF REPORT =================
def report_pdf(sid,name,course,df):
    file=f"{sid}_report.pdf"
    doc=SimpleDocTemplate(file,pagesize=A4)
    styles=getSampleStyleSheet()

    content=[]

    qr_file=generate_qr(f"{name}-{sid}-{course}",f"{sid}_qr.png")

    if os.path.exists("logo.png"):
        content.append(Image("logo.png",1.2*inch,1.2*inch,hAlign='CENTER'))

    title=ParagraphStyle(name="t",fontSize=24,alignment=TA_CENTER,textColor=colors.HexColor("#C9A227"))
    content.append(Paragraph("ACADEMIC REPORT CARD",title))
    content.append(Spacer(1,20))

    content.append(Paragraph(f"Name: {name}",styles["Normal"]))
    content.append(Paragraph(f"Student ID: {sid}",styles["Normal"]))
    content.append(Paragraph(f"Course: {course}",styles["Normal"]))
    content.append(Spacer(1,10))

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

    avg=df["marks"].astype(int).mean()
    content.append(Spacer(1,10))
    content.append(Paragraph(f"Average: {round(avg,2)}",styles["Normal"]))

    content.append(Spacer(1,20))
    content.append(Paragraph("Scan to Verify",styles["Normal"]))
    content.append(Image(qr_file,1.5*inch,1.5*inch))

    if os.path.exists("signature.png"):
        content.append(Spacer(1,30))
        content.append(Image("signature.png",2*inch,1*inch))

    def border(c,d):
        c.setStrokeColor(colors.HexColor("#C9A227"))
        c.setLineWidth(4)
        c.rect(20,20,555,800)

    doc.build(content,onFirstPage=border)
    return file

# ================= PDF CERTIFICATE =================
def certificate_pdf(sid,name,course):
    file=f"{sid}_certificate.pdf"
    doc=SimpleDocTemplate(file,pagesize=A4)
    styles=getSampleStyleSheet()

    content=[]

    qr_file=generate_qr(f"Certificate-{name}-{sid}",f"{sid}_certqr.png")

    if os.path.exists("logo.png"):
        content.append(Image("logo.png",1.5*inch,1.5*inch,hAlign='CENTER'))

    title=ParagraphStyle(name="t",fontSize=28,alignment=TA_CENTER,textColor=colors.HexColor("#C9A227"))
    content.append(Paragraph("CERTIFICATE OF COMPLETION",title))
    content.append(Spacer(1,30))

    content.append(Paragraph(f"This is to certify {name}",styles["Normal"]))
    content.append(Paragraph(f"Student ID: {sid}",styles["Normal"]))
    content.append(Paragraph(f"Completed course {course}",styles["Normal"]))

    content.append(Spacer(1,30))
    content.append(Paragraph("Scan to Verify",styles["Normal"]))
    content.append(Image(qr_file,1.6*inch,1.6*inch))

    if os.path.exists("signature.png"):
        content.append(Spacer(1,40))
        content.append(Image("signature.png",2*inch,1*inch))

    def border(c,d):
        c.setStrokeColor(colors.HexColor("#C9A227"))
        c.setLineWidth(4)
        c.rect(20,20,555,800)

    doc.build(content,onFirstPage=border)
    return file

# ================= CONFLICT =================
def conflicts(df):
    res=[]
    for i in range(len(df)):
        for j in range(i+1,len(df)):
            a,b=df.iloc[i],df.iloc[j]
            if a["time_slot"]==b["time_slot"]:
                if a["room_id"]==b["room_id"]:
                    res.append("Room conflict")
                if a["faculty_id"]==b["faculty_id"]:
                    res.append("Faculty conflict")
    return res

# ================= SESSION =================
if "db" not in st.session_state:
    st.session_state.db=load()

if "login" not in st.session_state:
    st.session_state.login=False

# ================= LOGIN =================
if not st.session_state.login:
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

# ================= MAIN =================
else:
    db=st.session_state.db
    role=st.session_state.role

    if role=="Admin":
        menu=["Dashboard","Students","Faculty","Schedule","Logout"]
    elif role=="Faculty":
        menu=["Students","My Schedule","Marks","Attendance","Performance","Logout"]
    else:
        menu=["My Results","Certificate","Logout"]

    ch=st.sidebar.radio("Menu",menu)

    # ===== ADMIN =====
    if ch=="Dashboard":
        st.metric("Students",len(db["students"]))
        st.metric("Faculty",len(db["faculty"]))
        st.metric("Rooms",len(db["rooms"]))

        c=conflicts(db["schedule"])
        if c: st.error(c)

    elif ch=="Students":
        st.dataframe(db["students"])
        sid=st.text_input("ID")
        name=st.text_input("Name")
        course=st.text_input("Course")
        email=st.text_input("Email")

        if st.button("Add"):
            db["students"]=pd.concat([db["students"],
            pd.DataFrame([[sid,name,course,email]],columns=SCHEMA["students"])])
            save(db)

    elif ch=="Faculty":
        st.dataframe(db["faculty"])
        fid=st.text_input("FID")
        name=st.text_input("Name")
        sub=st.text_input("Subject")

        if st.button("Add Faculty"):
            db["faculty"]=pd.concat([db["faculty"],
            pd.DataFrame([[fid,name,sub]],columns=SCHEMA["faculty"])])
            save(db)

    elif ch=="Schedule":
        st.dataframe(db["schedule"])
        cid=st.text_input("CID")
        fid=st.text_input("FID")
        room=st.text_input("Room")
        time=st.text_input("Time")
        sub=st.text_input("Subject")

        if st.button("Add Schedule"):
            db["schedule"]=pd.concat([db["schedule"],
            pd.DataFrame([[cid,fid,room,time,sub]],columns=SCHEMA["schedule"])])
            save(db)

    # ===== FACULTY =====
    elif ch=="Students":
        st.dataframe(db["students"])

    elif ch=="My Schedule":
        fid=st.session_state.link
        st.dataframe(db["schedule"][db["schedule"]["faculty_id"]==fid])

    elif ch=="Marks":
        sid=st.text_input("SID")
        sub=st.text_input("Subject")
        m=st.number_input("Marks",0,100)

        if st.button("Save"):
            g=grade(m)
            db["results"]=pd.concat([db["results"],
            pd.DataFrame([[sid,sub,m,g]],columns=SCHEMA["results"])])
            save(db)

    elif ch=="Attendance":
        for _,r in db["students"].iterrows():
            status=st.selectbox(r["name"],["Present","Absent"],key=r["student_id"])
            if st.button(f"Save {r['student_id']}"):
                db["attendance"]=pd.concat([db["attendance"],
                pd.DataFrame([[r["student_id"],"C1",str(datetime.today().date()),status]],
                columns=SCHEMA["attendance"])])
                save(db)

    elif ch=="Performance":
        st.plotly_chart(px.bar(db["results"],x="student_id",y="marks"))

    # ===== STUDENT =====
    elif ch=="My Results":
        sid=st.session_state.link
        df=db["results"][db["results"]["student_id"]==sid]

        st.dataframe(df)

        info=db["students"][db["students"]["student_id"]==sid]
        name=info.iloc[0]["name"] if not info.empty else "Student"
        course=info.iloc[0]["course"] if not info.empty else "Course"

        if st.button("Download Report"):
            f=report_pdf(sid,name,course,df)
            with open(f,"rb") as file:
                st.download_button("Download PDF",file)

    elif ch=="Certificate":
        sid=st.session_state.link
        info=db["students"][db["students"]["student_id"]==sid]
        name=info.iloc[0]["name"] if not info.empty else "Student"
        course=info.iloc[0]["course"] if not info.empty else "Course"

        if st.button("Download Certificate"):
            f=certificate_pdf(sid,name,course)
            with open(f,"rb") as file:
                st.download_button("Download Certificate",file)

    elif ch=="Logout":
        st.session_state.clear()
        st.rerun()
