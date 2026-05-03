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
def make_qr(data,file):
    img=qrcode.make(data)
    img.save(file)
    return file

# ================= REPORT PDF =================
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

    table=Table(data)
    table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#C9A227")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("GRID",(0,0),(-1,-1),1,colors.black)
    ]))
    content.append(table)

    content.append(Spacer(1,20))
    q=make_qr(f"{sid}-{name}",f"{sid}_qr.png")
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
    content.append(Paragraph(f"<b>{name}</b>",ParagraphStyle(name="n",fontSize=28,alignment=TA_CENTER)))
    content.append(Spacer(1,20))

    content.append(Paragraph(f"Student ID: {sid}",styles["Normal"]))
    content.append(Paragraph(f"Completed course {course}",styles["Normal"]))

    content.append(Spacer(1,30))
    content.append(Paragraph(f"Date: {datetime.today().date()}",styles["Normal"]))

    q=make_qr(f"CERT-{sid}",f"{sid}_certqr.png")
    content.append(Image(q,1.6*inch,1.6*inch,hAlign='RIGHT'))

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
        menu=["Dashboard","Manage Students","Manage Faculty","Manage Schedule","Logout"]
    elif role=="Faculty":
        menu=["My Schedule","Students","Marks","Attendance","View Data","Logout"]
    else:
        menu=["My Results","My Attendance","Certificate","Logout"]

    ch=st.sidebar.radio("Menu",menu)

    # ===== ADMIN =====
    if role=="Admin":

        if ch=="Dashboard":
            st.metric("Students",len(db["students"]))
            st.metric("Faculty",len(db["faculty"]))
            st.metric("Schedule",len(db["schedule"]))

        elif ch=="Manage Students":
            st.dataframe(db["students"])
            sid=st.text_input("ID")
            name=st.text_input("Name")
            gender=st.text_input("Gender")
            cousre=st.text_input("Course")
            year=st.text_input("Year")
            section=st.text_input("Section")
            admission_date=st.text_input("Admission Date")
            attendance_percentage=st.text_input("Attendance Percentage")
            status=st.text_input("Status")
            email=st.text_input("Email")

            if st.button("Add"):
                db["students"]=pd.concat([db["students"],
                pd.DataFrame([[sid,name,course,email]],columns=SCHEMA["students"])])
                save(db)

            if st.button("Delete"):
                db["students"]=db["students"][db["students"]["student_id"]!=sid]
                save(db)

        elif ch=="Manage Faculty":
            st.dataframe(db["faculty"])
            fid=st.text_input("FID")
            name=st.text_input("Name")
            sub=st.text_input("Subject")

            if st.button("Add F"):
                db["faculty"]=pd.concat([db["faculty"],
                pd.DataFrame([[fid,name,sub]],columns=SCHEMA["faculty"])])
                save(db)

        elif ch=="Manage Schedule":
            st.dataframe(db["schedule"])

    # ===== FACULTY =====
    elif role=="Faculty":
       faculty_id = str(st.session_state.link)
    if ch == "My Schedule":
          st.subheader("📅 My Schedule")
          df = db["schedule"].copy()
        df["faculty_id"] = df["faculty_id"].astype(str)

        my_schedule = df[df["faculty_id"] == faculty_id]

        if my_schedule.empty:
            st.warning("No schedule assigned to you yet.")
        else:
            st.dataframe(my_schedule, use_container_width=True)
    

    if faculty_id is None:
        st.error("Faculty not logged in properly")
    else:
        df = db["schedule"].copy()

        # fix datatype mismatch
        df["faculty_id"] = df["faculty_id"].astype(str)
        faculty_id = str(faculty_id)

        my_schedule = df[df["faculty_id"] == faculty_id]

        if my_schedule.empty:
            st.warning("No schedule assigned to you yet.")
        else:
            st.dataframe(my_schedule, use_container_width=True)
             
        elif ch=="Students":
            st.dataframe(db["students"])

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
                pd.DataFrame([[sid,cid,str(datetime.today().date()),status]],
                columns=SCHEMA["attendance"])])
                save(db)

        elif ch=="View Data":
            st.write("Marks")
            st.dataframe(db["results"])
            st.write("Attendance")
            st.dataframe(db["attendance"])

    # ===== STUDENT =====
    elif role=="Student":

        sid=str(st.session_state.link).replace(".0","")

        if ch=="My Results":
            df=db["results"].copy()
            df["student_id"]=df["student_id"].astype(str).str.replace(".0","",regex=False)
            st.dataframe(df[df["student_id"]==sid])

        elif ch=="My Attendance":
            df=db["attendance"].copy()
            df["student_id"]=df["student_id"].astype(str).str.replace(".0","",regex=False)
            st.dataframe(df[df["student_id"]==sid])

        elif ch=="Certificate":
            info=db["students"][db["students"]["student_id"]==sid]
            name=info.iloc[0]["name"] if not info.empty else "Student"
            course=info.iloc[0]["course"] if not info.empty else "Course"

            if st.button("Download Certificate"):
                f=certificate_pdf(sid,name,course)
                with open(f,"rb") as file:
                    st.download_button("Download",file.read(),file_name=f"{sid}_certificate.pdf",mime="application/pdf")

    if ch=="Logout":
        st.session_state.clear()
        st.rerun()
