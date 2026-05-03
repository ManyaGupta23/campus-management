# ================= IMPORT =================
import streamlit as st
import pandas as pd
import os
from datetime import datetime
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4

# ================= CONFIG =================
st.set_page_config(page_title="Campus ERP", layout="wide")

FILE = "campus_flow_template.xlsx"

SCHEMA = {
    "students": ["student_id","name","course","email"],
    "faculty": ["faculty_id","name","subject"],
    "schedule": ["class_id","faculty_id","room_id","time_slot","subject"],
    "users": ["username","password","role","student_id","faculty_id"],
    "results": ["student_id","subject","marks","grade"]
}

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
        db[s]=df

    return db

def save(db):
    with pd.ExcelWriter(FILE, engine="openpyxl", mode="w") as w:
        for s,df in db.items():
            df.to_excel(w, sheet_name=s, index=False)
    st.session_state.db = load()

# ================= PDF =================
def report_pdf(sid, df):
    file=f"{sid}_report.pdf"
    doc=SimpleDocTemplate(file,pagesize=A4)

    data=[["Subject","Marks","Grade"]]
    for _,r in df.iterrows():
        data.append([r["subject"],r["marks"],r["grade"]])

    table=Table(data)
    table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),1,colors.black)
    ]))

    doc.build([table])
    return file

# ================= SESSION =================
if "db" not in st.session_state:
    st.session_state.db = load()

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
            st.session_state.link=m.iloc[0]["student_id"]
            st.rerun()

# ================= MAIN =================
else:
    db=st.session_state.db
    role=st.session_state.role

    if role=="Admin":
        menu=["Students","Faculty","Schedule","Logout"]
    elif role=="Faculty":
        menu=["Marks","Logout"]
    else:
        menu=["Results","Logout"]

    ch=st.sidebar.radio("Menu",menu)

    # ================= ADMIN =================
    if role=="Admin":

        # ===== STUDENTS =====
        if ch=="Students":
            st.subheader("Students")
            st.dataframe(db["students"])

            sid=st.text_input("ID")
            name=st.text_input("Name")
            course=st.text_input("Course")
            email=st.text_input("Email")

            col1,col2,col3=st.columns(3)

            # ADD
            if col1.button("Add"):
                new=pd.DataFrame([[sid,name,course,email]],columns=SCHEMA["students"])
                db["students"]=pd.concat([db["students"],new],ignore_index=True)
                save(db)
                st.success("Added")

            # UPDATE
            if col2.button("Update"):
                db["students"].loc[db["students"]["student_id"]==sid,["name","course","email"]] = [name,course,email]
                save(db)
                st.success("Updated")

            # DELETE
            if col3.button("Delete"):
                db["students"]=db["students"][db["students"]["student_id"]!=sid]
                save(db)
                st.success("Deleted")

        # ===== FACULTY =====
        elif ch=="Faculty":
            st.subheader("Faculty")
            st.dataframe(db["faculty"])

            fid=st.text_input("FID")
            name=st.text_input("Name")
            sub=st.text_input("Subject")

            col1,col2,col3=st.columns(3)

            if col1.button("Add F"):
                db["faculty"]=pd.concat([db["faculty"],
                pd.DataFrame([[fid,name,sub]],columns=SCHEMA["faculty"])])
                save(db)

            if col2.button("Update F"):
                db["faculty"].loc[db["faculty"]["faculty_id"]==fid,["name","subject"]] = [name,sub]
                save(db)

            if col3.button("Delete F"):
                db["faculty"]=db["faculty"][db["faculty"]["faculty_id"]!=fid]
                save(db)

        # ===== SCHEDULE =====
        elif ch=="Schedule":
            st.subheader("Schedule")
            st.dataframe(db["schedule"])

            cid=st.text_input("CID")
            fid=st.text_input("FID")
            room=st.text_input("Room")
            time=st.text_input("Time")
            sub=st.text_input("Subject")

            col1,col2,col3=st.columns(3)

            if col1.button("Add S"):
                db["schedule"]=pd.concat([db["schedule"],
                pd.DataFrame([[cid,fid,room,time,sub]],columns=SCHEMA["schedule"])])
                save(db)

            if col2.button("Update S"):
                db["schedule"].loc[db["schedule"]["class_id"]==cid,
                ["faculty_id","room_id","time_slot","subject"]] = [fid,room,time,sub]
                save(db)

            if col3.button("Delete S"):
                db["schedule"]=db["schedule"][db["schedule"]["class_id"]!=cid]
                save(db)

        elif ch=="Logout":
            st.session_state.clear()
            st.rerun()

    # ================= FACULTY =================
    elif role=="Faculty":

        if ch=="Marks":
            sid=st.text_input("Student ID")
            sub=st.text_input("Subject")
            m=st.number_input("Marks",0,100)

            if st.button("Save"):
                g="A+" if m>=90 else "A" if m>=75 else "B" if m>=60 else "C" if m>=40 else "F"
                db["results"]=pd.concat([db["results"],
                pd.DataFrame([[sid,sub,m,g]],columns=SCHEMA["results"])])
                save(db)
                st.success("Saved")

        elif ch=="Logout":
            st.session_state.clear()
            st.rerun()

    # ================= STUDENT =================
    elif role=="Student":

        sid=st.session_state.link
        df=db["results"][db["results"]["student_id"]==sid]

        st.dataframe(df)

        if st.button("Download Report"):
            f=report_pdf(sid,df)

            with open(f,"rb") as file:
                st.download_button(
                    "Download PDF",
                    data=file.read(),
                    file_name=f"{sid}_report.pdf",
                    mime="application/pdf"
                )

        if ch=="Logout":
            st.session_state.clear()
            st.rerun()
