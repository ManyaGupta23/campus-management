import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
from reportlab.pdfgen import canvas

EXCEL_FILE = "campus_flow_template.xlsx"

st.set_page_config(page_title="Campus Flow System", layout="wide")

# ------------------------
# LOAD / SAVE
# ------------------------
@st.cache_data
def load_sheet(sheet):
    return pd.read_excel(EXCEL_FILE, sheet_name=sheet)

def save_sheet(sheet, df):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=sheet, index=False)

# ------------------------
# GRADE FUNCTION
# ------------------------
def calculate_grade(m):
    if m >= 80:
        return "A"
    elif m >= 60:
        return "B"
    elif m >= 40:
        return "C"
    else:
        return "Fail"

# ------------------------
# SESSION
# ------------------------
if "login" not in st.session_state:
    st.session_state.login = False

# ------------------------
# LOGIN
# ------------------------
if not st.session_state.login:
    st.title("🔐 Login")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    role = st.selectbox("Role", ["Admin", "Faculty", "Student"])

    users = load_sheet("users")

    if st.button("Login"):
        user = users[
            (users["username"] == username) &
            (users["password"] == password) &
            (users["role"] == role)
        ]

        if not user.empty:
            st.session_state.login = True
            st.session_state.role = role
            st.session_state.user = username

            if role == "Student":
                st.session_state.student_id = str(user.iloc[0]["student_id"])

            if role == "Faculty":
                st.session_state.faculty_id = str(user.iloc[0]["faculty_id"])

            st.rerun()
        else:
            st.error("Invalid Credentials")

# ------------------------
# MAIN SYSTEM
# ------------------------
else:
    role = st.session_state.role

    st.sidebar.title("🏫 Campus Flow")
    st.sidebar.write(st.session_state.user)

    # ROLE BASED MENU
    if role == "Admin":
        menu = ["Dashboard","Students","Schedule","Campus Flow","Conflict Detection"]

    elif role == "Faculty":
        menu = ["My Classes","Mark Attendance","Enter Marks","My Analytics"]

    else:
        menu = ["My Data","Performance Analytics","Certificate","Report Card"]

    choice = st.sidebar.radio("Menu", menu)

    # ---------------- ADMIN ----------------
    if role == "Admin":

        if choice == "Dashboard":
            st.title("📊 Dashboard")

            students = load_sheet("students")
            attendance = load_sheet("attendance")

            c1, c2 = st.columns(2)
            c1.metric("Students", len(students))

            if not attendance.empty:
                present = attendance[attendance["status"]=="Present"].shape[0]
                total = len(attendance)
                c2.metric("Attendance %", round((present/total)*100,2))

        elif choice == "Students":
            st.title("👨‍🎓 Students")

            df = load_sheet("students")

            with st.form("add"):
                sid = st.text_input("ID")
                name = st.text_input("Name")

                if st.form_submit_button("Add"):
                    new = pd.DataFrame([[sid,name,"F","BCA",1,"A",str(datetime.now().date()),0,"Active",""]],
                                       columns=df.columns)
                    df = pd.concat([df,new],ignore_index=True)
                    save_sheet("students",df)

            st.dataframe(df)

        elif choice == "Schedule":
            st.title("📚 Schedule")

            df = load_sheet("schedule")

            with st.form("sch"):
                cid = st.text_input("Class ID")
                fid = st.text_input("Faculty ID")
                room = st.text_input("Room")
                time = st.text_input("Time")

                if st.form_submit_button("Add"):
                    new = pd.DataFrame([[cid,fid,room,time,"Subject"]],columns=df.columns)
                    df = pd.concat([df,new],ignore_index=True)
                    save_sheet("schedule",df)

            st.dataframe(df)

        elif choice == "Campus Flow":
            st.title("🚶 Campus Flow")

            att = load_sheet("attendance")
            logs = load_sheet("logs")
            sch = load_sheet("schedule")

            if not att.empty:
                att["date"] = pd.to_datetime(att["date"])
                trend = att.groupby("date").size().reset_index(name="count")
                st.plotly_chart(px.line(trend,x="date",y="count"))

            if not logs.empty:
                logs["entry_time"] = pd.to_datetime(logs["entry_time"],errors="coerce")
                logs["hour"] = logs["entry_time"].dt.hour
                st.plotly_chart(px.histogram(logs,x="hour"))

            if not sch.empty:
                usage = sch["room_id"].value_counts().reset_index()
                usage.columns=["room","count"]
                st.plotly_chart(px.pie(usage,names="room",values="count"))

        elif choice == "Conflict Detection":
            st.title("⚠️ Conflict Detection")

            df = load_sheet("schedule")
            con = df[df.duplicated(subset=["room_id","time_slot"],keep=False)]

            if not con.empty:
                st.error("Conflicts Found")
                st.dataframe(con)
            else:
                st.success("No Conflict")

    # ---------------- FACULTY ----------------
    elif role == "Faculty":

        if choice == "My Classes":
            st.title("📚 My Classes")

            sch = load_sheet("schedule")
            my = sch[sch["faculty_id"].astype(str)==st.session_state.faculty_id]

            st.dataframe(my)

        elif choice == "Mark Attendance":
            st.title("📅 Attendance")

            att = load_sheet("attendance")

            sid = st.text_input("Student ID")
            cid = st.text_input("Class ID")
            status = st.selectbox("Status",["Present","Absent"])

            if st.button("Submit"):
                new = pd.DataFrame([[sid,cid,str(datetime.now().date()),status]],
                                   columns=att.columns)
                att = pd.concat([att,new],ignore_index=True)
                save_sheet("attendance",att)

        elif choice == "Enter Marks":
            st.title("📊 Enter Marks")

            res = load_sheet("results")

            sid = st.text_input("Student ID")
            subj = st.text_input("Subject")
            marks = st.number_input("Marks",0,100)

            if st.button("Save"):
                grade = calculate_grade(marks)
                new = pd.DataFrame([[sid,subj,marks,grade]],columns=res.columns)
                res = pd.concat([res,new],ignore_index=True)
                save_sheet("results",res)

        elif choice == "My Analytics":
            st.title("📈 Analytics")

            res = load_sheet("results")

            if not res.empty:
                st.plotly_chart(px.bar(res,x="subject",y="marks"))
                st.metric("Avg Marks",round(res["marks"].mean(),2))

    # ---------------- STUDENT ----------------
    else:

        if choice == "My Data":
            st.title("👨‍🎓 My Data")

            stu = load_sheet("students")
            stu = stu[stu["student_id"].astype(str)==st.session_state.student_id]

            st.dataframe(stu)

        elif choice == "Performance Analytics":
            st.title("📈 Performance")

            res = load_sheet("results")
            res = res[res["student_id"].astype(str)==st.session_state.student_id]

            if not res.empty:
                st.plotly_chart(px.bar(res,x="subject",y="marks"))
                st.metric("Avg",round(res["marks"].mean(),2))

        elif choice == "Certificate":
            st.title("🏆 Certificate")

            stu = load_sheet("students")
            s = stu[stu["student_id"].astype(str)==st.session_state.student_id]

            if not s.empty:
                name = s.iloc[0]["name"]

                if st.button("Generate"):
                    file="cert.pdf"
                    c=canvas.Canvas(file)
                    c.drawString(100,700,f"Certificate for {name}")
                    c.save()

                    with open(file,"rb") as f:
                        st.download_button("Download",f,file_name=file)

        elif choice == "Report Card":
            st.title("📄 Report Card")

            res = load_sheet("results")
            res = res[res["student_id"].astype(str)==st.session_state.student_id]

            if not res.empty:
                file="report.pdf"
                c=canvas.Canvas(file)

                y=700
                for _,r in res.iterrows():
                    c.drawString(100,y,f"{r['subject']} {r['marks']} {r['grade']}")
                    y-=30

                c.save()

                with open(file,"rb") as f:
                    st.download_button("Download Report",f,file_name=file)