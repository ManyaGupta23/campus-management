import streamlit as st
import pandas as pd
import os
from datetime import datetime

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Campus Flow ERP", layout="wide", page_icon="🏫")

# Your exact Excel filename
FILE_NAME = "campus_flow_template.xlsx"

# =========================
# DATA ENGINE
# =========================
def load_all_data():
    """Reads all sheets from the Excel file."""
    if not os.path.exists(FILE_NAME):
        st.error(f"File '{FILE_NAME}' not found! Please ensure it is in your GitHub folder.")
        st.stop()
    
    # Mapping to your specific sheet names
    db = {
        "students": pd.read_excel(FILE_NAME, sheet_name="students"),
        "faculty": pd.read_excel(FILE_NAME, sheet_name="faculty"),
        "rooms": pd.read_excel(FILE_NAME, sheet_name="rooms"),
        "schedule": pd.read_excel(FILE_NAME, sheet_name="schedule"),
        "attendance": pd.read_excel(FILE_NAME, sheet_name="attendance"),
        "logs": pd.read_excel(FILE_NAME, sheet_name="logs")
    }
    return db

def save_all_data(db_dict):
    """Writes all dataframes back to the Excel sheets."""
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
        for sheet_name, df in db_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

# Initialize data in session state for speed
if "db" not in st.session_state:
    st.session_state.db = load_all_data()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.user_id = None

# =========================
# LOGIN SYSTEM
# =========================
if not st.session_state.logged_in:
    st.title("🏫 Campus Flow Management Login")
    
    with st.form("login_form"):
        user_id_input = st.text_input("Enter ID (Faculty ID or Student ID)")
        # For this template, any valid ID in the sheets acts as the login
        submit = st.form_submit_button("Login")
        
        if submit:
            # Validate against sheets
            fac_match = st.session_state.db["faculty"][st.session_state.db["faculty"]["faculty_id"].astype(str) == user_id_input]
            stud_match = st.session_state.db["students"][st.session_state.db["students"]["student_id"].astype(str) == user_id_input]
            
            if user_id_input == "admin":
                st.session_state.logged_in = True
                st.session_state.role = "Admin"
                st.rerun()
            elif not fac_match.empty:
                st.session_state.logged_in = True
                st.session_state.role = "Faculty"
                st.session_state.user_id = user_id_input
                st.rerun()
            elif not stud_match.empty:
                st.session_state.logged_in = True
                st.session_state.role = "Student"
                st.session_state.user_id = user_id_input
                st.rerun()
            else:
                st.error("User ID not found in database.")

# =========================
# MAIN DASHBOARD
# =========================
else:
    st.sidebar.title(f"Portal: {st.session_state.role}")
    st.sidebar.write(f"Logged in as: **{st.session_state.user_id if st.session_state.user_id else 'Admin'}**")
    
    # Navigation logic
    if st.session_state.role in ["Admin", "Faculty"]:
        menu = st.sidebar.radio("Navigation", 
            ["Student Directory", "Add/Update Student", "Attendance Records", "Faculty List", "Schedule & Rooms", "System Logs", "Logout"])
    else:
        menu = st.sidebar.radio("Navigation", ["My Profile", "My Attendance", "Class Schedule", "Logout"])

    # 1. STUDENT DIRECTORY (Admin/Faculty)
    if menu == "Student Directory":
        st.header("📋 Student Master Records")
        st.dataframe(st.session_state.db["students"], use_container_width=True)

    # 2. ADD / UPDATE STUDENT (Admin/Faculty)
    elif menu == "Add/Update Student":
        st.header("📝 Register or Update Student")
        with st.form("student_form"):
            c1, c2 = st.columns(2)
            with c1:
                s_id = st.text_input("Student ID (Unique)")
                s_name = st.text_input("Full Name")
                s_gender = st.selectbox("Gender", ["M", "F", "Other"])
                s_course = st.text_input("Course")
            with c2:
                s_year = st.number_input("Current Year", min_value=1, step=1)
                s_section = st.text_input("Section")
                s_email = st.text_input("Email Address")
                s_date = st.date_input("Admission Date", value=datetime.now())
            
            if st.form_submit_button("Save Changes"):
                new_student = {
                    "student_id": s_id, "name": s_name, "gender": s_gender, 
                    "course": s_course, "year": s_year, "section": s_section, 
                    "admission_date": str(s_date), "email": s_email,
                    "attendance_percentage": 0, "status": "Active"
                }
                
                df = st.session_state.db["students"]
                # If ID exists, update; else, append
                if s_id in df["student_id"].astype(str).values:
                    df.loc[df["student_id"].astype(str) == s_id, list(new_student.keys())] = list(new_student.values())
                    st.success(f"Updated record for {s_name}")
                else:
                    st.session_state.db["students"] = pd.concat([df, pd.DataFrame([new_student])], ignore_index=True)
                    st.success(f"Added new student: {s_name}")
                
                save_all_data(st.session_state.db)

    # 3. SCHEDULE & ROOMS
    elif menu == "Schedule & Rooms":
        st.header("🏢 Campus Logistics")
        tab1, tab2 = st.tabs(["Class Schedule", "Room Capacity"])
        with tab1:
            st.dataframe(st.session_state.db["schedule"], use_container_width=True)
        with tab2:
            st.dataframe(st.session_state.db["rooms"], use_container_width=True)

    # 4. STUDENT VIEW (Self)
    elif menu == "My Profile":
        st.header("👤 Personal Profile")
        profile = st.session_state.db["students"][st.session_state.db["students"]["student_id"].astype(str) == st.session_state.user_id]
        if not profile.empty:
            st.table(profile.T) # Transposed for better readability
        else:
            st.error("Profile not found.")

    elif menu == "My Attendance":
        st.header("📅 My Attendance History")
        attn = st.session_state.db["attendance"][st.session_state.db["attendance"]["student_id"].astype(str) == st.session_state.user_id]
        st.dataframe(attn, use_container_width=True)

    # 5. LOGOUT
    elif menu == "Logout":
        st.session_state.logged_in = False
        st.rerun()
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
