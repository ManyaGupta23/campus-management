# =========================================
# CAMPUS FLOW ERP SYSTEM (EXCEL FIXED VERSION)
# Streamlit + Pandas + Excel (Stable Version)
# =========================================

import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
from io import BytesIO
from reportlab.pdfgen import canvas
import qrcode

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Campus ERP", layout="wide")

FILE_NAME = "campus_flow_template.xlsx"

SHEETS = [
    "students", "faculty", "rooms",
    "schedule", "attendance", "logs",
    "users", "results"
]

# =========================
# DEFAULT SCHEMA
# =========================
SCHEMA = {
    "students": ["student_id","name","gender","course","year","section","admission_date","attendance_percentage","status","email"],
    "faculty": ["faculty_id","name","subject"],
    "rooms": ["room_id","capacity","type"],
    "schedule": ["class_id","faculty_id","room_id","time_slot","subject"],
    "attendance": ["student_id","class_id","date","status"],
    "logs": ["log_id","student_id","entry_time","exit_time"],
    "users": ["username","password","role","student_id","faculty_id"],
    "results": ["student_id","subject","marks","grade"]
}

# =========================
# LOAD DATA (FIXED)
# =========================
def load_data():
    db = {}

    if not os.path.exists(FILE_NAME):
        with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
            for sheet, cols in SCHEMA.items():
                pd.DataFrame(columns=cols).to_excel(writer, sheet_name=sheet, index=False)

    for sheet in SHEETS:
        db[sheet] = pd.read_excel(FILE_NAME, sheet_name=sheet)

    return db

# =========================
# SAVE DATA (FIXED)
# =========================
def save_data(db):
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="w") as writer:
        for sheet, df in db.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

    # 🔥 IMPORTANT: reload after saving
    st.session_state.db = load_data()

# =========================
# GRADE SYSTEM
# =========================
def get_grade(marks):
    if marks >= 90:
        return "A+"
    elif marks >= 75:
        return "A"
    elif marks >= 60:
        return "B"
    elif marks >= 40:
        return "C"
    return "F"

# =========================
# SESSION INIT
# =========================
if "db" not in st.session_state:
    st.session_state.db = load_data()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.user = None
    st.session_state.link_id = None

# =========================
# LOGIN SYSTEM
# =========================
if not st.session_state.logged_in:
    st.title("🏫 Campus Flow ERP Login")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    role = st.selectbox("Role", ["Admin", "Faculty", "Student"])

    if st.button("Login"):
        users = st.session_state.db["users"]

        match = users[
            (users["username"] == username) &
            (users["password"] == password) &
            (users["role"] == role)
        ]

        if username == "admin" and password == "admin123":
            st.session_state.logged_in = True
            st.session_state.role = "Admin"
            st.session_state.user = username

        elif not match.empty:
            st.session_state.logged_in = True
            st.session_state.role = role
            st.session_state.user = username

            if role == "Student":
                st.session_state.link_id = str(match.iloc[0]["student_id"])
            elif role == "Faculty":
                st.session_state.link_id = str(match.iloc[0]["faculty_id"])
        else:
            st.error("Invalid Login")

        st.rerun()

# =========================
# MAIN DASHBOARD
# =========================
else:
    st.sidebar.title(f"Role: {st.session_state.role}")

    if st.session_state.role == "Admin":
        menu = ["Dashboard","Students","Faculty","Schedule","Analytics","Conflict","Logout"]
    elif st.session_state.role == "Faculty":
        menu = ["Attendance","Marks Entry","Logout"]
    else:
        menu = ["My Results","Logout"]

    choice = st.sidebar.radio("Menu", menu)

    db = st.session_state.db

    # ================= ADMIN DASHBOARD =================
    if choice == "Dashboard":
        st.title("📊 Admin Dashboard")

        st.metric("Students", len(db["students"]))
        st.metric("Faculty", len(db["faculty"]))
        st.metric("Classes", len(db["schedule"]))

    # ================= STUDENTS =================
    elif choice == "Students":
        st.title("👨‍🎓 Students")

        df = db["students"]
        st.dataframe(df)

        st.subheader("Add Student")

        sid = st.text_input("Student ID")
        name = st.text_input("Name")
        course = st.text_input("Course")

        if st.button("Add Student"):
            new = pd.DataFrame([[sid,name,None,course,None,None,str(datetime.today().date()),0,"Active",None]],
                               columns=SCHEMA["students"])

            db["students"] = pd.concat([df,new], ignore_index=True)
            save_data(db)

            st.success("Student Added")
            st.rerun()

        st.subheader("Delete Student")

        del_id = st.text_input("Student ID to delete")

        if st.button("Delete Student"):
            df = df[df["student_id"] != del_id]
            db["students"] = df
            save_data(db)

            st.success("Deleted")
            st.rerun()

    # ================= FACULTY =================
    elif choice == "Faculty":
        st.title("👨‍🏫 Faculty")
        st.dataframe(db["faculty"])

    # ================= SCHEDULE =================
    elif choice == "Schedule":
        st.title("📅 Schedule")
        st.dataframe(db["schedule"])

    # ================= CONFLICT =================
    elif choice == "Conflict":
        st.title("⚠️ Conflict Detection")

        sch = db["schedule"]

        if not sch.empty:
            dup = sch[sch.duplicated(subset=["room_id","time_slot"], keep=False)]
            st.dataframe(dup)
        else:
            st.warning("No schedule data")

    # ================= ANALYTICS =================
    elif choice == "Analytics":
        st.title("📈 Analytics")

        res = db["results"]

        if not res.empty and "marks" in res.columns:
            st.plotly_chart(px.histogram(res, x="marks"))
        else:
            st.warning("No data")

    # ================= MARKS ENTRY =================
    elif choice == "Marks Entry":
        st.title("📝 Marks Entry")

        sid = st.text_input("Student ID")
        subject = st.text_input("Subject")
        marks = st.number_input("Marks",0,100)

        grade = get_grade(marks)
        st.info(f"Grade: {grade}")

        if st.button("Save Marks"):
            new = pd.DataFrame([[sid,subject,marks,grade]],
                               columns=SCHEMA["results"])

            db["results"] = pd.concat([db["results"],new], ignore_index=True)
            save_data(db)

            st.success("Saved")
            st.rerun()

    # ================= ATTENDANCE =================
    elif choice == "Attendance":
        st.title("📌 Attendance")

        sid = st.text_input("Student ID")
        cid = st.text_input("Class ID")
        status = st.selectbox("Status",["Present","Absent"])

        if st.button("Save Attendance"):
            new = pd.DataFrame([[sid,cid,str(datetime.today().date()),status]],
                               columns=SCHEMA["attendance"])

            db["attendance"] = pd.concat([db["attendance"],new], ignore_index=True)
            save_data(db)

            st.success("Saved")
            st.rerun()

    # ================= STUDENT RESULTS =================
    elif choice == "My Results":
        st.title("🎓 My Results")

        df = db["results"]

        if st.session_state.link_id:
            df = df[df["student_id"].astype(str) == str(st.session_state.link_id)]

        st.dataframe(df)

        # QR CODE
        if not df.empty:
            qr = qrcode.make(f"Student ID: {st.session_state.link_id}")
            buf = BytesIO()
            qr.save(buf)
            st.image(buf.getvalue())

        # PDF
        if st.button("Download PDF"):
            buffer = BytesIO()
            c = canvas.Canvas(buffer)

            c.drawString(100,800,"Campus ERP Report")

            y = 760
            for _,row in df.iterrows():
                c.drawString(100,y,f"{row['subject']} - {row['marks']} ({row['grade']})")
                y -= 20

            c.save()
            buffer.seek(0)

            st.download_button("Download", buffer, file_name="result.pdf")

    # ================= LOGOUT =================
    elif choice == "Logout":
        st.session_state.logged_in = False
        st.rerun()
