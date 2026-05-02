# =========================================
# CAMPUS FLOW ERP SYSTEM (EXCEL BASED - NO SQL)
# Streamlit + Pandas + Excel (Production Simplified)
# =========================================

import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
from io import BytesIO
from reportlab.pdfgen import canvas

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Campus ERP", layout="wide")

FILE_NAME = "campus_flow_template.xlsx"

SHEETS = ["students", "faculty", "rooms", "schedule", "attendance", "users", "results"]

# =========================
# LOAD DATA
# =========================

def load_data():
    db = {}

    if not os.path.exists(FILE_NAME):
        st.warning("Excel file not found. Creating new template...")
        with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
            for s in SHEETS:
                pd.DataFrame().to_excel(writer, sheet_name=s, index=False)

    for s in SHEETS:
        try:
            db[s] = pd.read_excel(FILE_NAME, sheet_name=s)
        except:
            db[s] = pd.DataFrame()

    return db

# =========================
# SAVE DATA
# =========================

def save_data(db):
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="w") as writer:
        for sheet, df in db.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

# =========================
# INIT SESSION
# =========================

if "db" not in st.session_state:
    st.session_state.db = load_data()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.user = None
    st.session_state.link_id = None

# =========================
# LOGIN
# =========================

if not st.session_state.logged_in:
    st.title("🏫 Campus ERP Login")

    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    r = st.selectbox("Role", ["Admin", "Faculty", "Student"])

    if st.button("Login"):
        users = st.session_state.db["users"]

        if not users.empty and all(col in users.columns for col in ["username","password","role"]):
            match = users[(users["username"] == u) & (users["password"] == p) & (users["role"] == r)]

            if not match.empty or (u == "admin" and p == "admin123"):
                st.session_state.logged_in = True
                st.session_state.role = r
                st.session_state.user = u

                if r == "Student" and "student_id" in match.columns:
                    st.session_state.link_id = str(match.iloc[0]["student_id"])
                elif r == "Faculty" and "faculty_id" in match.columns:
                    st.session_state.link_id = str(match.iloc[0]["faculty_id"])

                st.rerun()
        else:
            st.error("Invalid credentials or user sheet missing")

# =========================
# DASHBOARD
# =========================
else:
    st.sidebar.title(f"Role: {st.session_state.role}")

    if st.session_state.role == "Admin":
        menu = ["Dashboard", "Manage Students", "Conflict Detection", "Analytics", "Logout"]
    elif st.session_state.role == "Faculty":
        menu = ["Mark Attendance", "Enter Marks", "My Classes", "Logout"]
    else:
        menu = ["My Profile", "Results", "Logout"]

    choice = st.sidebar.radio("Menu", menu)

    # ================= ADMIN =================

    if choice == "Dashboard":
        st.header("Admin Dashboard")
        st.metric("Students", len(st.session_state.db["students"]))
        st.metric("Faculty", len(st.session_state.db["faculty"]))

    elif choice == "Manage Students":
        st.header("Students")
        st.dataframe(st.session_state.db["students"])

    elif choice == "Conflict Detection":
        st.header("Conflict Detection")
        schedule = st.session_state.db["schedule"]

        if not schedule.empty and "room" in schedule.columns:
            dup = schedule[schedule.duplicated(subset=["room","time"], keep=False)]
            if not dup.empty:
                st.error("Conflicts Found")
                st.dataframe(dup)
            else:
                st.success("No Conflicts")
        else:
            st.warning("Schedule data missing")

    elif choice == "Analytics":
        st.header("Analytics")
        df = st.session_state.db["results"]

        if not df.empty and "marks" in df.columns:
            st.plotly_chart(px.histogram(df, x="marks"))

    # ================= FACULTY =================

    elif choice == "Mark Attendance":
        st.header("Attendance")

        sid = st.text_input("Student ID")
        cid = st.text_input("Class ID")
        status = st.selectbox("Status", ["Present","Absent"])

        if st.button("Submit"):
            df = st.session_state.db["attendance"]

            new_row = pd.DataFrame([[sid,cid,str(datetime.now().date()),status]],
                                   columns=["student_id","class_id","date","status"])

            st.session_state.db["attendance"] = pd.concat([df,new_row], ignore_index=True)
            save_data(st.session_state.db)
            st.success("Saved")

    elif choice == "Enter Marks":
        st.header("Marks Entry")

        sid = st.text_input("Student ID")
        subject = st.text_input("Subject")
        marks = st.number_input("Marks",0,100)
        grade = st.text_input("Grade")

        if st.button("Save"):
            df = st.session_state.db["results"]

            new_row = pd.DataFrame([[sid,subject,marks,grade]],
                                   columns=["student_id","subject","marks","grade"])

            st.session_state.db["results"] = pd.concat([df,new_row], ignore_index=True)
            save_data(st.session_state.db)
            st.success("Saved")

    elif choice == "My Classes":
        st.header("My Classes")
        st.dataframe(st.session_state.db["schedule"])

    # ================= STUDENT =================

    elif choice == "My Profile":
        st.header("Profile")
        st.write(st.session_state.user)

    elif choice == "Results":
        st.header("Results")

        df = st.session_state.db["results"]
        st.dataframe(df)

        if st.button("Download PDF"):
            buffer = BytesIO()
            c = canvas.Canvas(buffer)

            c.drawString(100,800,"Campus ERP Report")
            y = 760

            for _,row in df.iterrows():
                c.drawString(100,y,f"{row['subject']} - {row['marks']}")
                y -= 20

            c.save()
            buffer.seek(0)

            st.download_button("Download PDF", buffer, file_name="report.pdf")

    elif choice == "Logout":
        st.session_state.logged_in = False
        st.rerun()

