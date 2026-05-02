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

SHEETS = ["students","faculty","rooms","schedule","attendance","logs","users","results"]

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
# CREATE FILE
# =========================
def create_file():
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
        for sheet, cols in SCHEMA.items():
            pd.DataFrame(columns=cols).to_excel(writer, sheet_name=sheet, index=False)

# =========================
# LOAD DATA (SAFE)
# =========================
def load_data():
    db = {}

    if not os.path.exists(FILE_NAME):
        create_file()

    xls = pd.ExcelFile(FILE_NAME)

    for sheet in SHEETS:
        if sheet in xls.sheet_names:
            db[sheet] = pd.read_excel(FILE_NAME, sheet_name=sheet)
        else:
            db[sheet] = pd.DataFrame(columns=SCHEMA[sheet])

    return db

# =========================
# SAVE DATA
# =========================
def save_data(db):
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="w") as writer:
        for sheet, df in db.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

    st.session_state.db = load_data()

# =========================
# GRADE SYSTEM
# =========================
def get_grade(m):
    if m >= 90: return "A+"
    elif m >= 75: return "A"
    elif m >= 60: return "B"
    elif m >= 40: return "C"
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
# LOGIN
# =========================
if not st.session_state.logged_in:
    st.title("🏫 Campus ERP Login")

    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    r = st.selectbox("Role", ["Admin","Faculty","Student"])

    if st.button("Login"):
        users = st.session_state.db["users"]

        users["username"] = users["username"].astype(str)
        users["password"] = users["password"].astype(str)
        users["role"] = users["role"].astype(str)

        u = str(u).strip()
        p = str(p).strip()
        r = str(r).strip()

        # ADMIN LOGIN
        if u == "admin" and p == "admin123":
            st.session_state.logged_in = True
            st.session_state.role = "Admin"
            st.rerun()

        else:
            match = users[
                (users["username"] == u) &
                (users["password"] == p) &
                (users["role"].str.lower() == r.lower())
            ]

            if not match.empty:
                st.session_state.logged_in = True
                st.session_state.role = r
                st.session_state.user = u

                if r.lower() == "student":
                    st.session_state.link_id = str(match.iloc[0]["student_id"]).strip()
                elif r.lower() == "faculty":
                    st.session_state.link_id = str(match.iloc[0]["faculty_id"]).strip()

                st.rerun()
            else:
                st.error("Invalid login")

# =========================
# MAIN APP
# =========================
else:
    st.sidebar.title(f"Role: {st.session_state.role}")

    if st.session_state.role == "Admin":
        menu = ["Dashboard","Students","Analytics","Logout"]
    elif st.session_state.role == "Faculty":
        menu = ["Attendance","Marks Entry","Logout"]
    else:
        menu = ["My Results","Logout"]

    choice = st.sidebar.radio("Menu", menu)

    db = st.session_state.db

    # ================= ADMIN =================
    if choice == "Dashboard":
        st.title("Admin Dashboard")
        st.metric("Students", len(db["students"]))
        st.metric("Faculty", len(db["faculty"]))

    # ================= STUDENTS =================
    elif choice == "Students":
        st.title("Students")
        df = db["students"]
        st.dataframe(df)

        sid = st.text_input("ID")
        name = st.text_input("Name")
        course = st.text_input("Course")

        if st.button("Add"):
            new = pd.DataFrame([[sid,name,None,course,None,None,str(datetime.today().date()),0,"Active",None]],
                               columns=SCHEMA["students"])
            db["students"] = pd.concat([df,new], ignore_index=True)
            save_data(db)
            st.success("Added")
            st.rerun()

        del_id = st.text_input("Delete ID")

        if st.button("Delete"):
            db["students"] = df[df["student_id"] != del_id]
            save_data(db)
            st.success("Deleted")
            st.rerun()

    # ================= ANALYTICS =================
    elif choice == "Analytics":
        st.title("Analytics")

        res = db["results"]

        if not res.empty:
            st.plotly_chart(px.histogram(res, x="marks"))
        else:
            st.warning("No data")

    # ================= MARKS =================
    elif choice == "Marks Entry":
        st.title("Marks Entry")

        sid = st.text_input("Student ID")
        sub = st.text_input("Subject")
        m = st.number_input("Marks",0,100)

        g = get_grade(m)
        st.info(f"Grade: {g}")

        if st.button("Save"):
            new = pd.DataFrame([[sid,sub,m,g]],
                               columns=SCHEMA["results"])
            db["results"] = pd.concat([db["results"],new], ignore_index=True)
            save_data(db)
            st.success("Saved")
            st.rerun()

    # ================= ATTENDANCE =================
    elif choice == "Attendance":
        st.title("Attendance")

        sid = st.text_input("Student ID")
        cid = st.text_input("Class ID")
        status = st.selectbox("Status",["Present","Absent"])

        if st.button("Save"):
            new = pd.DataFrame([[sid,cid,str(datetime.today().date()),status]],
                               columns=SCHEMA["attendance"])
            db["attendance"] = pd.concat([db["attendance"],new], ignore_index=True)
            save_data(db)
            st.success("Saved")
            st.rerun()

    # ================= RESULTS =================
    elif choice == "My Results":
        st.title("My Results")

        df = db["results"]

        # 🔥 FINAL FIX (IMPORTANT)
        df["student_id"] = df["student_id"].astype(str).str.strip().str.replace(".0","")
        sid = str(st.session_state.link_id).strip()

        df = df[df["student_id"] == sid]

        st.dataframe(df)

    # ================= LOGOUT =================
    elif choice == "Logout":
        st.session_state.logged_in = False
        st.rerun()
