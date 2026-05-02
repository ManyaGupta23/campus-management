import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
from reportlab.pdfgen import canvas

# =========================
# CONFIG & DATA ENGINE
# =========================
EXCEL_FILE = "campus_flow_template.xlsx"
st.set_page_config(page_title="Campus Flow ERP", layout="wide")

def load_all_data():
    """Load all sheets at once to prevent connection errors."""
    sheets = ["students", "faculty", "rooms", "schedule", "attendance", "logs", "users", "results"]
    db = {}
    for s in sheets:
        try:
            db[s] = pd.read_excel(EXCEL_FILE, sheet_name=s)
        except:
            # Create empty df if sheet is missing
            db[s] = pd.DataFrame() 
    return db

def save_all_data(db):
    """Safely saves all sheets back to the main Excel file."""
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        for sheet_name, df in db.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

# Initialize Session Data
if "db" not in st.session_state:
    st.session_state.db = load_all_data()

# =========================
# LOGIN LOGIC
# =========================
if "login" not in st.session_state:
    st.session_state.login = False

if not st.session_state.login:
    st.title("🔐 Campus Flow Login")
    with st.form("login"):
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        r = st.selectbox("Role", ["Admin", "Faculty", "Student"])
        if st.form_submit_button("Login"):
            users = st.session_state.db["users"]
            user_row = users[(users["username"] == u) & (users["password"] == p) & (users["role"] == r)]
            
            if not user_row.empty:
                st.session_state.login = True
                st.session_state.role = r
                st.session_state.user = u
                # Store IDs for filtering
                if r == "Student": st.session_state.sid = str(user_row.iloc[0]["student_id"])
                if r == "Faculty": st.session_state.fid = str(user_row.iloc[0]["faculty_id"])
                st.rerun()
            else:
                st.error("Invalid Credentials")

# =========================
# MAIN DASHBOARD
# =========================
else:
    role = st.session_state.role
    st.sidebar.title(f"🏫 {role} Portal")
    
    if role == "Admin":
        menu = ["Dashboard", "Campus Flow", "Conflict Detection", "Logout"]
    elif role == "Faculty":
        menu = ["Mark Attendance", "Enter Marks", "Logout"]
    else:
        menu = ["My Performance", "Report Card", "Logout"]

    choice = st.sidebar.radio("Navigation", menu)

    # --- ADMIN: CAMPUS FLOW ANALYTICS ---
    if choice == "Campus Flow":
        st.header("🚶 Live Campus Analytics")
        logs = st.session_state.db["logs"]
        if not logs.empty:
            # Convert time strings to usable hour numbers
            logs['hour'] = pd.to_datetime(logs['entry_time'], format='%H:%M:%S').dt.hour
            fig = px.histogram(logs, x="hour", title="Peak Entry Hours", nbins=24)
            st.plotly_chart(fig)
        
        rooms = st.session_state.db["rooms"]
        fig2 = px.pie(rooms, names="type", values="capacity", title="Campus Capacity Distribution")
        st.plotly_chart(fig2)

    # --- FACULTY: ATTENDANCE ---
    elif choice == "Mark Attendance":
        st.header("📅 Record Attendance")
        with st.form("att_form"):
            sid = st.text_input("Student ID")
            cid = st.text_input("Class ID")
            status = st.selectbox("Status", ["Present", "Absent"])
            if st.form_submit_button("Submit"):
                new_att = pd.DataFrame([[sid, cid, str(datetime.now().date()), status]], 
                                       columns=st.session_state.db["attendance"].columns)
                st.session_state.db["attendance"] = pd.concat([st.session_state.db["attendance"], new_att], ignore_index=True)
                save_all_data(st.session_state.db)
                st.success("Attendance Logged!")

    # --- STUDENT: REPORT CARD ---
    elif choice == "Report Card":
        st.header("📄 Download Report Card")
        res = st.session_state.db["results"]
        my_res = res[res["student_id"].astype(str) == st.session_state.sid]
        
        if not my_res.empty:
            st.dataframe(my_res)
            if st.button("Generate PDF"):
                filename = f"Report_{st.session_state.sid}.pdf"
                c = canvas.Canvas(filename)
                c.setFont("Helvetica-Bold", 16)
                c.drawString(100, 800, f"Official Report Card: {st.session_state.user}")
                c.setFont("Helvetica", 12)
                y = 750
                for _, row in my_res.iterrows():
                    c.drawString(100, y, f"Subject: {row['subject']} | Marks: {row['marks']} | Grade: {row['grade']}")
                    y -= 25
                c.save()
                with open(filename, "rb") as f:
                    st.download_button("Download PDF", f, file_name=filename)
        else:
            st.warning("No results found.")

    elif choice == "Logout":
        st.session_state.login = False
        st.rerun()
