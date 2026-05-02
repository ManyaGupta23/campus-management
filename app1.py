import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
from reportlab.pdfgen import canvas

# =========================
# CONFIG & DATA ENGINE
# =========================
st.set_page_config(page_title="Campus Flow ERP", layout="wide", page_icon="🏫")

FILE_NAME = "campus_flow_template.xlsx"

def load_all_data():
    """Reads all sheets. If a sheet is missing, it returns an empty DataFrame."""
    sheets = ["students", "faculty", "rooms", "schedule", "attendance", "logs", "users", "results"]
    db = {}
    if not os.path.exists(FILE_NAME):
        st.error(f"File {FILE_NAME} not found! Please upload it to GitHub.")
        st.stop()
    
    for s in sheets:
        try:
            db[s] = pd.read_excel(FILE_NAME, sheet_name=s)
        except Exception:
            db[s] = pd.DataFrame() 
    return db

def save_all_data(db_dict):
    """Saves the entire database state back to Excel."""
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
        for sheet_name, df in db_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

# Initialize Session Data
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
        u_input = st.text_input("Username")
        p_input = st.text_input("Password", type="password")
        r_input = st.selectbox("Role", ["Admin", "Faculty", "Student"])
        submit = st.form_submit_button("Login")
        
        if submit:
            users_df = st.session_state.db["users"]
            if u_input == "admin" and p_input == "admin123":
                st.session_state.logged_in = True
                st.session_state.role = "Admin"
                st.rerun()
            elif not users_df.empty:
                match = users_df[(users_df["username"] == u_input) & 
                                 (users_df["password"] == p_input) & 
                                 (users_df["role"] == r_input)]
                if not match.empty:
                    st.session_state.logged_in = True
                    st.session_state.role = r_input
                    st.session_state.user_id = u_input
                    # Link ID for Student/Faculty
                    if r_input == "Student": st.session_state.link_id = str(match.iloc[0]["student_id"])
                    if r_input == "Faculty": st.session_state.link_id = str(match.iloc[0]["faculty_id"])
                    st.rerun()
                else:
                    st.error("Invalid Credentials")
            else:
                st.error("User database is empty. Please check the 'users' sheet.")

# =========================
# MAIN DASHBOARD
# =========================
else:
    st.sidebar.title(f"Portal: {st.session_state.role}")
    st.sidebar.write(f"User: **{st.session_state.user_id if st.session_state.user_id else 'Admin'}**")
    
    if st.session_state.role == "Admin":
        menu = ["Dashboard", "Campus Flow Analytics", "Manage Students", "Conflict Detection", "Logout"]
    elif st.session_state.role == "Faculty":
        menu = ["Mark Attendance", "Enter Marks", "My Classes", "Logout"]
    else:
        menu = ["My Profile", "Performance Analytics", "Report Card", "Logout"]

    choice = st.sidebar.radio("Navigation", menu)

    # --- ADMIN: ANALYTICS ---
    if choice == "Dashboard":
        st.header("📊 Campus Overview")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Students", len(st.session_state.db["students"]))
        c2.metric("Faculty Strength", len(st.session_state.db["faculty"]))
        c3.metric("Room Capacity", st.session_state.db["rooms"]["capacity"].sum())

    elif choice == "Campus Flow Analytics":
        st.header("📈 Infrastructure & Flow")
        logs = st.session_state.db["logs"]
        if not logs.empty:
            logs['hour'] = pd.to_datetime(logs['entry_time'], format='%H:%M:%S', errors='coerce').dt.hour
            st.plotly_chart(px.histogram(logs, x="hour", title="Entry Traffic by Hour"))
        
        rooms = st.session_state.db["rooms"]
        st.plotly_chart(px.pie(rooms, names="type", values="capacity", title="Room Type Distribution"))

    # --- FACULTY: RECORDING DATA ---
    elif choice == "Mark Attendance":
        st.header("📅 Attendance Entry")
        with st.form("att"):
            sid = st.text_input("Student ID")
            cid = st.text_input("Class ID")
            status = st.selectbox("Status", ["Present", "Absent"])
            if st.form_submit_button("Submit"):
                new_data = pd.DataFrame([[sid, cid, str(datetime.now().date()), status]], 
                                        columns=st.session_state.db["attendance"].columns)
                st.session_state.db["attendance"] = pd.concat([st.session_state.db["attendance"], new_data], ignore_index=True)
                save_all_data(st.session_state.db)
                st.success("Attendance Recorded!")

    # --- STUDENT: RESULTS & PDF ---
    elif choice == "Report Card":
        st.header("📄 Academic Report Card")
        res = st.session_state.db["results"]
        my_res = res[res["student_id"].astype(str) == st.session_state.link_id]
        
        if not my_res.empty:
            st.table(my_res)
            if st.button("Download PDF Report"):
                pdf_file = f"Report_{st.session_state.link_id}.pdf"
                c = canvas.Canvas(pdf_file)
                c.drawString(100, 800, f"Campus Flow ERP - Official Report")
                c.drawString(100, 780, f"Student: {st.session_state.user_id}")
                y = 750
                for _, row in my_res.iterrows():
                    c.drawString(100, y, f"{row['subject']}: {row['marks']} ({row['grade']})")
                    y -= 20
                c.save()
                with open(pdf_file, "rb") as f:
                    st.download_button("Click to Download", f, file_name=pdf_file)
        else:
            st.info("No marks uploaded yet.")

    elif choice == "Logout":
        st.session_state.logged_in = False
        st.rerun()
