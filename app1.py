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
