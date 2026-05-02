import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px

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
# LOAD DATA SAFE
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
# SESSION INIT
# =========================
if "db" not in st.session_state:
    st.session_state.db = load_data()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.role = None
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

                if r.lower() == "student":
                    st.session_state.link_id = str(match.iloc[0]["student_id"])

                st.rerun()
            else:
                st.error("Invalid login")

# =========================
# MAIN APP
# =========================
else:
    db = st.session_state.db

    st.sidebar.title(f"Role: {st.session_state.role}")

    if st.session_state.role == "Admin":
        menu = ["Dashboard","Students","Analytics","Logout"]
    elif st.session_state.role == "Faculty":
        menu = ["Marks Entry","Attendance","Logout"]
    else:
        menu = ["My Results","Logout"]

    choice = st.sidebar.radio("Menu", menu)

    # ================= ADMIN =================
    if choice == "Dashboard":
        st.title("Admin Dashboard")
        st.metric("Students", len(db["students"]))
        st.metric("Faculty", len(db["faculty"]))

    # ================= STUDENTS =================
    elif choice == "Students":
        st.title("Students List")
        st.dataframe(db["students"])

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

        if st.button("Save"):
            grade = "A+" if m>=90 else "A" if m>=75 else "B" if m>=60 else "C" if m>=40 else "F"

            new = pd.DataFrame([[sid,sub,m,grade]],
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

    # ================= REPORT CARD =================
    elif choice == "My Results":
        st.title("🎓 Student Report Card")

        df = db["results"].copy()

        # CLEAN DATA (IMPORTANT FIX)
        df["student_id"] = df["student_id"].astype(str).str.strip().str.replace(".0","",regex=False)

        sid = str(st.session_state.link_id).strip()

        student_df = df[df["student_id"] == sid]

        # ================= NO RESULT =================
        if student_df.empty:
            st.error("❌ No result found for your Student ID")
            st.info("Contact faculty to add marks")
        else:
            st.success("🎉 Report Card Loaded")

            st.markdown(f"""
            ### 🏫 Campus ERP Report Card  
            **Student ID:** {sid}  
            **Date:** {datetime.today().date()}  
            ---  
            """)

            st.table(student_df)

            total = student_df["marks"].sum()
            avg = student_df["marks"].mean()

            col1, col2 = st.columns(2)

            col1.metric("📊 Total Marks", int(total))
            col2.metric("📈 Average", round(avg,2))

            if avg >= 75:
                st.success("Excellent Performance ⭐")
            elif avg >= 60:
                st.info("Good Performance 👍")
            elif avg >= 40:
                st.warning("Average Performance ⚠️")
            else:
                st.error("Needs Improvement ❌")

    # ================= LOGOUT =================
    elif choice == "Logout":
        st.session_state.logged_in = False
        st.rerun()
