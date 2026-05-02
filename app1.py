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
# GRADE SYSTEM
# =========================
def get_grade(m):
    if m >= 90: return "A+"
    elif m >= 75: return "A"
    elif m >= 60: return "B"
    elif m >= 40: return "C"
    return "F"

# =========================
# REPORT CARD + QR
# =========================
def generate_report_card(student_id, df):
    buffer = BytesIO()
    c = canvas.Canvas(buffer)

    c.setFont("Helvetica-Bold", 16)
    c.drawString(180, 800, "CAMPUS ERP - REPORT CARD")

    c.setFont("Helvetica", 12)
    c.drawString(100, 770, f"Student ID: {student_id}")
    c.drawString(100, 750, f"Date: {datetime.now().date()}")

    y = 700
    total = 0
    count = 0

    for _, row in df.iterrows():
        line = f"{row['subject']} : {row['marks']} ({row['grade']})"
        c.drawString(100, y, line)
        y -= 20

        total += float(row["marks"])
        count += 1

    avg = total / count if count > 0 else 0
    c.drawString(100, y-20, f"Average: {avg:.2f}")

    # ================= QR CODE =================
    qr_text = f"Student:{student_id}|Avg:{avg:.2f}|Verified:CampusERP"
    qr = qrcode.make(qr_text)

    qr_buffer = BytesIO()
    qr.save(qr_buffer)
    qr_buffer.seek(0)

    c.drawImage(qr_buffer, 400, 650, width=120, height=120)

    c.save()
    buffer.seek(0)

    return buffer

# =========================
# SESSION
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
        st.title("Students")
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

        g = get_grade(m)

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

    # ================= RESULTS + CERTIFICATE =================
    elif choice == "My Results":
        st.title("My Report Card")

        df = db["results"].copy()

        df["student_id"] = df["student_id"].astype(str).str.strip().str.replace(".0","",regex=False)
        sid = str(st.session_state.link_id).strip()

        result_df = df[df["student_id"] == sid]

        if result_df.empty:
            st.warning("No results found")
        else:
            st.success("Your Results")
            st.dataframe(result_df)

            # ================= DOWNLOAD CERTIFICATE =================
            if st.button("Download Report Card (PDF)"):
                pdf = generate_report_card(sid, result_df)

                st.download_button(
                    "Download Certificate",
                    pdf,
                    file_name=f"report_card_{sid}.pdf",
                    mime="application/pdf"
                )

    # ================= LOGOUT =================
    elif choice == "Logout":
        st.session_state.logged_in = False
        st.rerun()
