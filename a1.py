# ================= IMPORT =================
import streamlit as st
import pandas as pd
import os

# ================= CONFIG =================
st.set_page_config(page_title="Campus ERP", layout="wide")

FILE = "campus_flow_template.xlsx"

SCHEMA = {
    "users": ["username","password","role","student_id","faculty_id"]
}

# ================= UTIL =================
def clean(x):
    return str(x).strip().replace(".0","")

# ================= FILE =================
def load():
    if not os.path.exists(FILE):
        with pd.ExcelWriter(FILE, engine="openpyxl") as w:
            pd.DataFrame(columns=SCHEMA["users"]).to_excel(w, sheet_name="users", index=False)

    xls = pd.ExcelFile(FILE)

    if "users" in xls.sheet_names:
        df = pd.read_excel(FILE, sheet_name="users")
    else:
        df = pd.DataFrame(columns=SCHEMA["users"])

    # FIXED CLEANING (NO applymap)
    df = df.astype(str)
    for col in df.columns:
        df[col] = df[col].apply(clean)

    return {"users": df}

# ================= UI BACKGROUND =================
def set_bg():
    st.markdown(
        """
        <style>
        .stApp {
            background-image: url("campus.jpg");
            background-size: cover;
            background-position: center;
        }

        .login-box {
            background-color: rgba(255,255,255,0.92);
            padding: 35px;
            border-radius: 15px;
            width: 350px;
            margin: auto;
            margin-top: 120px;
            box-shadow: 0px 0px 25px rgba(0,0,0,0.3);
        }

        .title {
            text-align: center;
            font-size: 28px;
            font-weight: bold;
            color: #C9A227;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

# ================= SESSION =================
if "db" not in st.session_state:
    st.session_state.db = load()

if "login" not in st.session_state:
    st.session_state.login = False

# ================= LOGIN =================
if not st.session_state.login:

    set_bg()

    st.markdown('<div class="login-box">', unsafe_allow_html=True)

    st.markdown('<p class="title">🏫 Campus ERP Login</p>', unsafe_allow_html=True)

    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    r = st.selectbox("Role", ["Admin","Faculty","Student"])

    if st.button("Login"):

        users = st.session_state.db["users"]

        # ADMIN LOGIN
        if u == "admin" and p == "admin123":
            st.session_state.login = True
            st.session_state.role = "Admin"
            st.rerun()

        # USER LOGIN
        match = users[
            (users["username"] == u) &
            (users["password"] == p) &
            (users["role"] == r)
        ]

        if not match.empty:
            st.session_state.login = True
            st.session_state.role = r

            if r == "Student":
                st.session_state.link = match.iloc[0]["student_id"]
            else:
                st.session_state.link = match.iloc[0]["faculty_id"]

            st.rerun()
        else:
            st.error("❌ Invalid Login")

    st.markdown('</div>', unsafe_allow_html=True)

# ================= AFTER LOGIN =================
else:
    st.success(f"Welcome {st.session_state.role} 🎉")

    if st.button("Logout"):
        st.session_state.clear()
        st.rerun()
# ================= MAIN APP =================
else:
    db = st.session_state.db
    role = st.session_state.role

    st.sidebar.title(f"👤 {role} Panel")

    # ================= MENU =================
    if role == "Admin":
        menu = ["Dashboard","Students","Faculty","Schedule","Logout"]
    elif role == "Faculty":
        menu = ["Students","My Schedule","Marks Entry","Attendance","Logout"]
    else:
        menu = ["My Results","Certificate","Logout"]

    choice = st.sidebar.radio("Menu", menu)

    # ================= ADMIN =================
    if role == "Admin":

        if choice == "Dashboard":
            st.title("📊 Admin Dashboard")

            st.metric("Students", len(db["students"]))
            st.metric("Faculty", len(db["faculty"]))
            st.metric("Rooms", len(db["rooms"]))

        # ===== STUDENTS =====
        elif choice == "Students":
            st.subheader("Students Data")
            st.dataframe(db["students"])

            st.subheader("➕ Add Student")

            sid = st.text_input("Student ID")
            name = st.text_input("Name")
            course = st.text_input("Course")
            email = st.text_input("Email")

            if st.button("Add Student"):
                new = pd.DataFrame([[sid,name,course,email]],
                                   columns=SCHEMA["students"])
                db["students"] = pd.concat([db["students"],new],ignore_index=True)
                save(db)
                st.success("Student Added")
                st.rerun()

        # ===== FACULTY =====
        elif choice == "Faculty":
            st.subheader("Faculty Data")
            st.dataframe(db["faculty"])

            fid = st.text_input("Faculty ID")
            fname = st.text_input("Name")
            sub = st.text_input("Subject")

            if st.button("Add Faculty"):
                new = pd.DataFrame([[fid,fname,sub]],
                                   columns=SCHEMA["faculty"])
                db["faculty"] = pd.concat([db["faculty"],new],ignore_index=True)
                save(db)
                st.success("Faculty Added")
                st.rerun()

        # ===== SCHEDULE =====
        elif choice == "Schedule":
            st.subheader("Schedule")
            st.dataframe(db["schedule"])

            cid = st.text_input("Class ID")
            fid = st.text_input("Faculty ID")
            room = st.text_input("Room ID")
            time = st.text_input("Time Slot")
            sub = st.text_input("Subject")

            if st.button("Add Schedule"):
                new = pd.DataFrame([[cid,fid,room,time,sub]],
                                   columns=SCHEMA["schedule"])
                db["schedule"] = pd.concat([db["schedule"],new],ignore_index=True)
                save(db)
                st.success("Schedule Added")
                st.rerun()

        elif choice == "Logout":
            st.session_state.clear()
            st.rerun()

    # ================= FACULTY =================
    elif role == "Faculty":

        if choice == "Students":
            st.title("👩‍🎓 Students")
            st.dataframe(db["students"])

        elif choice == "My Schedule":
            fid = st.session_state.link
            st.title("📅 My Schedule")
            st.dataframe(db["schedule"][db["schedule"]["faculty_id"] == fid])

        elif choice == "Marks Entry":
            st.title("Marks Entry")

            sid = st.text_input("Student ID")
            sub = st.text_input("Subject")
            m = st.number_input("Marks",0,100)

            if st.button("Save"):
                g = "A+" if m>=90 else "A" if m>=75 else "B" if m>=60 else "C" if m>=40 else "F"

                new = pd.DataFrame([[sid,sub,m,g]],
                                   columns=SCHEMA["results"])
                db["results"] = pd.concat([db["results"],new],ignore_index=True)
                save(db)
                st.success("Marks Saved")
                st.rerun()

        elif choice == "Attendance":
            st.title("Attendance")

            for _,r in db["students"].iterrows():
                status = st.selectbox(r["name"],["Present","Absent"], key=r["student_id"])

                if st.button(f"Save {r['student_id']}"):
                    new = pd.DataFrame([[r["student_id"],"C1",str(datetime.today().date()),status]],
                                       columns=SCHEMA["attendance"])
                    db["attendance"] = pd.concat([db["attendance"],new],ignore_index=True)
                    save(db)
                    st.success("Saved")

        elif choice == "Logout":
            st.session_state.clear()
            st.rerun()

    # ================= STUDENT =================
    elif role == "Student":

        sid = st.session_state.link

        if choice == "My Results":
            st.title("📄 My Results")

            df = db["results"][db["results"]["student_id"] == sid]
            st.dataframe(df)

            if st.button("Download Report"):
                file = report_pdf(sid,"Student","Course",df)

                with open(file,"rb") as f:
                    st.download_button(
                        label="📄 Download Report Card",
                        data=f.read(),
                        file_name=f"{sid}_report.pdf",
                        mime="application/pdf"
                    )

        elif choice == "Certificate":
            st.title("🎓 Certificate")

            if st.button("Download Certificate"):
                file = certificate_pdf(sid,"Student","Course")

                with open(file,"rb") as f:
                    st.download_button(
                        label="🎓 Download Certificate",
                        data=f.read(),
                        file_name=f"{sid}_certificate.pdf",
                        mime="application/pdf"
                    )

        elif choice == "Logout":
            st.session_state.clear()
            st.rerun()
