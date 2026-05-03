# ================= IMPORT =================
import streamlit as st
import pandas as pd
import os
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import inch

# ================= CONFIG =================
st.set_page_config(page_title="Campus ERP", layout="wide")

FILE = "campus_flow_template.xlsx"

# ================= CLEAN =================
def clean(x):
    return str(x).strip().replace(".0","")

# ================= LOAD =================
def load():
    if not os.path.exists(FILE):
        pd.DataFrame(columns=["username","password","role","student_id","faculty_id"])\
        .to_excel(FILE, sheet_name="users", index=False)

    df = pd.read_excel(FILE, sheet_name="users")

    df = df.astype(str)
    for col in df.columns:
        df[col] = df[col].apply(clean)

    return {"users": df}

# ================= LOGIN UI =================
def set_bg():
    st.markdown("""
    <style>
    .stApp {
        background-image: url("campus.jpg");
        background-size: cover;
    }

    .box {
        background: rgba(255,255,255,0.9);
        padding: 30px;
        border-radius: 15px;
        width: 350px;
        margin: auto;
        margin-top: 120px;
        box-shadow: 0px 0px 20px rgba(0,0,0,0.3);
    }

    .title {
        text-align: center;
        font-size: 28px;
        color: #C9A227;
        font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)

# ================= PDF (PREMIUM REPORT) =================
def report_pdf(sid, df):
    file = f"{sid}_report.pdf"
    doc = SimpleDocTemplate(file, pagesize=A4)

    content = []

    title = ParagraphStyle(name="t", fontSize=24, alignment=TA_CENTER, textColor=colors.HexColor("#C9A227"))

    content.append(Paragraph("ACADEMIC REPORT CARD", title))
    content.append(Spacer(1,20))

    data = [["Subject","Marks","Grade"]]
    for _,r in df.iterrows():
        data.append([r["subject"], r["marks"], r["grade"]])

    t = Table(data)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#C9A227")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("GRID",(0,0),(-1,-1),1,colors.black)
    ]))

    content.append(t)

    def border(c,d):
        c.setStrokeColor(colors.HexColor("#C9A227"))
        c.setLineWidth(4)
        c.rect(20,20,555,800)

    doc.build(content, onFirstPage=border)
    return file

# ================= SESSION =================
if "db" not in st.session_state:
    st.session_state.db = load()

if "login" not in st.session_state:
    st.session_state.login = False

# ================= LOGIN =================
if not st.session_state.login:

    set_bg()

    st.markdown('<div class="box">', unsafe_allow_html=True)
    st.markdown('<p class="title">🏫 Campus ERP</p>', unsafe_allow_html=True)

    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    r = st.selectbox("Role", ["Admin","Faculty","Student"])

    if st.button("Login"):

        users = st.session_state.db["users"]

        # ADMIN
        if u == "admin" and p == "admin123":
            st.session_state.login = True
            st.session_state.role = "Admin"
            st.rerun()

        # USER MATCH
        match = users[
            (users["username"] == clean(u)) &
            (users["password"] == clean(p)) &
            (users["role"].str.lower() == r.lower())
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
            st.error("❌ Invalid Username / Password")

    st.markdown('</div>', unsafe_allow_html=True)

# ================= AFTER LOGIN =================
else:
    st.success(f"Welcome {st.session_state.role} 🎉")

    if st.session_state.role == "Student":
        sid = st.session_state.link

        st.subheader("📄 My Results")

        # demo empty
        df = pd.DataFrame({
            "subject":["Math","Python","DBMS"],
            "marks":[80,75,90],
            "grade":["A","A","A+"]
        })

        st.dataframe(df)

        if st.button("Download Report Card"):
            f = report_pdf(sid, df)

            with open(f,"rb") as file:
                st.download_button(
                    "📄 Download PDF",
                    data=file.read(),
                    file_name=f"{sid}_report.pdf",
                    mime="application/pdf"
                )

    if st.button("Logout"):
        st.session_state.clear()
        st.rerun()
