import streamlit as st
import pandas as pd
import os
from reportlab.platypus 
import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Student ERP System", layout="wide")

FILE_NAME = "student_performance.xlsx"

# =========================
# CREATE FILE IF NOT EXISTS
# =========================
def create_file():
    if not os.path.exists(FILE_NAME):
        df_students = pd.DataFrame(columns=["ID", "Name", "Course", "Marks"])
        df_users = pd.DataFrame([["admin", "admin123"]], columns=["username", "password"])

        with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
            df_students.to_excel(writer, sheet_name="students", index=False)
            df_users.to_excel(writer, sheet_name="users", index=False)

create_file()

# =========================
# LOAD DATA
# =========================
def load_data():
    df_students = pd.read_excel(FILE_NAME, sheet_name="students")
    df_users = pd.read_excel(FILE_NAME, sheet_name="users")
    return df_students, df_users

def save_students(df):
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="students", index=False)

# =========================
# LOGIN FUNCTION
# =========================
def login(username, password, users_df):
    user = users_df[(users_df["username"] == username) &
                    (users_df["password"] == password)]
    return not user.empty

# =========================
# PDF REPORT
# =========================
def generate_pdf(dataframe):
    file_path = "report.pdf"
    doc = SimpleDocTemplate(file_path)

    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Student Performance Report", styles["Title"]))
    elements.append(Spacer(1, 12))

    data = [dataframe.columns.tolist()] + dataframe.values.tolist()

    table = Table(data)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.grey),
        ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
        ("GRID", (0,0), (-1,-1), 0.5, colors.black)
    ]))

    elements.append(table)
    doc.build(elements)

    return file_path

# =========================
# MAIN APP
# =========================
df_students, df_users = load_data()

st.title("🎓 Student ERP System (Streamlit + Excel)")

# =========================
# SESSION STATE LOGIN
# =========================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# =========================
# LOGIN PAGE
# =========================
if not st.session_state.logged_in:
    st.subheader("Login Page")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if login(username, password, df_users):
            st.session_state.logged_in = True
            st.success("Login Successful")
            st.rerun()
        else:
            st.error("Invalid Credentials")

# =========================
# DASHBOARD
# =========================
else:
    menu = st.sidebar.selectbox(
        "Menu",
        ["Add Student", "View Students", "Update Student", "Delete Student", "Reports", "Logout"]
    )

    # -------------------------
    # ADD STUDENT
    # -------------------------
    if menu == "Add Student":
        st.subheader("➕ Add Student")

        id = st.number_input("ID", step=1)
        name = st.text_input("Name")
        course = st.text_input("Course")
        marks = st.number_input("Marks", step=1)

        if st.button("Add"):
            new_data = pd.DataFrame([[id, name, course, marks]],
                                    columns=["ID", "Name", "Course", "Marks"])
            df_students = pd.concat([df_students, new_data], ignore_index=True)
            save_students(df_students)
            st.success("Student Added Successfully")

    # -------------------------
    # VIEW STUDENTS
    # -------------------------
    elif menu == "View Students":
        st.subheader("📋 Student Records")
        st.dataframe(df_students)

    # -------------------------
    # UPDATE STUDENT
    # -------------------------
    elif menu == "Update Student":
        st.subheader("✏️ Update Student")

        sid = st.number_input("Enter Student ID", step=1)

        if st.button("Search"):
            student = df_students[df_students["ID"] == sid]

            if not student.empty:
                name = st.text_input("Name", student.iloc[0]["Name"])
                course = st.text_input("Course", student.iloc[0]["Course"])
                marks = st.number_input("Marks", value=int(student.iloc[0]["Marks"]))

                if st.button("Update"):
                    df_students.loc[df_students["ID"] == sid, ["Name", "Course", "Marks"]] = [name, course, marks]
                    save_students(df_students)
                    st.success("Updated Successfully")
            else:
                st.error("Student Not Found")

    # -------------------------
    # DELETE STUDENT
    # -------------------------
    elif menu == "Delete Student":
        st.subheader("❌ Delete Student")

        sid = st.number_input("Enter Student ID", step=1)

        if st.button("Delete"):
            df_students = df_students[df_students["ID"] != sid]
            save_students(df_students)
            st.success("Deleted Successfully")

    # -------------------------
    # REPORTS
    # -------------------------
    elif menu == "Reports":
        st.subheader("📊 Reports")

        st.bar_chart(df_students.set_index("Name")["Marks"])

        if st.button("Generate PDF"):
            file = generate_pdf(df_students)
            st.success("PDF Generated: report.pdf")

    # -------------------------
    # LOGOUT
    # -------------------------
    elif menu == "Logout":
        st.session_state.logged_in = False
        st.success("Logged Out")
        st.rerun()
