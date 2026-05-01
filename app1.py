import streamlit as st
import pandas as pd
import os
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# =========================
# CONFIG & FILE HANDLING
# =========================
st.set_page_config(page_title="Student ERP System", layout="wide", page_icon="🎓")

FILE_NAME = "student_performance.xlsx"

def load_data():
    """Initializes the Excel file if missing and loads data."""
    if not os.path.exists(FILE_NAME):
        df_students = pd.DataFrame(columns=["ID", "Name", "Course", "Marks"])
        df_users = pd.DataFrame([["admin", "admin123"]], columns=["username", "password"])
        save_data(df_students, df_users)
    
    df_students = pd.read_excel(FILE_NAME, sheet_name="students")
    df_users = pd.read_excel(FILE_NAME, sheet_name="users")
    return df_students, df_users

def save_data(df_students, df_users):
    """Saves both sheets to the Excel file safely."""
    with pd.ExcelWriter(FILE_NAME, engine="openpyxl") as writer:
        df_students.to_excel(writer, sheet_name="students", index=False)
        df_users.to_excel(writer, sheet_name="users", index=False)

# Load data at startup
df_students, df_users = load_data()

# =========================
# PDF GENERATION
# =========================
def generate_pdf(dataframe):
    file_path = "report.pdf"
    doc = SimpleDocTemplate(file_path)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Student Performance Report", styles["Title"]))
    elements.append(Spacer(1, 12))

    # Convert dataframe to list format for ReportLab Table
    data = [dataframe.columns.tolist()] + dataframe.values.tolist()
    table = Table(data)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
        ("GRID", (0, 0), (-1, -1), 1, colors.black)
    ]))

    elements.append(table)
    doc.build(elements)
    return file_path

# =========================
# AUTHENTICATION SESSION
# =========================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# =========================
# LOGIN PAGE
# =========================
if not st.session_state.logged_in:
    st.title("🎓 Student ERP System")
    with st.container():
        st.subheader("Login to Access Dashboard")
        user_input = st.text_input("Username")
        pass_input = st.text_input("Password", type="password")
        
        if st.button("Login"):
            # Check credentials
            auth_success = not df_users[(df_users["username"] == user_input) & 
                                       (df_users["password"] == pass_input)].empty
            if auth_success:
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("Invalid Username or Password")

# =========================
# DASHBOARD (LOGGED IN)
# =========================
else:
    st.sidebar.title(f"Welcome, Admin")
    menu = st.sidebar.selectbox(
        "Navigation",
        ["View Students", "Add Student", "Update Student", "Delete Student", "Reports", "Logout"]
    )

    if menu == "View Students":
        st.header("📋 All Student Records")
        st.dataframe(df_students, use_container_width=True)

    elif menu == "Add Student":
        st.header("➕ Add New Student")
        with st.form("add_form", clear_on_submit=True):
            sid = st.number_input("Student ID", step=1, min_value=1)
            name = st.text_input("Full Name")
            course = st.text_input("Course Name")
            marks = st.number_input("Marks", min_value=0, max_value=100)
            
            if st.form_submit_button("Submit"):
                if sid in df_students["ID"].values:
                    st.error("Student ID already exists!")
                else:
                    new_student = pd.DataFrame([[sid, name, course, marks]], columns=df_students.columns)
                    df_students = pd.concat([df_students, new_student], ignore_index=True)
                    save_data(df_students, df_users)
                    st.success("Student added successfully!")

    elif menu == "Update Student":
        st.header("✏️ Edit Student Information")
        search_id = st.number_input("Enter ID to Search", step=1, min_value=1)
        
        # Check if student exists
        student_row = df_students[df_students["ID"] == search_id]
        
        if not student_row.empty:
            idx = student_row.index[0]
            with st.form("update_form"):
                u_name = st.text_input("Name", value=df_students.at[idx, "Name"])
                u_course = st.text_input("Course", value=df_students.at[idx, "Course"])
                u_marks = st.number_input("Marks", value=int(df_students.at[idx, "Marks"]))
                
                if st.form_submit_button("Update Records"):
                    df_students.at[idx, "Name"] = u_name
                    df_students.at[idx, "Course"] = u_course
                    df_students.at[idx, "Marks"] = u_marks
                    save_data(df_students, df_users)
                    st.success("Record updated successfully!")
        else:
            st.info("Search for a valid Student ID to reveal the update form.")

    elif menu == "Delete Student":
        st.header("❌ Remove Student")
        del_id = st.number_input("Enter ID to Delete", step=1, min_value=1)
        if st.button("Delete Permanently"):
            if del_id in df_students["ID"].values:
                df_students = df_students[df_students["ID"] != del_id]
                save_data(df_students, df_users)
                st.success(f"Student ID {del_id} deleted.")
            else:
                st.error("ID not found.")

    elif menu == "Reports":
        st.header("📊 Performance Analytics")
        if not df_students.empty:
            st.bar_chart(df_students.set_index("Name")["Marks"])
            
            if st.button("Export to PDF"):
                pdf_path = generate_pdf(df_students)
                with open(pdf_path, "rb") as f:
                    st.download_button("Download PDF Report", f, file_name="Student_Report.pdf")
        else:
            st.warning("No data available to generate reports.")

    elif menu == "Logout":
        st.session_state.logged_in = False
        st.rerun()
