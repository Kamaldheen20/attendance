import pandas as pd
from fpdf import FPDF
from datetime import datetime
import smtplib
from email.message import EmailMessage
import os

# ───── CONFIG ─────
CSV_PATH = r'C:\Users\academytraining\Desktop\landscape\employee_attendance_dataset.csv'
EMAIL_ADDRESS = "dheenkamal681@gmail.com"
EMAIL_PASSWORD = "frij wsub ckqe zalg"
LOG_FOLDER = 'logs'
EMPLOYEE_FOLDER = 'employee_reports'

# ───── Ensure Folders Exist ─────
for folder in [LOG_FOLDER, EMPLOYEE_FOLDER]:
    os.makedirs(folder, exist_ok=True)

# ───── Email Sending Function ─────
def send_email(to_email, subject, body, attachments=[]):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email
    msg.set_content(body)

    for file_path in attachments:
        with open(file_path, "rb") as f:
            file_data = f.read()
            msg.add_attachment(
                file_data,
                maintype="application",
                subtype="pdf",
                filename=os.path.basename(file_path),
            )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)

# ───── PDF Class ─────
class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 14)
        self.cell(0, 10, "Attendance Report", ln=True, align="C")
        self.ln(5)

# ───── Load Data ─────
df = pd.read_csv(CSV_PATH, encoding='utf-8')
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
sent_log = []
failed_log = []

# ───── Process Each Employee ─────
for emp_id, group in df.groupby('EmployeeID'):
    try:
        name = group['Name'].iloc[0]
        email = group['email'].iloc[0]
        start_date = group['Date'].min().strftime('%d-%m-%Y')
        end_date = group['Date'].max().strftime('%d-%m-%Y')

        present_days = group[group['Status'].str.lower() == 'present'].shape[0]
        total_days = group.shape[0]
        percentage = (present_days / total_days) * 100 if total_days > 0 else 0

        # ───── Create Employee Folder ─────
        safe_name = name.replace(' ', '_')
        employee_folder = os.path.join(EMPLOYEE_FOLDER, f"{emp_id}_{safe_name}")
        os.makedirs(employee_folder, exist_ok=True)

        # ───── One Combined PDF: Certificate + Summary Table ─────
        pdf = PDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        # Certificate Content
        pdf.multi_cell(0, 10, f"This is to certify that {name} (Employee ID: {emp_id})")
        pdf.multi_cell(0, 10, f"has successfully attended work from {start_date} to {end_date}.")
        pdf.multi_cell(0, 10, f"The employee was present for {present_days} out of {total_days} working days.")
        pdf.multi_cell(0, 10, f"Attendance Percentage: {percentage:.2f}%")
        pdf.ln(10)
       
    

        # Attendance Table Heading
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, f"Daily Attendance Sheet: {name}", ln=True)
        pdf.ln(5)

        # Table Header
        pdf.set_font("Arial", "B", 11)
        pdf.cell(60, 10, "Date", border=1, align='C')
        pdf.cell(60, 10, "Present", border=1, align='C')
        pdf.cell(60, 10, "Absent", border=1, ln=True, align='C')

        # Table Rows
        pdf.set_font("Arial", "", 11)
        for _, row in group.iterrows():
            date_str = row['Date'].strftime('%d-%m-%Y')
            status = row['Status'].strip().lower()
            pdf.cell(60, 10, date_str, border=1)
            if status == "present":
                pdf.cell(60, 10, "P", border=1, align='C')
                pdf.cell(60, 10, "", border=1, ln=True)
            else:
                pdf.cell(60, 10, "", border=1)
                pdf.cell(60, 10, "A", border=1, ln=True)

        # Save final PDF
        final_path = os.path.join(employee_folder, "attendance_report_combined.pdf")
        pdf.output(final_path)

        # ───── Email ─────
        subject = "Your Attendance Certificate and Summary"
        body = f"""Dear {name},

Please find attached your attendance certificate and summary report for the period {start_date} to {end_date}.

Best regards,
HR Department
"""
        send_email(email, subject, body, attachments=[final_path])
        print(f"✅ Sent to {name} ({email})")

        sent_log.append({
            "Name": name,
            "Email": email,
            "Combined_Report": final_path
        })

    except Exception as e:
        print(f"❌ Failed for {name}: {e}")
        failed_log.append({
            "Name": name,
            "Email": email,
            "Error": str(e)
        })

# ───── Save Logs ─────
pd.DataFrame(sent_log).to_csv(os.path.join(LOG_FOLDER, "sent_log.csv"), index=False)
pd.DataFrame(failed_log).to_csv(os.path.join(LOG_FOLDER, "failed_log.csv"), index=False)

print("\n🎉 All attendance reports created with continuous tables and emails sent.")
