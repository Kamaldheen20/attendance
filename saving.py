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

        # ───── Certificate PDF ─────
        cert_pdf = PDF()
        cert_pdf.add_page()
        cert_pdf.set_font("Arial", size=12)
        cert_pdf.multi_cell(0, 10, f"This is to certify that {name} (Employee ID: {emp_id})")
        cert_pdf.multi_cell(0, 10, f"has successfully attended work from {start_date} to {end_date}.")
        cert_pdf.multi_cell(0, 10, f"The employee was present for {present_days} out of {total_days} working days.")
        cert_pdf.multi_cell(0, 10, f"Attendance Percentage: {percentage:.2f}%")
        cert_pdf.ln(10)
        cert_pdf.multi_cell(0, 10, f"Date of Issue: {datetime.today().strftime('%d-%m-%Y')}")
        cert_pdf.ln(10)
        cert_pdf.multi_cell(0, 10, "Authorized Signature:")

        cert_path = os.path.join(employee_folder, "certificate.pdf")
        cert_pdf.output(cert_path)

        # ───── Summary PDF ─────
        summary_pdf = PDF()
        summary_pdf.add_page()
        summary_pdf.set_font("Arial", "B", 12)
        summary_pdf.cell(0, 10, f"Attendance Sheet: {name}", ln=True)
        summary_pdf.ln(5)

        summary_pdf.set_font("Arial", "B", 11)
        summary_pdf.cell(60, 10, "Date", border=1, align='C')
        summary_pdf.cell(60, 10, "Present", border=1, align='C')
        summary_pdf.cell(60, 10, "Absent", border=1, ln=True, align='C')

        summary_pdf.set_font("Arial", "", 11)
        for _, row in group.iterrows():
            date_str = row['Date'].strftime('%d-%m-%Y')
            status = row['Status'].strip().lower()
            summary_pdf.cell(60, 10, date_str, border=1)
            if status == "present":
                summary_pdf.cell(60, 10, "P", border=1, align='C')
                summary_pdf.cell(60, 10, "", border=1, ln=True)
            else:
                summary_pdf.cell(60, 10, "", border=1)
                summary_pdf.cell(60, 10, "A", border=1, ln=True)

        summary_path = os.path.join(employee_folder, "summary.pdf")
        summary_pdf.output(summary_path)

        # ───── Email ─────
        subject = "Your Attendance Certificate and Summary"
        body = f"""Dear {name},

Please find attached your attendance certificate and summary report for the period {start_date} to {end_date}.

Best regards,
HR Department
"""
        send_email(email, subject, body, attachments=[cert_path, summary_path])
        print(f"✅ Sent to {name} ({email})")

        sent_log.append({"Name": name, "Email": email, "Certificate": cert_path, "Summary": summary_path})

    except Exception as e:
        print(f"❌ Failed for {name}: {e}")
        failed_log.append({"Name": name, "Email": email, "Error": str(e)})

# ───── Save Logs ─────
pd.DataFrame(sent_log).to_csv(os.path.join(LOG_FOLDER, "sent_log.csv"), index=False)
pd.DataFrame(failed_log).to_csv(os.path.join(LOG_FOLDER, "failed_log.csv"), index=False)

print("\n🎉 All attendance reports processed and emails sent.")
