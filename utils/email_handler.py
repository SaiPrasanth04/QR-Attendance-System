import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import List

def send_absent_email(to_email: str, student_name: str, date_str: str, batch: str, sender_email: str, sender_password: str, smtp_server: str, smtp_port: int, correction: bool = False):
    if correction:
        subject = f"Attendance Correction - {date_str}"
        body = f"""
Dear {student_name},

Correction: You were previously marked absent for your batch ({batch}) on {date_str}, but you are now marked PRESENT.

Best regards,
Vcodez Attendance System
"""
    else:
        subject = f"Absence Notification - {date_str}"
        body = f"""
Dear {student_name},

This is to inform you that you have been marked absent for your batch ({batch}) on {date_str}.

If you believe this is a mistake, please contact your mentor or the administration as soon as possible.

Best regards,
Vcodez Attendance System
"""
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, to_email, msg.as_string())
        return True
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")
        return False
