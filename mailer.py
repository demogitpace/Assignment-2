import os
import logging
from dotenv import load_dotenv
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import timedelta

# Load environment variables
load_dotenv()

# Set up logging
logging.basicConfig(filename='email_errors.log', level=logging.ERROR,
                    format='%(asctime)s:%(levelname)s:%(message)s')

def minutes_to_hours_minutes(minutes):
    if minutes == 0:
        return "Absent"
    td = timedelta(minutes=int(minutes))
    hours, remainder = divmod(td.seconds, 3600)
    minutes, _ = divmod(remainder, 60)
    return f"{hours}h {minutes}m"

def calculate_total_attendance(student_data):
    sessions = [f"Session {i}" for i in range(1, 22)]
    total_minutes = sum(int(student_data[session]) if student_data[session] else 0 for session in sessions)
    hours, minutes = divmod(total_minutes, 60)
    return f"{hours}h {minutes}m"

def generate_html_content(student_data):
    sessions = [f"Session {i}" for i in range(1, 22)]
    
    html_content = f"""
    <html>
    <body>
    <p>Dear {student_data['Name']},</p>
    <p>Here is your attendance report for Java FSD Class being conducted by <strong>42 Learn</strong> for <strong>PACE</strong>:</p>
    <table border='1'>
        <tr><th>College</th><th>Roll Number</th><th>Name</th><th>Email</th><th>Branch</th></tr>
        <tr><td>{student_data['College']}</td><td>{student_data['Roll Number']}</td><td>{student_data['Name']}</td><td>{student_data['Email']}</td><td>{student_data['Branch']}</td></tr>
    </table><br>
    <table border='1'>
        <tr><th>Session</th><th>Duration</th></tr>
    """
    
    total_minutes = 0
    for session in sessions:
        duration = int(student_data[session]) if student_data[session] else 0
        total_minutes += duration
        formatted_duration = minutes_to_hours_minutes(duration)
        style = " style='background-color: #FFB3BA;'" if duration == 0 else ""
        html_content += f"<tr><td>{session}</td><td{style}>{formatted_duration}</td></tr>"
    
    total_attendance = calculate_total_attendance(student_data)
    html_content += f"""
    </table>
    <p>Total attendance: <strong>{total_attendance}</strong></p>
    <p>If your attendance is not tracked, kindly join from now on using the email address to which this mail is sent ({student_data['Email']}). 
    This will allow us to track your attendance going forward. Please use only this email address to join the meet.</p>
    <p>If you have any questions, please contact us at hello@join42.com</p>
    </body>
    </html>
    """
    return html_content

def send_email(sender_email, sender_password, student_data):
    recipient_email = student_data['Email']
    subject = f"Attendance Report for {student_data['Name']}"
    html_content = generate_html_content(student_data)
    
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = recipient_email
    message['Subject'] = subject
    
    message.attach(MIMEText(html_content, 'html'))
    
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, sender_password)
            server.send_message(message)
        return True
    except smtplib.SMTPException as e:
        total_attendance = calculate_total_attendance(student_data)
        logging.error(f"Failed to send email. Details:\n"
                      f"Name: {student_data['Name']}\n"
                      f"Email: {recipient_email}\n"
                      f"Total Attendance: {total_attendance}\n"
                      f"Error: {str(e)}")
        return False

def main():
    # Get email configuration from environment variables
    sender_email = os.getenv('EMAIL_USER')
    sender_password = os.getenv('EMAIL_PASSWORD')
    excel_file = os.getenv('EXCEL_FILE', 'attendance_sep4.xlsx')
    
    if not sender_email or not sender_password:
        print("Error: Email configuration not found in .env file")
        return
    
    try:
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active
    except FileNotFoundError:
        print(f"Error: Excel file '{excel_file}' not found.")
        return
    except openpyxl.utils.exceptions.InvalidFileException:
        print(f"Error: '{excel_file}' is not a valid Excel file.")
        return
    
    headers = [cell.value for cell in sheet[1]]
    
    successful_emails = 0
    failed_emails = 0

    for row in sheet.iter_rows(min_row=2, values_only=True):
        student_data = dict(zip(headers, row))
        
        if send_email(sender_email, sender_password, student_data):
            print(f"Email sent successfully to {student_data['Name']} at {student_data['Email']}")
            successful_emails += 1
        else:
            print(f"Failed to send email to {student_data['Name']} at {student_data['Email']}")
            failed_emails += 1

    print(f"\nEmail sending completed.")
    print(f"Successful emails: {successful_emails}")
    print(f"Failed emails: {failed_emails}")
    print("Check 'email_errors.log' for details on failed emails.")

if __name__ == "__main__":
    main()