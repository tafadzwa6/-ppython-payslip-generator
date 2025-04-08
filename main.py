import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import formataddr
import os

# Define constants for email setup
SMTP_SERVER = 'smtp.gmail.com'
SENDER_EMAIL = 'your-email@gmail.com'  # Replace with your email
SENDER_PASSWORD = 'your-email-password'  # Use an App Password for Gmail
PAYSILP_DIR = 'payslips'

# Function to read employee data from Excel
def load_employee_data(file_path):
    try:
        # Load the Excel file into a DataFrame
        data = pd.read_excel(file_path)
        return data
    except Exception as e:
        print(f"Error loading employee data: {str(e)}")
        return None

# Function to calculate net salary
def calculate_net_salary(row):
    return row['Basic Salary'] + row['Allowances'] - row['Deductions']

# Function to generate payslip (text format for simplicity)
def generate_payslip(row):
    pdf_content = (
        f"Employee ID: {row['Employee ID']}\n"
        f"Name: {row['Name']}\n"
        f"Basic Salary: {row['Basic Salary']}\n"
        f"Allowances: {row['Allowances']}\n"
        f"Deductions: {row['Deductions']}\n"
        f"Net Salary: {calculate_net_salary(row)}\n"
    )

    # Create directory for payslips if it doesn't exist
    if not os.path.exists(PAYSILP_DIR):
        os.makedirs(PAYSILP_DIR)

    # Save to text file (You can modify this to save as a PDF if required)
    payslip_filename = f"{PAYSILP_DIR}/{row['Employee ID']}_payslip.txt"
    with open(payslip_filename, 'w') as f:
        f.write(pdf_content)

    return payslip_filename

# Function to send the payslip via email
def send_payslip_email(row, payslip_path):
    try:
        # Prepare the email
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = row['Email']
        msg['Subject'] = "Your Payslip for This Month"
        
        body = "Please find your payslip attached."
        msg.attach(MIMEText(body, 'plain'))

        # Attach the payslip file
        with open(payslip_path, 'rb') as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(payslip_path))
            msg.attach(part)

        # Connect to the email server and send the email
        server = smtplib.SMTP(SMTP_SERVER, 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, row['Email'], msg.as_string())
        server.quit()

        print(f"Payslip for {row['Name']} sent successfully.")
    except Exception as e:
        print(f"Error sending email to {row['Name']}: {str(e)}")

# Main function
def main():
    # Load the employee data from the Excel file
    file_path = 'employees.xlsx'  # Update with your Excel file path
    data = load_employee_data(file_path)

    if data is None:
        return

    # Process each employee in the data
    for index, row in data.iterrows():
        # Generate payslip and get the file path
        payslip_path = generate_payslip(row)
        # Send the payslip via email
        send_payslip_email(row, payslip_path)

if __name__ == "__main__":
    main()