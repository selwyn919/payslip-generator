# import pandas as pd
# from fpdf import FPDF
# import os

# def validate_excel_path(file_path: str) -> bool:
#     """Validate if the Excel file exists and is accessible."""
#     return os.path.isfile(file_path) and os.access(file_path, os.R_OK)

# def load_employee_data(file_path: str) -> pd.DataFrame:
#     """Load employee data from Excel file with detailed error reporting."""
#     try:
#         df = pd.read_excel(file_path)
        
#         # Print all column information for debugging
#         print("\n📋 Detailed Column Information:")
#         print("-" * 40)
#         print("Column Names:", df.columns.tolist())
#         print("Column Types:\n", df.dtypes)
#         print("\nSample of Data:")
#         print(df.head())
        
#         # Clean column names
#         df.columns = df.columns.str.strip().str.title()
        
#         # Verify required columns
#         required_columns = ['Name', 'EmployeeID', 'Department', 'Salary']
#         missing_cols = [col for col in required_columns if col not in df.columns]
        
#         if missing_cols:
#             print("\n❌ Missing Required Columns:")
#             for col in missing_cols:
#                 print(f"- {col}")
#             raise ValueError("Missing required columns")
            
#         return df
        
#     except FileNotFoundError:
#         print(f"\n❌ File not found: {file_path}")
#         raise
#     except Exception as e:
#         print(f"\n❌ Error reading Excel file: {str(e)}")
#         raise

# def create_payslip(pdf: FPDF, employee_data: dict) -> None:
#     """Create a formatted payslip for an employee."""
#     pdf.add_page()
#     pdf.set_font("Arial", size=12)
    
#     # Header
#     pdf.cell(200, 10, txt="Payslip", ln=True, align='C')
#     pdf.ln(10)
    
#     # Employee details
#     pdf.cell(200, 10, txt=f"Name: {employee_data['Name']}", ln=True)
#     pdf.cell(200, 10, txt=f"Employee ID: {employee_data['EmployeeID']}", ln=True)
#     pdf.cell(200, 10, txt=f"Department: {employee_data['Department']}", ln=True)
#     pdf.cell(200, 10, txt=f"Salary: ${employee_data['Salary']:.2f}", ln=True)

# def main():
#     # Set the path to your Excel file
#     excel_file = r"C:\Users\uncommonStudent\OneDrive\Desktop\selwyn python\employees.xlsx"
    
#     # Validate file existence
#     if not validate_excel_path(excel_file):
#         exit()
    
#     try:
#         # Load employee data with detailed debugging
#         df = load_employee_data(excel_file)
        
#         # Create output folder for payslips
#         output_folder = os.path.join(os.path.dirname(excel_file), "payslips")
#         os.makedirs(output_folder, exist_ok=True)
        
#         # Generate PDF payslip for each employee
#         success_count = 0
#         total_employees = len(df)
        
#         for _, row in df.iterrows():
#             try:
#                 pdf = FPDF()
#                 create_payslip(pdf, dict(row))
                
#                 safe_name = f"{row['Name'].replace(' ', '_')}_payslip.pdf"
#                 filename = os.path.join(output_folder, safe_name)
#                 pdf.output(filename)
#                 success_count += 1
                
#             except Exception as e:
#                 print(f"\n❌ Error generating payslip for {row['Name']}: {str(e)}")
#                 continue
        
#         print(f"\n✅ Generated {success_count}/{total_employees} payslips successfully!")
        
#     except Exception as e:
#         print(f"\n❌ Fatal error: {str(e)}")

# if __name__ == "__main__":
#     main()

import pandas as pd
from fpdf import FPDF
import os

# Create properly formatted DataFrame
data = {
    'EmployeeID': ['07-2348769L43', '63-6784438G48', '03-46445832M44', '44-6889759D63', '63-8897652R44'],
    'Name': ['TINASHE WUTETE', 'LLOYD CHOGARI', 'TAFADZWA', 'DEMINIOUS', 'CARLTON SITHOLE'],
    'Department': ['HR', 'IT', 'Finance', 'Marketing', 'Sales'],
    'Salary': [50000, 55000, 48000, 52000, 58000],
    'Allowances': [100, 120, 110, 100, 130],
    'Deductions': [25, 55, 45, 34, 34]
}

df = pd.DataFrame(data)

# Display the properly formatted DataFrame
print("\n✅ Properly formatted DataFrame:")
print(df)

# Save to Excel file
excel_file = r"C:\Users\uncommonStudent\OneDrive\Desktop\selwyn python\employees.xlsx"
df.to_excel(excel_file, index=False)

print("\n✅ Data has been saved to:", excel_file)
print("✅ All required columns are present:")
print("- EmployeeID")
print("- Name")
print("- Department")
print("- Salary")
print("- Additional columns: Allowances, Deductions")

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors

def calculate_net_salary(employee_data):
    """Calculate Net Salary"""
    basic = float(employee_data['Basic Salary'])
    allowance = float(employee_data['Allowance'])
    deduction = float(employee_data['Deduction'])
    return basic + allowance - deduction

def generate_payslip(employee_data, filename):
    """Generate a single payslip PDF"""
    c = canvas.Canvas(filename, pagesize=letter)
    
    # Header
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, 750, "Employee Payslip")
    
    # Employee details
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 720, f"Employee ID: {employee_data['Employee ID']}")
    c.drawString(300, 720, f"Name: {employee_data['Name']}")
    
    # Salary breakdown
    salary_table = [
        ['Component', 'Amount'],
        ['Basic Salary', f"${float(employee_data['Basic Salary']):.2f}"],
        ['Allowance', f"${float(employee_data['Allowance']):.2f}"],
        ['Deduction', f"${float(employee_data['Deduction']):.2f}"],
        ['Net Salary', f"${float(calculate_net_salary(employee_data)):.2f}"]
    ]
    
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, 680, "Salary Breakdown:")
    c.setFont("Helvetica", 12)
    
    # Create and draw the salary table
    table = Table(salary_table, style=[
        ('GRID', (0,0), (-1,-1), 1, colors.grey),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('ALIGN', (1,0), (1,-1), 'RIGHT'),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
    ])
    
    table.wrapOn(c, 400, 300)
    table.drawOn(c, 50, 580)
    
    c.save()

    from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
import pandas as pd
import os
import qrcode  # For QR code generation
import smtplib  # For sending emails
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

def calculate_net_salary(employee_data):
    """Calculate Net Salary."""
    basic = float(employee_data.get('Basic Salary', 0))
    allowance = float(employee_data.get('Allowance', 0))
    deduction = float(employee_data.get('Deduction', 0))
    return basic + allowance - deduction

def generate_payslip(employee_data, filename, logo_path=None, current_month=None):
    """Generate a single payslip PDF."""
    c = canvas.Canvas(filename, pagesize=letter)
    
    # Add company logo if provided
    if logo_path and os.path.exists(logo_path):
        c.drawImage(logo_path, 50, 740, width=100, height=50)

    # Header
    c.setFont("Helvetica-Bold", 16)
    c.drawString(200, 750, "Employee Payslip")

    # Employee details
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 700, f"Employee ID: {employee_data.get('Employee ID', 'N/A')}")
    c.drawString(300, 700, f"Name: {employee_data.get('Name', 'N/A')}")

    # Date (current month/year)
    c.setFont("Helvetica", 10)
    c.drawString(400, 700, f"Month: {current_month}")

    # Salary breakdown
    net_salary = calculate_net_salary(employee_data)
    salary_table = [
        ['Component', 'Amount'],
        ['Basic Salary', f"${float(employee_data.get('Basic Salary', 0)):.2f}"],
        ['Allowance', f"${float(employee_data.get('Allowance', 0)):.2f}"],
        ['Deduction', f"${float(employee_data.get('Deduction', 0)):.2f}"],
        ['Net Salary', f"${net_salary:.2f}"]
    ]

    table = Table(salary_table, colWidths=[200, 150])
    table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
    ]))

    # Draw the table
    table.wrapOn(c, 400, 200)
    table.drawOn(c, 50, 550)

    # QR Code for Employee ID
    qr_code = qrcode.make(employee_data.get('Employee ID', 'N/A'))
    qr_code_path = "qr_code.png"
    qr_code.save(qr_code_path)
    c.drawImage(qr_code_path, 450, 600, width=100, height=100)

    # Footer
    c.setFont("Helvetica-Oblique", 10)
    c.drawString(50, 50, "Company Name - Confidential Payslip")
    c.drawString(400, 50, "Generated by Payroll System")

    c.save()
    os.remove(qr_code_path)  # Clean up QR code image file

def send_email(pdf_filename, recipient_email, sender_email, sender_password):
    """Send an email with the PDF attached."""
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = 'Your Employee Payslip'

    # Attach the PDF
    part = MIMEBase('application', 'octet-stream')
    with open(pdf_filename, 'rb') as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(pdf_filename)}')
    msg.attach(part)

    # Send email
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())

    print(f"📧 Payslip sent to {recipient_email}")

# Load employee data from Excel
excel_file = r"C:\Users\uncommonStudent\OneDrive\Desktop\selwyn python\employees.xlsx"
df = pd.read_excel(excel_file)
df.columns = df.columns.str.strip()  # Remove leading/trailing spaces

# Print out the actual column names for troubleshooting
print("Column names in the Excel file:", df.columns)

# Validate required columns
required_columns = {'Employee ID', 'Name', 'Basic Salary', 'Allowance', 'Deduction'}
missing = required_columns - set(df.columns)
if missing:
    raise KeyError(f"Missing required columns in Excel file: {missing}")

# Output folder for PDFs
output_folder = os.path.join(os.path.dirname(excel_file), "payslips")
os.makedirs(output_folder, exist_ok=True)

# Optional: logo
logo_path = os.path.join(os.path.dirname(excel_file), "company_logo.png")

# Get current month
current_month = datetime.now().strftime("%B %Y")

# Email details
sender_email = "youremail@example.com"  # Your email address
sender_password = "yourpassword"  # Your email password or app password
recipient_email = "employeeemail@example.com"  # Replace with the actual employee's email or loop through employee emails

# Generate PDF and send via email
for _, row in df.iterrows():
    name_safe = str(row['Name']).strip().replace(' ', '_')
    pdf_filename = os.path.join(output_folder, f"{name_safe}_payslip.pdf")
    generate_payslip(row, pdf_filename, logo_path=logo_path, current_month=current_month)

    # Optionally send email with the payslip
    send_email(pdf_filename, recipient_email, sender_email, sender_password)

print("✅ Payslips generated and sent successfully.")

def generate_payslip(employee_data, filename, logo_path=None, current_month=None):
    """Generate a decorated payslip PDF with company branding."""
    c = canvas.Canvas(filename, pagesize=letter)
    
    # Company Header with bold font
    c.setFont("Helvetica-Bold", 20)
    c.drawString(50, 750, "PHONIX MEDIA")
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, 730, "Employee Payslip")
    
    # Decorative line under header
    c.setLineWidth(2)
    c.line(50, 720, 550, 720)
    
    # Employee details with decorative elements
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 700, f"Employee ID: {employee_data.get('EmployeeID', 'N/A')}")
    c.drawString(300, 700, f"Name: {employee_data.get('Name', 'N/A')}")
    
    # Date with decorative frame
    c.setFont("Helvetica", 10)
    c.drawString(400, 700, f"Month: {current_month}")
    
    # Salary breakdown table with styling
    net_salary = calculate_net_salary(employee_data)
    salary_table = [
        ['Component', 'Amount'],
        ['Salary', f"${float(employee_data.get('Salary', 0)):.2f}"],
        ['Allowance', f"${float(employee_data.get('Allowances', 0)):.2f}"],
        ['Deduction', f"${float(employee_data.get('Deductions', 0)):.2f}"],
        ['Net Salary', f"${net_salary:.2f}"]
    ]
    
    table = Table(salary_table, colWidths=[200, 150])
    table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('TEXTCOLOR', (1, -1), (1, -1), colors.blue)  # Highlight net salary
    ]))
    
    # Draw the table with decorative border
    table.wrapOn(c, 400, 200)
    table.drawOn(c, 50, 550)
    
    # Decorative footer with company information
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 50, "PHONIX MEDIA")
    c.setFont("Helvetica", 10)
    c.drawString(50, 40, "Digital Media Solutions")
    c.drawString(400, 50, "Generated by Payroll System")
    c.drawString(400, 40, f"Date: {datetime.now().strftime('%d %B %Y')}")
    
    # Decorative border around the payslip
    c.setLineWidth(1)
    c.rect(40, 30, 520, 680, stroke=1, fill=0)
    
    c.save()
    return True

import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import qrcode
from datetime import datetime
import time

def calculate_net_salary(employee_data):
    """Calculate Net Salary."""
    salary = float(employee_data.get('Salary', 0))
    allowance = float(employee_data.get('Allowances', 0))
    deduction = float(employee_data.get('Deductions', 0))
    return salary + allowance - deduction

def generate_payslip(employee_data, filename, logo_path=None, current_month=None):
    """Generate a decorated payslip PDF with company branding."""
    c = canvas.Canvas(filename, pagesize=letter)
    
    # Company Header with blue color
    c.setFont("Helvetica-Bold", 20)
    c.setFillColor(colors.blue)
    c.drawString(50, 750, "PHONIX MEDIA")
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, 730, "Employee Payslip")
    
    # Decorative line under header
    c.setLineWidth(2)
    c.line(50, 720, 550, 720)
    
    # Employee details with decorative elements
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, 700, f"Employee ID: {employee_data.get('EmployeeID', 'N/A')}")
    c.drawString(300, 700, f"Name: {employee_data.get('Name', 'N/A')}")
    
    # Date with decorative frame
    c.setFont("Helvetica", 10)
    c.drawString(400, 700, f"Month: {current_month}")
    
    # Salary breakdown table with styling
    net_salary = calculate_net_salary(employee_data)
    salary_table = [
        ['Component', 'Amount'],
        ['Salary', f"${float(employee_data.get('Salary', 0)):.2f}"],
        ['Allowance', f"${float(employee_data.get('Allowances', 0)):.2f}"],
        ['Deduction', f"${float(employee_data.get('Deductions', 0)):.2f}"],
        ['Net Salary', f"${net_salary:.2f}"]
    ]
    
    table = Table(salary_table, colWidths=[200, 150])
    table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('TEXTCOLOR', (1, -1), (1, -1), colors.blue)  # Highlight net salary
    ]))
    
    # Draw the table with decorative border
    table.wrapOn(c, 400, 200)
    table.drawOn(c, 50, 550)
    
    # Decorative footer with company information
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.blue)
    c.drawString(50, 50, "PHONIX MEDIA")
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 10)
    c.drawString(50, 40, "Digital Media Solutions")
    c.drawString(400, 50, "Generated by Payroll System")
    c.drawString(400, 40, f"Date: {datetime.now().strftime('%d %B %Y')}")
    
    # Decorative border around the payslip
    c.setLineWidth(1)
    c.rect(40, 30, 520, 680, stroke=1, fill=0)
    
    c.save()
    return True

def send_email(pdf_filename, recipient_email, sender_email, sender_password):
    """Send an email with the PDF attached."""
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = 'Your Employee Payslip'
        
        # Attach PDF
        with open(pdf_filename, 'rb') as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 
                           f'attachment; filename={os.path.basename(pdf_filename)}')
            msg.attach(part)
        
        # Send email with retry logic
        max_retries = 3
        for attempt in range(max_retries):
            try:
                with smtplib.SMTP('smtp.gmail.com', 587) as server:
                    server.starttls()
                    server.login(sender_email, sender_password)
                    server.sendmail(sender_email, recipient_email, msg.as_string())
                print(f"✅ Email sent successfully to {recipient_email}")
                return True
            except smtplib.SMTPAuthenticationError:
                print(f"❌ Authentication failed for {recipient_email}")
                return False
            except smtplib.SMTPException as e:
                if attempt < max_retries - 1:
                    print(f"⚠️ Retry {attempt + 1}/{max_retries} for {recipient_email}: {str(e)}")
                    time.sleep(2)  # Wait before retrying
                else:
                    print(f"❌ Failed to send email to {recipient_email} after {max_retries} attempts: {str(e)}")
                    return False
    except Exception as e:
        print(f"❌ Unexpected error sending email to {recipient_email}: {str(e)}")
        return False

def main():
    # Set the path to your Excel file
    excel_file = r"C:\Users\uncommonStudent\OneDrive\Desktop\selwyn python\employees.xlsx"
    
    # Validate file existence
    if not os.path.exists(excel_file):
        print(f"❌ File not found: {excel_file}")
        return
    
    try:
        # Load employee data
        df = pd.read_excel(excel_file)
        df.columns = df.columns.str.strip()
        
        # Validate required columns
        required_columns = {'EmployeeID', 'Name', 'Salary', 'Allowances', 'Deductions'}
        missing = required_columns - set(df.columns)
        if missing:
            raise KeyError(f"Missing required columns in Excel file: {missing}")
        
        # Output folder for PDFs
        output_folder = os.path.join(os.path.dirname(excel_file), "payslips")
        os.makedirs(output_folder, exist_ok=True)
        
        # Optional: logo
        logo_path = os.path.join(os.path.dirname(excel_file), "company_logo.png")
        
        # Get current month
        current_month = datetime.now().strftime("%B %Y")
        
        # Email details
        sender_email = "your-email@gmail.com"  # Your email address
        sender_password = "your-password"  # Your email password or app password
        
        # Generate PDF and send via email
        for _, row in df.iterrows():
            name_safe = str(row['Name']).strip().replace(' ', '_')
            pdf_filename = os.path.join(output_folder, f"{name_safe}_payslip.pdf")
            generate_payslip(row, pdf_filename, logo_path=logo_path, current_month=current_month)
            
            # Send email with the payslip
            if 'Email' in row:
                recipient_email = row['Email']
                send_email(pdf_filename, recipient_email, sender_email, sender_password)
            else:
                print(f"❌ Email not found for {row['Name']}")
        
        print("✅ Payslips generated and sent successfully.")
        
    except Exception as e:
        print(f"❌ Fatal error: {str(e)}")

if __name__ == "__main__":
    main()

import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
import qrcode
from datetime import datetime
import time

def calculate_net_salary(employee_data):
    """Calculate Net Salary."""
    salary = float(employee_data.get('Salary', 0))
    allowance = float(employee_data.get('Allowances', 0))
    deduction = float(employee_data.get('Deductions', 0))
    return salary + allowance - deduction

def generate_payslip(employee_data, filename, logo_path=None, current_month=None):
    """Generate a decorated payslip PDF with company branding."""
    try:
        c = canvas.Canvas(filename, pagesize=letter)
        
        # Company Header with blue color
        c.setFont("Helvetica-Bold", 20)
        c.setFillColor(colors.blue)
        c.drawString(50, 750, "PHONIX MEDIA")
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, 730, "Employee Payslip")
        
        # Decorative line under header
        c.setLineWidth(2)
        c.line(50, 720, 550, 720)
        
        # Employee details with decorative elements
        c.setFont("Helvetica-Bold", 12)
        c.drawString(50, 700, f"Employee ID: {employee_data.get('EmployeeID', 'N/A')}")
        c.drawString(300, 700, f"Name: {employee_data.get('Name', 'N/A')}")
        
        # Date with decorative frame
        c.setFont("Helvetica", 10)
        c.drawString(400, 700, f"Month: {current_month}")
        
        # Salary breakdown table with styling
        net_salary = calculate_net_salary(employee_data)
        salary_table = [
            ['Component', 'Amount'],
            ['Salary', f"${float(employee_data.get('Salary', 0)):.2f}"],
            ['Allowance', f"${float(employee_data.get('Allowances', 0)):.2f}"],
            ['Deduction', f"${float(employee_data.get('Deductions', 0)):.2f}"],
            ['Net Salary', f"${net_salary:.2f}"]
        ]
        
        table = Table(salary_table, colWidths=[200, 150])
        table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('TEXTCOLOR', (1, -1), (1, -1), colors.blue)  # Highlight net salary
        ]))
        
        # Draw the table with decorative border
        table.wrapOn(c, 400, 200)
        table.drawOn(c, 50, 550)
        
        # Decorative footer with company information
        c.setFont("Helvetica-Bold", 12)
        c.setFillColor(colors.blue)
        c.drawString(50, 50, "PHONIX MEDIA")
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 10)
        c.drawString(50, 40, "Digital Media Solutions")
        c.drawString(400, 50, "Generated by Payroll System")
        c.drawString(400, 40, f"Date: {datetime.now().strftime('%d %B %Y')}")
        
        # Decorative border around the payslip
        c.setLineWidth(1)
        c.rect(40, 30, 520, 680, stroke=1, fill=0)
        
        c.save()
        return True
        
    except Exception as e:
        print(f"❌ Error generating payslip: {str(e)}")
        return False

def send_email(pdf_filename, recipient_email, sender_email, sender_password):
    """Send an email with the PDF attached."""
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = 'Your Employee Payslip'
        
        # Attach PDF
        with open(pdf_filename, 'rb') as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 
                           f'attachment; filename={os.path.basename(pdf_filename)}')
            msg.attach(part)
        
        # Send email with retry logic and debug info
        max_retries = 3
        for attempt in range(max_retries):
            try:
                with smtplib.SMTP('smtp.gmail.com', 587) as server:
                    server.starttls()
                    server.ehlo()
                    server.login(sender_email, sender_password)
                    server.sendmail(sender_email, recipient_email, msg.as_string())
                print(f"✅ Email sent successfully to {recipient_email}")
                return True
            except smtplib.SMTPAuthenticationError:
                print(f"❌ Authentication failed for {recipient_email}. "
                      "Please check your email and password.")
                return False
            except smtplib.SMTPException as e:
                if attempt < max_retries - 1:
                    print(f"⚠️ Retry {attempt + 1}/{max_retries} for {recipient_email}: {str(e)}")
                    time.sleep(2)  # Wait before retrying
                else:
                    print(f"❌ Failed to send email to {recipient_email} after {max_retries} attempts: {str(e)}")
                    return False
    except Exception as e:
        print(f"❌ Unexpected error sending email to {recipient_email}: {str(e)}")
        return False

def main():
    # Set the path to your Excel file
    excel_file = r"C:\Users\uncommonStudent\OneDrive\Desktop\selwyn python\employees.xlsx"
    
    # Validate file existence
    if not os.path.exists(excel_file):
        print(f"❌ File not found: {excel_file}")
        return
    
    try:
        # Load employee data
        df = pd.read_excel(excel_file)
        df.columns = df.columns.str.strip()
        
        # Validate required columns
        required_columns = {'EmployeeID', 'Name', 'Salary', 'Allowances', 'Deductions'}
        missing = required_columns - set(df.columns)
        if missing:
            raise KeyError(f"Missing required columns in Excel file: {missing}")
        
        # Output folder for PDFs
        output_folder = os.path.join(os.path.dirname(excel_file), "payslips")
        os.makedirs(output_folder, exist_ok=True)
        
        # Optional: logo
        logo_path = os.path.join(os.path.dirname(excel_file), "company_logo.png")
        
        # Get current month
        current_month = datetime.now().strftime("%B %Y")
        
        # Email details
        sender_email = "your_email@gmail.com"  # Your email address
        sender_password = "app_password"  # Your app password
        
        # Generate PDF and send via email
        success_count = 0
        total_employees = len(df)
        
        for _, row in df.iterrows():
            try:
                name_safe = str(row['Name']).strip().replace(' ', '_')
                pdf_filename = os.path.join(output_folder, f"{name_safe}_payslip.pdf")
                
                # Generate payslip
                if generate_payslip(row, pdf_filename, logo_path=logo_path, current_month=current_month):
                    # Send email with the payslip
                    if 'Email' in row:
                        recipient_email = row['Email']
                        if send_email(pdf_filename, recipient_email, sender_email, sender_password):
                            success_count += 1
                    else:
                        print(f"❌ Email not found for {row['Name']}")
                else:
                    print(f"❌ Failed to generate payslip for {row['Name']}")
                    
            except Exception as e:
                print(f"❌ Error processing {row['Name']}: {str(e)}")
                continue
        
        print(f"\n✅ Generated {success_count}/{total_employees} payslips successfully!")
        
    except Exception as e:
        print(f"❌ Fatal error: {str(e)}")

if __name__ == "__main__":
    main()
