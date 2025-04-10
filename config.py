import pandas as pd
from fpdf import FPDF
import os

# Path to your Excel file
excel_file = r"C:\Users\uncommonStudent\OneDrive\Desktop\selwyn python\employees.xlsx"

# Read the Excel file
try:
    df = pd.read_excel(excel_file)
except FileNotFoundError:
    print(f"❌ File not found: {excel_file}")
    exit()

# Clean column names (strip spaces)
df.columns = df.columns.str.strip()

# Show actual column names (optional, for debug)
print("Columns in Excel:", df.columns.tolist())

# Required columns for payslip
required_columns = ['Name', 'EmployeeID', 'Department', 'Salary']
for col in required_columns:
    if col not in df.columns:
        print(f"❌ Missing column in Excel: '{col}'")
        exit()

# Create output folder for PDF payslips
output_folder = os.path.join(os.path.dirname(excel_file), "payslips")
os.makedirs(output_folder, exist_ok=True)

# Function to create a PDF for one employee
def generate_payslip(row):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Payslip Header
    pdf.set_font("Arial", 'B', size=16)
    pdf.cell(200, 10, txt="Employee Payslip", ln=True, align='C')
    pdf.ln(10)

    # Reset font for details
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt=f"Name: {row['Name']}", ln=True)
    pdf.cell(200, 10, txt=f"Employee ID: {row['EmployeeID']}", ln=True)
    pdf.cell(200, 10, txt=f"Department: {row['Department']}", ln=True)
    pdf.cell(200, 10, txt=f"Salary: ${row['Salary']:.2f}", ln=True)

    # Save the PDF
    filename = f"{row['Name'].replace(' ', '_')}_payslip.pdf"
    pdf.output(os.path.join(output_folder, filename))

# Generate a payslip for each employee
for index, row in df.iterrows():
    generate_payslip(row)

print("🎉 All payslips generated successfully in:", output_folder)
