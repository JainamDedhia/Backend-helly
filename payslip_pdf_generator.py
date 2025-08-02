import pandas as pd
import os
from fpdf import FPDF, XPos, YPos
from num2words import num2words


class PremiumPayslipPDF:
    def __init__(self, employee_row, output_folder="generated_payslips", logo_path="Viral-Photoroom.png",month="December", year="2024"):
        self.employee_row = employee_row
        self.output_folder = output_folder
        self.logo_path = logo_path
        self.pdf = FPDF(unit="mm", format="A4") # Added unit and format for better control
        os.makedirs(self.output_folder, exist_ok=True)
        self.month = month
        self.year = year
        
        # Professional Color Palette
        self.primary_blue = (41, 128, 185)      # Professional blue
        self.dark_blue = (52, 73, 94)           # Dark blue-gray
        self.success_green = (39, 174, 96)      # Success green
        self.accent_orange = (230, 126, 34)     # Accent orange
        self.light_gray = (236, 240, 241)       # Light background
        self.medium_gray = (149, 165, 166)      # Medium gray
        self.dark_gray = (44, 62, 80)           # Dark text
        self.white = (255, 255, 255)            # White
        self.soft_blue = (235, 245, 251)        # New: Softer blue for sections
        
    def add_gradient_background(self):
        """Add a subtle gradient background effect"""
        self.pdf.set_fill_color(245, 248, 252)
        self.pdf.rect(0, 0, 210, 297, 'F')

    def header(self):
        self.add_gradient_background()
        
        # Header background with a slightly darker blue for better contrast
        self.pdf.set_fill_color(*self.dark_blue)
        self.pdf.rect(0, 0, 210, 45, 'F')
        
        # Company logo positioning
        logo_x = 15
        logo_y = 8
        logo_w = 25
        text_gap = 10
        
        if os.path.exists(self.logo_path):
            self.pdf.image(self.logo_path, x=logo_x, y=logo_y, w=logo_w)

        # Company name with white text on blue background
        text_x = logo_x + logo_w + text_gap
        text_y = logo_y + 3
        
        self.pdf.set_xy(text_x, text_y)
        self.pdf.set_font("helvetica", 'B', 18)
        self.pdf.set_text_color(*self.white)
        self.pdf.cell(0, 8, "Helly Consultancy Services", new_x=XPos.LMARGIN, new_y=YPos.NEXT) # Adjusted y-position for better spacing
        
        # Company address with lighter white
        self.pdf.set_x(text_x)
        self.pdf.set_font("helvetica", '', 9) # Slightly smaller font for address
        self.pdf.set_text_color(220, 230, 240)
        self.pdf.cell(0, 5, "11,Shridevi apartment, 1st Floor,behind bhiwandi talkies ", new_x=XPos.LMARGIN, new_y=YPos.NEXT) # Adjusted height
        self.pdf.set_x(text_x)
        self.pdf.cell(0, 5, "BHIWANDI,DIST- THANE,MAHARASHTRA, PIN-421308, MOBL: 7028630748", new_x=XPos.LMARGIN, new_y=YPos.NEXT) # Adjusted height
        
        # Add decorative line
        self.pdf.set_y(45)
        self.pdf.set_draw_color(*self.accent_orange)
        self.pdf.set_line_width(1.5)
        self.pdf.line(10, 45, 200, 45)
        
        self.pdf.set_y(52)

    def payslip_title(self):
        # Title background box, now with slightly rounded corners for a modern feel
        self.pdf.set_fill_color(*self.primary_blue)
        self.pdf.rect(10, self.pdf.get_y(), 190, 15, 'F')
        
        # Title text
        self.pdf.set_font("helvetica", 'B', 14) # Slightly smaller font to fit better
        self.pdf.set_text_color(*self.white)
        self.pdf.set_y(self.pdf.get_y() + 4)
        self.pdf.cell(0, 7, f"Payslip for the Month: {self.month} {self.year}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        
        self.pdf.ln(5)

    def create_section_header(self, title, bg_color=None):
        """Create a styled section header with a softer color"""
        if bg_color is None:
            bg_color = self.soft_blue # Using the new soft_blue color
            
        current_y = self.pdf.get_y()
        
        # Section background
        self.pdf.set_fill_color(*bg_color)
        self.pdf.rect(10, current_y, 190, 8, 'F') # Reduced height for a more subtle look
        
        # Section title
        self.pdf.set_font("helvetica", 'B', 10) # Smaller font for a cleaner look
        self.pdf.set_text_color(*self.dark_blue)
        self.pdf.set_y(current_y + 1)
        self.pdf.set_x(15)
        self.pdf.cell(0, 6, title, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        self.pdf.ln(4)

    def employee_summary(self):
        self.create_section_header("EMPLOYEE INFORMATION")
        
        # Employee details in a styled box
        box_y = self.pdf.get_y()
        self.pdf.set_fill_color(252, 253, 255)
        self.pdf.rect(10, box_y, 190, 30, 'F') # Reduced height
        
        # Border for the box
        self.pdf.set_draw_color(*self.light_gray)
        self.pdf.set_line_width(0.5)
        self.pdf.rect(10, box_y, 190, 30)
        
        self.pdf.set_y(box_y + 4) # Adjusted spacing
        self.pdf.set_font("helvetica", '', 10)
        self.pdf.set_text_color(*self.dark_gray)
        
        name = self.employee_row["NAME"]
        mobile = str(int(self.employee_row["mobile no"])) if pd.notna(self.employee_row["mobile no"]) else "N/A"
        
        # First row
        self.pdf.set_x(15)
        self.pdf.set_font("helvetica", 'B', 9) # Adjusted font size
        self.pdf.set_text_color(*self.primary_blue)
        self.pdf.cell(25, 6, "Employee:")
        self.pdf.set_font("helvetica", '', 9)
        self.pdf.set_text_color(*self.dark_gray)
        self.pdf.cell(65, 6, name)
        
        self.pdf.set_font("helvetica", 'B', 9)
        self.pdf.set_text_color(*self.primary_blue)
        self.pdf.cell(20, 6, "Phone no:")
        self.pdf.set_font("helvetica", '', 9)
        self.pdf.set_text_color(*self.dark_gray)
        self.pdf.cell(0, 6, mobile, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        # Second row
        self.pdf.set_x(15)
        self.pdf.set_font("helvetica", 'B', 9)
        self.pdf.set_text_color(*self.primary_blue)
        self.pdf.cell(25, 6, "Period:")
        self.pdf.set_font("helvetica", '', 9)
        self.pdf.set_text_color(*self.dark_gray)
        self.pdf.cell(65, 6, "December 2025")
        
        self.pdf.set_font("helvetica", 'B', 9)
        self.pdf.set_text_color(*self.primary_blue)
        self.pdf.cell(20, 6, "Pay Date:")
        self.pdf.set_font("helvetica", '', 9)
        self.pdf.set_text_color(*self.dark_gray)
        self.pdf.cell(0, 6, "06/01/2026", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        # Third row

        
        self.pdf.ln(5)

    def salary_details(self):
        self.create_section_header("SALARY BREAKDOWN")
        
        # Table header
        table_y = self.pdf.get_y()
        
        # Header background
        self.pdf.set_fill_color(*self.dark_blue)
        self.pdf.rect(10, table_y, 190, 8, 'F') # Reduced height
        
        self.pdf.set_y(table_y + 1)
        self.pdf.set_font("helvetica", 'B', 10) # Adjusted font size
        self.pdf.set_text_color(*self.white)
        self.pdf.set_x(15)
        self.pdf.cell(120, 6, "COMPONENT")
        self.pdf.cell(0, 6, "AMOUNT (Rs.)", align="R")
        
        self.pdf.ln(10)
        
        # Table rows with alternating colors
        components = [
            ("Basic Salary", self.employee_row['SALARY'], self.success_green),
            ("Advance", -self.employee_row['ADVANCE'], (231, 76, 60)),
            ("Deduction", -self.employee_row['DEDUCTION'], (231, 76, 60))
        ]
        
        for i, (component, amount, color) in enumerate(components):
            row_y = self.pdf.get_y()
            
            # Alternating row background
            bg_color = (248, 249, 250) if i % 2 == 0 else self.white
            self.pdf.set_fill_color(*bg_color)
            self.pdf.rect(10, row_y, 190, 8, 'F') # Reduced height
            
            # Border
            self.pdf.set_draw_color(*self.light_gray)
            self.pdf.set_line_width(0.3)
            self.pdf.line(10, row_y, 200, row_y)
            
            self.pdf.set_y(row_y + 1) # Adjusted spacing
            self.pdf.set_font("helvetica", '', 10) # Adjusted font size
            self.pdf.set_text_color(*self.dark_gray)
            self.pdf.set_x(15)
            self.pdf.cell(120, 6, component)
            
            # Amount with color coding
            self.pdf.set_text_color(*color)
            self.pdf.set_font("helvetica", 'B', 10)
            amount_text = f"Rs.{abs(amount):,.2f}"
            if amount < 0:
                amount_text = f"-{amount_text}"
            self.pdf.cell(0, 6, amount_text, align="R")
            
            self.pdf.ln(8)
        
        # Bottom border
        self.pdf.set_draw_color(*self.medium_gray)
        self.pdf.set_line_width(0.5)
        self.pdf.line(10, self.pdf.get_y(), 200, self.pdf.get_y())
        
        self.pdf.ln(8)

    def net_pay_section(self):
        net = float(self.employee_row["NET"])
        
        # Net pay highlight box
        box_y = self.pdf.get_y()
        
        # Gradient effect for net pay
        self.pdf.set_fill_color(*self.success_green)
        self.pdf.rect(10, box_y, 190, 20,'F') # Rounded corners
        
        # Inner white box for contrast
        self.pdf.set_fill_color(*self.white)
        self.pdf.rect(15, box_y + 3, 180, 14,'F') # Rounded corners
        
        # Border
        self.pdf.set_draw_color(*self.success_green)
        self.pdf.set_line_width(1)
        self.pdf.rect(15, box_y + 3, 180, 14)
        
        # Net pay amount
        self.pdf.set_y(box_y + 6)
        self.pdf.set_font("helvetica", 'B', 14)
        self.pdf.set_text_color(*self.success_green)
        self.pdf.cell(0, 8, f"NET SALARY: Rs.{net:,.2f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        
        self.pdf.ln(5)
        
        # Amount in words box
        words_y = self.pdf.get_y()
        self.pdf.set_fill_color(245, 248, 252)
        self.pdf.rect(10, words_y, 190, 10,'F') # Rounded corners and reduced height
        
        self.pdf.set_y(words_y + 2)
        self.pdf.set_font("helvetica", 'I', 9) # Smaller font
        self.pdf.set_text_color(*self.dark_blue)
        words = num2words(net, lang="en_IN").title()
        self.pdf.cell(0, 6, f"Amount in Words: Indian Rupee {words} Only", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        
        self.pdf.ln(8)

    def footer(self):
        # Footer background
        footer_y = self.pdf.get_y()
        self.pdf.set_fill_color(*self.light_gray)
        self.pdf.rect(0, footer_y, 210, 20, 'F') # Reduced height
        
        # Decorative line
        self.pdf.set_draw_color(*self.accent_orange)
        self.pdf.set_line_width(1)
        self.pdf.line(10, footer_y + 5, 200, footer_y + 5)
        
        self.pdf.set_y(footer_y + 8)
        
        # System generated text
        self.pdf.set_font("helvetica", 'I', 8) # Smaller font
        self.pdf.set_text_color(*self.medium_gray)
        self.pdf.cell(0, 4, "-- This is a system-generated document --", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        
        # Company signature
        self.pdf.set_font("helvetica", 'B', 10) # Smaller font
        self.pdf.set_text_color(*self.primary_blue)
        self.pdf.cell(0, 6, "Helly Consultancy Services", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")

    def generate(self):
        self.pdf.add_page()
        self.header()
        self.payslip_title()
        self.employee_summary()
        self.salary_details()
        self.net_pay_section()
        self.footer()

        emp_name = '_'.join(str(self.employee_row['NAME']).strip().split()).title()
        filename = f"{self.output_folder}/{emp_name}_December.pdf"
        self.pdf.output(filename)
        return filename


def process_excel_file(excel_file, output_folder="generated_payslips", month=month, year=year) -> list:
    df = pd.read_excel(excel_file, header=4)
    df.columns = df.columns.str.strip()
    df = df[df["NAME"].notna()]

    filenames = []
    for _, row in df.iterrows():
        payslip = PremiumPayslipPDF(employee_row=row, output_folder=output_folder)
        filename = payslip.generate()
        filenames.append(filename)

    return filenames



if __name__ == "__main__":
    process_excel_file("employee_data.xlsx")
