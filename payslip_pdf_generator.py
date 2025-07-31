import pandas as pd
import os
from fpdf import FPDF, XPos, YPos
from num2words import num2words


class PremiumPayslipPDF:
    def __init__(self, employee_row, output_folder="generated_payslips", logo_path="Viral-Photoroom.png"):
        self.employee_row = employee_row
        self.output_folder = output_folder
        self.logo_path = logo_path
        self.pdf = FPDF(unit="mm", format="A4")
        os.makedirs(self.output_folder, exist_ok=True)
        
        # Modern Color Palette - Enhanced
        self.primary_gradient_start = (16, 42, 67)    # Deep navy blue
        self.primary_gradient_end = (25, 64, 102)     # Rich blue
        self.accent_teal = (26, 188, 156)             # Modern teal
        self.accent_coral = (255, 107, 107)           # Soft coral
        self.success_emerald = (46, 204, 113)         # Vibrant emerald
        self.warning_amber = (241, 196, 15)           # Golden amber
        self.error_crimson = (231, 76, 60)            # Deep crimson
        self.background_light = (248, 251, 255)       # Very light blue
        self.background_card = (255, 255, 255)        # Pure white
        self.text_primary = (45, 55, 72)              # Dark blue-gray
        self.text_secondary = (107, 114, 128)         # Medium gray
        self.text_light = (156, 163, 175)             # Light gray
        self.border_light = (229, 231, 235)           # Very light border
        self.shadow_color = (0, 0, 0)                 # For subtle shadows
        
    def add_modern_background(self):
        """Add a modern gradient background"""
        # Main background
        self.pdf.set_fill_color(*self.background_light)
        self.pdf.rect(0, 0, 210, 297, 'F')
        
        # Subtle accent shapes for modern look
        self.pdf.set_fill_color(240, 248, 255)
        self.pdf.ellipse(180, -20, 60, 60, 'F')
        self.pdf.ellipse(-20, 250, 80, 80, 'F')

    def header(self):
        self.add_modern_background()
        
        # Modern header with gradient effect
        self.pdf.set_fill_color(*self.primary_gradient_start)
        self.pdf.rect(0, 0, 210, 50, 'F')
        
        # Add subtle overlay for depth
        self.pdf.set_fill_color(*self.primary_gradient_end)
        self.pdf.rect(0, 0, 210, 25, 'F')
        
        # Company logo positioning
        logo_x = 20
        logo_y = 12
        logo_w = 26
        text_gap = 12
        
        if os.path.exists(self.logo_path):
            self.pdf.image(self.logo_path, x=logo_x, y=logo_y, w=logo_w)

        # Company name with modern typography
        text_x = logo_x + logo_w + text_gap
        text_y = logo_y + 2
        
        self.pdf.set_xy(text_x, text_y)
        self.pdf.set_font("helvetica", 'B', 20)
        self.pdf.set_text_color(255, 255, 255)
        self.pdf.cell(0, 8, "Helly Consultancy Services", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        # Company address with better contrast
        self.pdf.set_x(text_x)
        self.pdf.set_font("helvetica", '', 10)
        self.pdf.set_text_color(200, 220, 240)
        self.pdf.cell(0, 5, "54, MALANI BHAVAN, G/7- PAR NAKA, BHIWANDI", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.pdf.set_x(text_x)
        self.pdf.cell(0, 5, "DIST- THANE, MAHARASTRA, PIN-421308", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        # Modern accent line with gradient effect
        self.pdf.set_y(50)
        self.pdf.set_draw_color(*self.accent_teal)
        self.pdf.set_line_width(2)
        self.pdf.line(15, 50, 195, 50)
        
        self.pdf.set_y(58)

    def payslip_title(self):
        # Modern title with card-like design
        title_y = self.pdf.get_y()
        
        # Card shadow effect
        self.pdf.set_fill_color(0, 0, 0, 5)  # Very transparent black
        self.pdf.rect(12, title_y + 1, 186, 18, 'F')
        
        # Main card
        self.pdf.set_fill_color(*self.background_card)
        self.pdf.rect(10, title_y, 190, 18, 'F')
        
        # Accent border
        self.pdf.set_draw_color(*self.accent_teal)
        self.pdf.set_line_width(0.8)
        self.pdf.rect(10, title_y, 190, 18)
        
        # Title text with modern styling
        self.pdf.set_font("helvetica", 'B', 16)
        self.pdf.set_text_color(*self.primary_gradient_start)
        self.pdf.set_y(title_y + 6)
        self.pdf.cell(0, 8, "PAYSLIP FOR THE MONTH OF DECEMBER 2025", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        
        self.pdf.ln(8)

    def create_modern_section_header(self, title):
        """Create a modern section header with gradient effect"""
        current_y = self.pdf.get_y()
        
        # Gradient background simulation
        self.pdf.set_fill_color(*self.accent_teal)
        self.pdf.rect(10, current_y, 190, 10, 'F')
        
        # Overlay for gradient effect
        self.pdf.set_fill_color(255, 255, 255, 20)  # Semi-transparent white
        self.pdf.rect(10, current_y, 190, 5, 'F')
        
        # Section title with modern typography
        self.pdf.set_font("helvetica", 'B', 11)
        self.pdf.set_text_color(255, 255, 255)
        self.pdf.set_y(current_y + 2)
        self.pdf.set_x(18)
        self.pdf.cell(0, 6, title, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        self.pdf.ln(6)

    def employee_summary(self):
        self.create_modern_section_header("EMPLOYEE INFORMATION")
        
        # Modern card design with shadow
        box_y = self.pdf.get_y()
        
        # Shadow effect
        self.pdf.set_fill_color(0, 0, 0, 8)
        self.pdf.rect(12, box_y + 1, 186, 24, 'F')
        
        # Main card
        self.pdf.set_fill_color(*self.background_card)
        self.pdf.rect(10, box_y, 190, 24, 'F')
        
        # Subtle border
        self.pdf.set_draw_color(*self.border_light)
        self.pdf.set_line_width(0.5)
        self.pdf.rect(10, box_y, 190, 24)
        
        self.pdf.set_y(box_y + 6)
        self.pdf.set_font("helvetica", '', 11)
        
        name = self.employee_row["NAME"]
        mobile = str(int(self.employee_row["mobile no"])) if pd.notna(self.employee_row["mobile no"]) else "N/A"
        
        # Employee info with modern styling
        self.pdf.set_x(20)
        self.pdf.set_font("helvetica", 'B', 10)
        self.pdf.set_text_color(*self.accent_teal)
        self.pdf.cell(30, 7, "Employee:")
        self.pdf.set_font("helvetica", '', 11)
        self.pdf.set_text_color(*self.text_primary)
        self.pdf.cell(80, 7, name)
        
        self.pdf.set_font("helvetica", 'B', 10)
        self.pdf.set_text_color(*self.accent_teal)
        self.pdf.cell(25, 7, "ID:")
        self.pdf.set_font("helvetica", '', 11)
        self.pdf.set_text_color(*self.text_primary)
        self.pdf.cell(0, 7, mobile, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
        self.pdf.ln(8)

    def salary_details(self):
        self.create_modern_section_header("SALARY BREAKDOWN")
        
        # Modern table design
        table_y = self.pdf.get_y()
        
        # Table shadow
        self.pdf.set_fill_color(0, 0, 0, 8)
        self.pdf.rect(12, table_y + 1, 186, 35, 'F')
        
        # Table background
        self.pdf.set_fill_color(*self.background_card)
        self.pdf.rect(10, table_y, 190, 35, 'F')
        
        # Header with gradient
        self.pdf.set_fill_color(*self.primary_gradient_start)
        self.pdf.rect(10, table_y, 190, 10, 'F')
        
        self.pdf.set_y(table_y + 2)
        self.pdf.set_font("helvetica", 'B', 11)
        self.pdf.set_text_color(255, 255, 255)
        self.pdf.set_x(18)
        self.pdf.cell(130, 6, "COMPONENT")
        self.pdf.cell(0, 6, "AMOUNT (Rs.)", align="R")
        
        self.pdf.ln(12)
        
        # Table rows with modern styling
        components = [
            ("Basic Salary", self.employee_row['SALARY'], self.success_emerald),
            ("Advance", -self.employee_row['ADVANCE'], self.error_crimson),
            ("Deduction", -self.employee_row['DEDUCTION'], self.error_crimson)
        ]
        
        for i, (component, amount, color) in enumerate(components):
            row_y = self.pdf.get_y()
            
            # Subtle row separators
            if i > 0:
                self.pdf.set_draw_color(*self.border_light)
                self.pdf.set_line_width(0.3)
                self.pdf.line(15, row_y, 195, row_y)
            
            self.pdf.set_y(row_y + 2)
            self.pdf.set_font("helvetica", '', 11)
            self.pdf.set_text_color(*self.text_primary)
            self.pdf.set_x(18)
            self.pdf.cell(130, 6, component)
            
            # Amount with modern color coding
            self.pdf.set_text_color(*color)
            self.pdf.set_font("helvetica", 'B', 11)
            amount_text = f"Rs.{abs(amount):,.2f}"
            if amount < 0:
                amount_text = f"-{amount_text}"
            self.pdf.cell(0, 6, amount_text, align="R")
            
            self.pdf.ln(8)
        
        self.pdf.ln(6)

    def net_pay_section(self):
        net = float(self.employee_row["NET"])
        
        # Modern net pay card with gradient
        box_y = self.pdf.get_y()
        
        # Shadow effect
        self.pdf.set_fill_color(0, 0, 0, 12)
        self.pdf.rect(12, box_y + 2, 186, 26, 'F')
        
        # Gradient background
        self.pdf.set_fill_color(*self.success_emerald)
        self.pdf.rect(10, box_y, 190, 26, 'F')
        
        # Inner card with modern design
        self.pdf.set_fill_color(*self.background_card)
        self.pdf.rect(16, box_y + 4, 178, 18, 'F')
        
        # Modern border
        self.pdf.set_draw_color(*self.success_emerald)
        self.pdf.set_line_width(1.5)
        self.pdf.rect(16, box_y + 4, 178, 18)
        
        # Net pay amount with modern typography
        self.pdf.set_y(box_y + 9)
        self.pdf.set_font("helvetica", 'B', 16)
        self.pdf.set_text_color(*self.success_emerald)
        self.pdf.cell(0, 8, f"NET SALARY: Rs.{net:,.2f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        
        self.pdf.ln(8)
        
        # Amount in words with modern styling
        words_y = self.pdf.get_y()
        self.pdf.set_fill_color(250, 252, 255)
        self.pdf.rect(10, words_y, 190, 12, 'F')
        
        # Subtle border
        self.pdf.set_draw_color(*self.border_light)
        self.pdf.set_line_width(0.5)
        self.pdf.rect(10, words_y, 190, 12)
        
        self.pdf.set_y(words_y + 3)
        self.pdf.set_font("helvetica", 'I', 10)
        self.pdf.set_text_color(*self.text_secondary)
        words = num2words(net, lang="en_IN").title()
        self.pdf.cell(0, 6, f"Amount in Words: Indian Rupee {words} Only", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        
        self.pdf.ln(12)

    def footer(self):
        # Modern footer design
        footer_y = self.pdf.get_y()
        
        # Footer background with gradient
        self.pdf.set_fill_color(248, 250, 252)
        self.pdf.rect(0, footer_y, 210, 25, 'F')
        
        # Decorative accent line
        self.pdf.set_draw_color(*self.accent_teal)
        self.pdf.set_line_width(1.5)
        self.pdf.line(15, footer_y + 8, 195, footer_y + 8)
        
        self.pdf.set_y(footer_y + 12)
        
        # System generated text with modern styling
        self.pdf.set_font("helvetica", 'I', 9)
        self.pdf.set_text_color(*self.text_light)
        self.pdf.cell(0, 4, "-- This is a system-generated document --", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
        
        # Company signature with accent color
        self.pdf.set_font("helvetica", 'B', 11)
        self.pdf.set_text_color(*self.accent_teal)
        self.pdf.cell(0, 6, "SKD Design Studio", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")

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


def process_excel_file(excel_file, output_folder="generated_payslips") -> list:
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
