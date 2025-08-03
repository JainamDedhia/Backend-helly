from fpdf import FPDF, XPos, YPos
from num2words import num2words
import pandas as pd
import os

class PremiumPayslipPDF:
    def __init__(self, employee_row, output_folder="generated_payslips", logo_path="Viral-Photoroom.png", month="December", year="2024", date=None):
        self.employee_row = employee_row
        self.output_folder = output_folder
        self.logo_path = logo_path
        self.pdf = FPDF(unit="mm", format="A4")
        self.pdf.add_page()
        os.makedirs(self.output_folder, exist_ok=True)
        self.month = month
        self.year = year
        self.date = date

        # Minimalist greyscale palette
        self.black = (0, 0, 0)
        self.white = (255, 255, 255)
        self.light_gray = (245, 245, 245)
        self.medium_gray = (220, 220, 220)
        self.dark_gray = (50, 50, 50)
        self.section_gray = (235, 235, 235)
        self.lime_green = (144, 238, 144)

    def add_gradient_background(self):
        self.pdf.set_fill_color(*self.white)
        self.pdf.rect(0, 0, 210, 297, 'F')

    def header(self):
        self.add_gradient_background()
        self.pdf.set_fill_color(*self.light_gray)
        self.pdf.rect(0, 0, 210, 42, 'F')

        if os.path.exists(self.logo_path):
            self.pdf.image(self.logo_path, x=17, y=11, w=18)  # Logo shifted down for balance

        # Main header text, perfectly aligned and black
        self.pdf.set_xy(42, 15)
        self.pdf.set_font("helvetica", 'B', 18)
        self.pdf.set_text_color(*self.black)
        self.pdf.cell(0, 8, "Helly Consultancy Services", ln=True)

        self.pdf.set_xy(42, 25)
        self.pdf.set_font("helvetica", '', 9)
        self.pdf.cell(0, 5, "11, Shridevi apartment, 1st Floor, behind bhiwandi talkies", ln=True)

        self.pdf.set_xy(42, 30)
        self.pdf.cell(0, 5, "BHIWANDI, DIST-THANE, MAHARASHTRA, PIN-421308, MOBL: 7028630748", ln=True)

        self.pdf.set_line_width(0.8)
        self.pdf.set_draw_color(*self.medium_gray)
        self.pdf.line(10, 42, 200, 42)
        self.pdf.set_y(48)

    def payslip_title(self):
        self.pdf.set_fill_color(*self.medium_gray)
        self.pdf.rect(10, self.pdf.get_y(), 190, 15, 'F')

        self.pdf.set_font("helvetica", 'B', 14)
        self.pdf.set_text_color(*self.black)
        self.pdf.set_y(self.pdf.get_y() + 5)
        self.pdf.cell(0, 7, f"Payslip for the Month: {self.month} {self.year}", ln=True, align="C")
        self.pdf.ln(4)

    def create_section_header(self, title, bg_color=None):
        if bg_color is None:
            bg_color = self.section_gray
        current_y = self.pdf.get_y()
        self.pdf.set_fill_color(*bg_color)
        self.pdf.rect(10, current_y, 190, 8, 'F')
        self.pdf.set_font("helvetica", 'B', 10)
        self.pdf.set_text_color(*self.black)
        self.pdf.set_y(current_y + 2)
        self.pdf.set_x(15)
        self.pdf.cell(0, 6, title, ln=True)
        self.pdf.ln(3)

    def employee_summary(self):
        self.create_section_header("EMPLOYEE INFORMATION")
        box_y = self.pdf.get_y()
        self.pdf.set_fill_color(*self.white)
        self.pdf.rect(10, box_y, 190, 24, 'F')

        self.pdf.set_draw_color(*self.medium_gray)
        self.pdf.set_line_width(0.5)
        self.pdf.rect(10, box_y, 190, 24)

        self.pdf.set_y(box_y + 3)
        self.pdf.set_font("helvetica", '', 10)
        self.pdf.set_text_color(*self.black)

        name = self.employee_row["NAME"]
        mobile = str(int(self.employee_row["mobile no"])) if pd.notna(self.employee_row["mobile no"]) else "N/A"

        self.pdf.set_x(15)
        self.pdf.set_font("helvetica", 'B', 9)
        self.pdf.cell(27, 6, "Employee:")
        self.pdf.set_font("helvetica", '', 9)
        self.pdf.cell(55, 6, name)

        self.pdf.set_font("helvetica", 'B', 9)
        self.pdf.cell(22, 6, "Phone no:")
        self.pdf.set_font("helvetica", '', 9)
        self.pdf.cell(0, 6, mobile, ln=True)

        self.pdf.set_x(15)
        self.pdf.set_font("helvetica", 'B', 9)
        self.pdf.cell(27, 6, "Period:")
        self.pdf.set_font("helvetica", '', 9)
        self.pdf.cell(55, 6, f"{self.month} {self.year}")

        self.pdf.set_font("helvetica", 'B', 9)
        self.pdf.cell(22, 6, "Pay Date:")
        self.pdf.set_font("helvetica", '', 9)
        self.pdf.cell(0, 6, self.date if self.date else "N/A", ln=True)

        self.pdf.ln(6)

    def salary_details(self):
        self.create_section_header("SALARY AGGREGATION")

        table_y = self.pdf.get_y()
        self.pdf.set_fill_color(*self.medium_gray)
        self.pdf.rect(10, table_y, 190, 8, 'F')

        self.pdf.set_y(table_y + 1.5)
        self.pdf.set_font("helvetica", 'B', 10)
        self.pdf.set_text_color(*self.black)
        self.pdf.set_x(15)
        self.pdf.cell(120, 5, "COMPONENT")
        self.pdf.cell(0, 5, "AMOUNT (Rs.)", align="R")
        self.pdf.ln(8)

        components = [
            ("Basic Salary", self.employee_row['SALARY']),
            ("(+)Advance", -self.employee_row['ADVANCE']),
            ("(-)Deduction", -self.employee_row['DEDUCTION'])
        ]

        for i, (component, amount) in enumerate(components):
            row_y = self.pdf.get_y()
            bg_color = self.white if i % 2 == 0 else self.section_gray
            self.pdf.set_fill_color(*bg_color)
            self.pdf.rect(10, row_y, 190, 8, 'F')

            self.pdf.set_draw_color(*self.medium_gray)
            self.pdf.set_line_width(0.3)
            self.pdf.line(10, row_y, 200, row_y)

            self.pdf.set_y(row_y + 1.3)
            self.pdf.set_font("helvetica", '', 10)
            self.pdf.set_text_color(*self.black)
            self.pdf.set_x(15)
            self.pdf.cell(120, 5, component)

            self.pdf.set_font("helvetica", 'B', 10)
            amount_text = f"Rs.{abs(amount):,.2f}"
            if amount < 0:
                amount_text = f"-{amount_text}"
            self.pdf.cell(0, 5, amount_text, align="R")
            self.pdf.ln(7)

        self.pdf.set_draw_color(*self.dark_gray)
        self.pdf.set_line_width(0.5)
        self.pdf.line(10, self.pdf.get_y(), 200, self.pdf.get_y())
        self.pdf.ln(4)

        total = self.employee_row['NET']
        self.pdf.set_fill_color(*self.light_gray)
        row_y = self.pdf.get_y()
        self.pdf.rect(10, row_y, 190, 10, 'F')

        self.pdf.set_font("helvetica", 'B', 11)
        self.pdf.set_text_color(*self.black)
        self.pdf.set_x(15)
        self.pdf.cell(120, 8, "Net Pay (Total)")
        self.pdf.cell(0, 8, f"Rs.{total:,.2f}", align="R")
        self.pdf.ln(11)

    def net_pay_section(self):
        net = float(self.employee_row["NET"])
        box_y = self.pdf.get_y()
        self.pdf.set_fill_color(*self.lime_green)
        self.pdf.rect(10, box_y, 190, 18, 'F')

        self.pdf.set_fill_color(*self.white)
        self.pdf.rect(15, box_y + 3, 180, 11, 'F')

        self.pdf.set_draw_color(*self.lime_green)
        self.pdf.set_line_width(1)
        self.pdf.rect(15, box_y + 3, 180, 11)

        self.pdf.set_y(box_y + 6)
        self.pdf.set_font("helvetica", 'B', 13)
        self.pdf.set_text_color(*self.black)
        self.pdf.cell(0, 7, f"NET SALARY: Rs.{net:,.2f}", ln=True, align="C")
        self.pdf.ln(3)

        words_y = self.pdf.get_y()
        self.pdf.set_fill_color(*self.lime_green)
        self.pdf.rect(10, words_y, 190, 8, 'F')

        self.pdf.set_y(words_y + 2)
        self.pdf.set_font("helvetica", 'I', 9)
        self.pdf.set_text_color(*self.dark_gray)
        words = num2words(net, lang="en_IN").title()
        self.pdf.cell(0, 5, f"Amount in Words: Indian Rupee {words} Only", ln=True, align="C")
        self.pdf.ln(8)

    def footer(self):
        footer_y = self.pdf.get_y()
        self.pdf.set_fill_color(*self.light_gray)
        self.pdf.rect(0, footer_y, 210, 18, 'F')

        self.pdf.set_draw_color(*self.section_gray)
        self.pdf.set_line_width(1)
        self.pdf.line(10, footer_y + 5, 200, footer_y + 5)

        self.pdf.set_y(footer_y + 6)
        self.pdf.set_font("helvetica", 'I', 8)
        self.pdf.set_text_color(*self.dark_gray)
        self.pdf.cell(0, 4, "-- This is a system-generated document --", ln=True, align="C")

        self.pdf.set_font("helvetica", 'B', 10)
        self.pdf.set_text_color(*self.black)
        self.pdf.cell(0, 6, "Helly Consultancy Services", ln=True, align="C")

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

def process_excel_file(excel_file, output_folder="generated_payslips", month="December", year="2024", date=None) -> list:
    df = pd.read_excel(excel_file, header=4)
    df.columns = df.columns.str.strip()
    df = df[df["NAME"].notna()]

    filenames = []
    for _, row in df.iterrows():
        payslip = PremiumPayslipPDF(employee_row=row, output_folder=output_folder, month=month, year=year, date=date)
        filename = payslip.generate()
        filenames.append(filename)

    return filenames

if __name__ == "__main__":
    process_excel_file("employee_data.xlsx")
