import pandas as pd
import os
from fpdf import FPDF, XPos, YPos
from num2words import num2words


class PremiumPayslipPDF:
    def __init__(self, employee_row, output_folder="generated_payslips", logo_path="Viral-Photoroom.png", month="December", year="2024", date=None):
        self.employee_row = employee_row
        self.output_folder = output_folder
        self.logo_path = logo_path
        self.pdf = FPDF(unit="mm", format="A4")
        os.makedirs(self.output_folder, exist_ok=True)
        self.month = month
        self.year = year
        self.date = date

        # Updated Grayscale Theme
        self.primary_blue = (60, 60, 60)
        self.dark_blue = (40, 40, 40)
        self.success_green = (30, 130, 76)
        self.accent_orange = (100, 100, 100)
        self.light_gray = (245, 245, 245)
        self.medium_gray = (180, 180, 180)
        self.dark_gray = (60, 60, 60)
        self.white = (255, 255, 255)
        self.soft_blue = (245, 245, 245)

    def add_gradient_background(self):
        self.pdf.set_fill_color(*self.light_gray)
        self.pdf.rect(0, 0, 210, 297, 'F')

    def header(self):
        self.add_gradient_background()

        self.pdf.set_fill_color(*self.dark_blue)
        self.pdf.rect(0, 0, 210, 45, 'F')

        logo_x = 15
        logo_y = 8
        logo_w = 25
        text_gap = 10

        if os.path.exists(self.logo_path):
            self.pdf.image(self.logo_path, x=logo_x, y=logo_y, w=logo_w)

        text_x = logo_x + logo_w + text_gap
        text_y = logo_y + 3

        self.pdf.set_xy(text_x, text_y)
        self.pdf.set_font("helvetica", 'B', 18)
        self.pdf.set_text_color(*self.white)
        self.pdf.cell(0, 8, "Helly Consultancy Services", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        self.pdf.set_x(text_x)
        self.pdf.set_font("helvetica", '', 9)
        self.pdf.set_text_color(220, 220, 220)
        self.pdf.cell(0, 5, "11, Shridevi apartment, 1st Floor, behind bhiwandi talkies", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.pdf.set_x(text_x)
        self.pdf.cell(0, 5, "BHIWANDI, DIST- THANE, MAHARASHTRA, PIN-421308, MOBL: 7028630748", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        self.pdf.set_y(45)
        self.pdf.set_draw_color(*self.accent_orange)
        self.pdf.set_line_width(1.5)
        self.pdf.line(10, 45, 200, 45)

        self.pdf.set_y(52)

    def payslip_title(self):
        self.pdf.set_fill_color(*self.primary_blue)
        self.pdf.rect(10, self.pdf.get_y(), 190, 15, 'F')

        self.pdf.set_font("helvetica", 'B', 14)
        self.pdf.set_text_color(*self.white)
        self.pdf.set_y(self.pdf.get_y() + 4)
        self.pdf.cell(0, 7, f"Payslip for the Month: {self.month} {self.year}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")

        self.pdf.ln(5)

    def create_section_header(self, title, bg_color=None):
        if bg_color is None:
            bg_color = self.soft_blue

        current_y = self.pdf.get_y()
        self.pdf.set_fill_color(*bg_color)
        self.pdf.rect(10, current_y, 190, 8, 'F')

        self.pdf.set_font("helvetica", 'B', 10)
        self.pdf.set_text_color(*self.dark_blue)
        self.pdf.set_y(current_y + 1)
        self.pdf.set_x(15)
        self.pdf.cell(0, 6, title, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.pdf.ln(4)

    def employee_summary(self):
        self.create_section_header("EMPLOYEE INFORMATION")

        box_y = self.pdf.get_y()
        self.pdf.set_fill_color(252, 253, 255)
        self.pdf.rect(10, box_y, 190, 30, 'F')
        self.pdf.set_draw_color(*self.light_gray)
        self.pdf.set_line_width(0.5)
        self.pdf.rect(10, box_y, 190, 30)

        self.pdf.set_y(box_y + 4)
        self.pdf.set_font("helvetica", '', 10)
        self.pdf.set_text_color(*self.dark_gray)

        name = self.employee_row["NAME"]
        mobile = str(int(self.employee_row["mobile no"])) if pd.notna(self.employee_row["mobile no"]) else "N/A"

        self.pdf.set_x(15)
        self.pdf.set_font("helvetica", 'B', 9)
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

        self.pdf.set_x(15)
        self.pdf.set_font("helvetica", 'B', 9)
        self.pdf.set_text_color(*self.primary_blue)
        self.pdf.cell(25, 6, "Period:")
        self.pdf.set_font("helvetica", '', 9)
        self.pdf.set_text_color(*self.dark_gray)
        period = f"{self.month} {self.year}"
        pay_date = self.date if self.date else "N/A"
        self.pdf.cell(65, 6, period)

        self.pdf.set_font("helvetica", 'B', 9)
        self.pdf.set_text_color(*self.primary_blue)
        self.pdf.cell(20, 6, "Pay Date:")
        self.pdf.set_font("helvetica", '', 9)
        self.pdf.set_text_color(*self.dark_gray)
        self.pdf.cell(0, 6, pay_date, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        self.pdf.ln(5)

    def salary_details(self):
        self.create_section_header("SALARY BREAKDOWN")

        table_y = self.pdf.get_y()
        self.pdf.set_fill_color(80, 80, 80)
        self.pdf.rect(10, table_y, 190, 8, 'F')

        self.pdf.set_y(table_y + 1)
        self.pdf.set_font("helvetica", 'B', 10)
        self.pdf.set_text_color(*self.white)
        self.pdf.set_x(15)
        self.pdf.cell(120, 6, "COMPONENT")
        self.pdf.cell(0, 6, "AMOUNT (Rs.)", align="R")
        self.pdf.ln(10)

        salary = float(self.employee_row['SALARY'])
        advance = float(self.employee_row['ADVANCE'])
        deduction = float(self.employee_row['DEDUCTION'])
        total = salary - advance - deduction

        components = [
            ("Basic Salary", salary),
            ("Advance", -advance),
            ("Deduction", -deduction),
        ]

        for i, (component, amount) in enumerate(components):
            row_y = self.pdf.get_y()
            self.pdf.set_fill_color(250, 250, 250 if i % 2 == 0 else 245)
            self.pdf.rect(10, row_y, 190, 8, 'F')

            self.pdf.set_draw_color(220, 220, 220)
            self.pdf.set_line_width(0.3)
            self.pdf.line(10, row_y, 200, row_y)

            self.pdf.set_y(row_y + 1)
            self.pdf.set_font("helvetica", '', 10)
            self.pdf.set_text_color(60, 60, 60)
            self.pdf.set_x(15)
            self.pdf.cell(120, 6, component)

            self.pdf.set_font("helvetica", 'B', 10)
            self.pdf.set_text_color(60, 60, 60 if amount >= 0 else (200, 30, 30))
            amount_text = f"Rs.{abs(amount):,.2f}"
            if amount < 0:
                amount_text = f"-{amount_text}"
            self.pdf.cell(0, 6, amount_text, align="R")
            self.pdf.ln(8)

        self.pdf.set_draw_color(180, 180, 180)
        self.pdf.set_line_width(0.5)
        self.pdf.line(10, self.pdf.get_y(), 200, self.pdf.get_y())

        self.pdf.set_y(self.pdf.get_y() + 2)
        self.pdf.set_font("helvetica", 'B', 11)
        self.pdf.set_text_color(20, 20, 20)
        self.pdf.set_x(15)
        self.pdf.cell(120, 7, "Total")
        self.pdf.cell(0, 7, f"Rs.{total:,.2f}", align="R")
        self.pdf.ln(10)

    def net_pay_section(self):
        net = float(self.employee_row["NET"])
        box_y = self.pdf.get_y()

        self.pdf.set_fill_color(*self.success_green)
        self.pdf.rect(10, box_y, 190, 20, 'F')

        self.pdf.set_fill_color(*self.white)
        self.pdf.rect(15, box_y + 3, 180, 14, 'F')

        self.pdf.set_draw_color(*self.success_green)
        self.pdf.set_line_width(1)
        self.pdf.rect(15, box_y + 3, 180, 14)

        self.pdf.set_y(box_y + 6)
        self.pdf.set_font("helvetica", 'B', 14)
        self.pdf.set_text_color(*self.success_green)
        self.pdf.cell(0, 8, f"NET SALARY: Rs.{net:,.2f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")

        self.pdf.ln(5)

        words_y = self.pdf.get_y()
        self.pdf.set_fill_color(*self.soft_blue)
        self.pdf.rect(10, words_y, 190, 10, 'F')

        self.pdf.set_y(words_y + 2)
        self.pdf.set_font("helvetica", 'I', 9)
        self.pdf.set_text_color(*self.dark_blue)
        words = num2words(net, lang="en_IN").title()
        self.pdf.cell(0, 6, f"Amount in Words: Indian Rupee {words} Only", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")

        self.pdf.ln(8)

    def footer(self):
        footer_y = self.pdf.get_y()
        self.pdf.set_fill_color(*self.light_gray)
        self.pdf.rect(0, footer_y, 210, 20, 'F')

        self.pdf.set_draw_color(*self.accent_orange)
        self.pdf.set_line_width(1)
        self.pdf.line(10, footer_y + 5, 200, footer_y + 5)

        self.pdf.set_y(footer_y + 8)
        self.pdf.set_font("helvetica", 'I', 8)
        self.pdf.set_text_color(*self.medium_gray)
        self.pdf.cell(0, 4, "-- This is a system-generated document --", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")

        self.pdf.set_font("helvetica", 'B', 10)
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
        filename = f"{self.output_folder}/{emp_name}_{self.month}.pdf"
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
