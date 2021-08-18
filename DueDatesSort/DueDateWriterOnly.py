from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

class CreateWorkbook:
    def __init__(self, workbook, sheet):
        self.workbook = workbook
        self.sheet = sheet

    def add_info_sheet(self, sheet):
        rows = 0
        cols = 0
        due_date = []
        finished_adding = False
        while not finished_adding:
            add_hw = input("What homework do you wish to add?: ")
            if 'quit' in add_hw or 'q' in add_hw:
                finished_adding = True
                break
            add_date = input("What date is this homework due?: ").upper()
            add_time = input("What time is this homework due?: ").upper()
            add_course = input("What course is this hw for?: ").upper()
            add_notes = input("Any additional notes?: ")
            total_info = (add_hw, add_date, add_time, add_course, add_notes)
            due_date.append(total_info)

        for data in due_date:
            sheet.append(data)

    def sheet_formatting(self, sheet):
        sheet.column_dimensions['A'].width = 20
        sheet.column_dimensions['B'].width = 18
        sheet.column_dimensions['D'].width = 16
        sheet.column_dimensions['E'].width = 40

        sheet['A1'] = 'Homework'
        sheet['A1'].font = Font(bold=True)
        sheet['B1'] = 'Date'
        sheet['B1'].font = Font(bold=True)
        sheet['C1'] = 'Time'
        sheet['C1'].font = Font(bold=True)
        sheet['D1'] = 'Course'
        sheet['D1'].font = Font(bold=True)
        sheet['E1'] = 'Notes'
        sheet['E1'].font = Font(bold=True)
