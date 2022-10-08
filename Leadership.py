import openpyxl

wb = openpyxl.load_workbook('Cadet_Full_Track_Report-8_23_2022.xlsx')
ws = wb.active
def leadTest_check():
    LeadTestStatus = []
    row_num = 1
    for row in ws.iter_rows(min_row=2, min_col=13, max_col=13, max_row=42, values_only=True):
        row_num +=1
        if row[0] > 79:
            LeadTestStatus.append(ws[f'A{row_num}'].value)
    return LeadTestStatus

