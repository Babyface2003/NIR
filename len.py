from openpyxl import load_workbook
file_name = 'D:\\NIR\\LIST_NEW\\first_course.xlsx'
wb = load_workbook(file_name)
sheet = wb['Поток Э2']
count = 0
for cell in sheet['D']:
    if cell.value is not None:
        count += 1
print(count)