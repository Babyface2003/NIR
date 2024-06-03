import xlwings as xw

wb = xw.Book('D:\\NIR\\LIST_NEW\\first_course.xlsx')

macro1 = wb.macro('Module18.Dlya_2_3_groups')
macro1()
wb.save()
wb.close()