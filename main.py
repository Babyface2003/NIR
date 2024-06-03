import pandas as pd
from selenium import webdriver
from LIST import parse_1
import time
import os
import win32com.client as client

import xlwings as xw

def excel_to_csv(df):
    csv_data = df.to_csv(index=False)
    return csv_data
if __name__ == "__main__":
    driver = webdriver.Chrome()

    districts = ["list_groups"]
    for district in districts:
        if district == 'list_groups':
            parsed_data = parse_1(driver)
    driver.quit()
    time.sleep(1)

    ##

    excel = client.Dispatch("excel.application")

    for file in os.listdir("D:\\NIR\\LIST\\"):
        filename, file_extension = os.path.splitext(file)
        input_path = os.path.join("D:\\NIR\\LIST\\", file)
        if filename == "first_course" and file_extension == ".xlsx":
            output_folder = "D:\\NIR\\LIST_NEW\\first_courses\\"
        elif filename == "second_course" and file_extension == ".xlsx":
            output_folder = "D:\\NIR\\LIST_NEW\\second_courses\\"
        elif filename == "third_course" and file_extension == ".xlsx":
            output_folder = "D:\\NIR\\LIST_NEW\\third_courses\\"
        elif filename == "fourth_course" and file_extension == ".xlsx":
            output_folder = "D:\\NIR\\LIST_NEW\\fourth_courses\\"
        elif filename == "fifth_course" and file_extension == ".xlsx":
            output_folder = "D:\\NIR\\LIST_NEW\\fifth_courses\\"
        elif filename == "mag_first_course" and file_extension == ".xlsx":
            output_folder = "D:\\NIR\\LIST_NEW\\mag_first_courses\\"
        elif filename == "mag_second_course" and file_extension == ".xlsx":
            output_folder = "D:\\NIR\\LIST_NEW\\mag_second_courses\\"
        wb = excel.Workbooks.Open(input_path)
        output_path = os.path.join(output_folder, file)
        wb.SaveAs(output_path, 52)  # 52 corresponds to xlOpenXMLWorkbook (xlsx) format
        wb.Close()
    excel.Quit()

    """    df = pd.read_excel(file_path)
        xml_data = excel_to_csv(df)

        csv_file_name = os.path.splitext(file)[0] + ".csv"
        with open(csv_file_name, 'w') as f:
            f.write(xml_data)

        os.remove(file)
   """