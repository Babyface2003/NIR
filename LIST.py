from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

import time
import glob
import os
import shutil
def parse_1(driver):
    download_path = r"D:\NIR\LIST"
    op = Options()
    op.add_argument('--disable-notifications')
    op.add_experimental_option("prefs", {
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    driver = webdriver.Chrome(options=op)

    driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': download_path}}
    command_result = driver.execute("send_command", params)

    driver.get('https://edu.stankin.ru/?')

    entrance = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/header/div[1]/div/div/div[2]/div[1]/div'))
    )
    entrance.click()

    time.sleep(1)

    login = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'username'))
    )
    login.send_keys('st621233')

    password = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'password'))
    )
    password.send_keys('Boom1979')
    password.send_keys(Keys.RETURN)
    time.sleep(1)
    guest_access = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                                                  '/section/aside[1]/section/div/div/div[1]/div[2]/div[2]/div[1]/a[3]/div[2]'))
    )
    guest_access.click()
    time.sleep(1)
    timetable_of_classes = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div'
                                                  '/section/div[2]/div/div[3]/div[2]/div/div[1]/a[1]'))
    )
    timetable_of_classes.click()

    course = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                                                  '/section/div[1]/nav/ul/li[1]/a'))
    )
    course.click()
    time.sleep(1)
    #Бакалавры первый курс скачивание
    list_1 = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]/section/div[2]/div/div/ul'
                                                 '/li[2]/div[2]/ul/li[1]/div/div[1]/div/div[1]/div/div[2]/div/a'))
    )
    list_1.click()
    time.sleep(2)
    download_first_course = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                                                   '/section/div[2]/div[2]/div/div/div/div/div/div/table'
                                                  '/tbody/tr/td[3]/span/a/span[2]'))
    )
    download_first_course.click()

    time.sleep(2)

    xls_files = glob.glob(os.path.join(download_path, "*.xls"))
    latest_xls_file = max(xls_files, key=os.path.getctime)

    new_filename = os.path.join(download_path, "first_course.xls")
    shutil.move(latest_xls_file, new_filename)

    driver.back()


    # Бакалавры второй курс скачивание
    list_2 = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]/section/div[2]'
                       '/div/div/ul/li[2]/div[2]/ul/li[2]/div/div[1]/div/div[1]/div/div[2]/div/a'))
    )
    list_2.click()

    download_second_course = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                                                  '/section/div[2]/div[2]/div/div/div/div/div/div/table/'
                                                  'tbody/tr/td[3]/span/a/span[2]'))
    )
    download_second_course.click()

    time.sleep(1)

    xls_files = glob.glob(os.path.join(download_path, "*.xls"))
    latest_xls_file = max(xls_files, key=os.path.getctime)

    new_filename = os.path.join(download_path, "second_course.xls")
    shutil.move(latest_xls_file, new_filename)

    driver.back()

    # Бакалавры третий курс скачивание
    list_3 = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                       '/section/div[2]/div/div/ul/li[2]/div[2]/ul/li[3]/div/div[1]/div/div[1]/div/div[2]/div/a'))
    )
    list_3.click()

    download_third_course = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                                                  '/section/div[2]/div[2]/div/div/div/div/div/div/table'
                                                  '/tbody/tr/td[3]/span/a/span[2]'))
    )
    download_third_course.click()

    time.sleep(1)

    xls_files = glob.glob(os.path.join(download_path, "*.xls"))
    latest_xls_file = max(xls_files, key=os.path.getctime)

    new_filename = os.path.join(download_path, "third_course.xls")
    shutil.move(latest_xls_file, new_filename)

    driver.back()
    # Бакалавры четвёртый курс скачивание
    list_4 = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                       '/section/div[2]/div/div/ul/li[2]/div[2]/ul/li[4]/div/div[1]/div/div[1]/div/div[2]/div/a'))
    )
    list_4.click()

    download_fourth_course = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                                                  '/section/div[2]/div[2]/div/div/div/div/div/div/table'
                                                  '/tbody/tr/td[3]/span/a/span[2]'))
    )
    download_fourth_course.click()

    time.sleep(1)

    xls_files = glob.glob(os.path.join(download_path, "*.xls"))
    latest_xls_file = max(xls_files, key=os.path.getctime)

    new_filename = os.path.join(download_path, "fourth_course.xls")
    shutil.move(latest_xls_file, new_filename)

    driver.back()
    # Бакалавры пятый курс скачивание
    list_5 = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                       '/section/div[2]/div/div/ul/li[2]/div[2]/ul/li[5]/div/div[1]/div/div[1]/div/div[2]/div/a'))
    )
    list_5.click()

    download_fifth_course = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                                                  '/section/div[2]/div[2]/div/div/div/div/div/div/table'
                                                  '/tbody/tr/td[3]/span/a/span[2]'))
    )
    download_fifth_course.click()

    time.sleep(1)

    xls_files = glob.glob(os.path.join(download_path, "*.xls"))
    latest_xls_file = max(xls_files, key=os.path.getctime)

    new_filename = os.path.join(download_path, "fifth_course.xls")
    shutil.move(latest_xls_file, new_filename)

    driver.back()

    # Магистратура первый курс скачивание
    list_mag_1 = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                       '/section/div[2]/div/div/ul/li[2]/div[2]/ul/li[6]/div/div[1]/div/div[1]/div/div[2]/div/a'))
    )
    list_mag_1.click()

    download_mag_first_course = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                                                  '/section/div[2]/div[2]/div/div/div/div/div/div/table'
                                                  '/tbody/tr/td[3]/span/a/span[2]'))
    )
    download_mag_first_course.click()

    time.sleep(1)

    xls_files = glob.glob(os.path.join(download_path, "*.xls"))
    latest_xls_file = max(xls_files, key=os.path.getctime)

    new_filename = os.path.join(download_path, "mag_first_course.xls")
    shutil.move(latest_xls_file, new_filename)

    driver.back()

    # Магистратура первый курс скачивание
    list_mag_2 = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                       '/section/div[2]/div/div/ul/li[2]/div[2]/ul/li[7]/div/div[1]/div/div[1]/div/div[2]/div/a'))
    )
    list_mag_2.click()

    download_mag_second_course = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/div/div[1]/div/div/div[1]'
                                                  '/section/div[2]/div[2]/div/div/div/div/div/div/table'
                                                  '/tbody/tr/td[3]/span/a/span[2]'))
    )
    download_mag_second_course.click()

    time.sleep(1)

    xls_files = glob.glob(os.path.join(download_path, "*.xls"))
    latest_xls_file = max(xls_files, key=os.path.getctime)

    new_filename = os.path.join(download_path, "mag_second_course.xls")
    shutil.move(latest_xls_file, new_filename)