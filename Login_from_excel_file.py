from selenium import webdriver
import unittest,time
import selenium.webdriver.common.keys
from selenium.webdriver.common.by import By
import os
import xlrd
import xlwt
import HtmlTestRunner
from xlutils.copy import copy
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException

direct = os.getcwd()

class FactorLab(unittest.TestCase):
    def setUp(self):
        dir = os.path.dirname(__file__)
        chrome_driver_path = dir + "\chromedriver.exe"

        self.driver = webdriver.Chrome(chrome_driver_path)
        self.driver.implicitly_wait(30)
        self.driver.maximize_window()
        self.driver.get("https://staging1.factorlablambdaapis.com/login")

    def test_Login(self):
        driver = self.driver

        excel_file = 'Sample_excel.xls'
        workbook = xlrd.open_workbook(excel_file)
        wb = copy(workbook)
        worksheet = workbook.sheet_by_index(0)
        s = wb.get_sheet(0)
        total_rows = worksheet.nrows
        total_cols = worksheet.ncols
        for j in range(1, total_rows):
            for i in range(1, total_rows):
                rownum = worksheet.cell_value(i, 0)
                colnum = worksheet.cell_value(i, 1)

                email_field = driver.find_element("id", 'login-username')
                email_field.send_keys(rownum)
                continue_button = driver.find_element("xpath", '/html/body/app-root/app-login/div/div[2]/div/div[2]/form/div[2]/button')
                continue_button.click()

                try:
                    signin_button = driver.find_element("xpath", '/html/body/app-root/app-login/div/div[2]/div/div[2]/form/div[3]/button[1]')
                    time.sleep(2)
                    if signin_button.is_displayed():
                        print('User ' + rownum + ' is valid')
                        #s.write(i, 1, 'User ' + rownum + ' is valid')
                        time.sleep(3)
                        i = i + 1
                    else:
                        j = j + 1
                except NoSuchElementException:
                    print('User ' + rownum + ' is not valid')
                    #s.write(i, 1, 'User ' + rownum + ' is not valid')

                password_field = driver.find_element("id", 'login-password')
                password_field.send_keys(colnum)
                signin_button = driver.find_element("xpath", '/html/body/app-root/app-login/div/div[2]/div/div[2]/form/div[3]/button[1]')
                signin_button.click()
                print('User ' + rownum + ' logged in successfully')
                time.sleep(5)

                logout_button = driver.find_element("xpath", '//*[@id="navbarDropdown"]/img')
                logout_button.click()
                sign_out = driver.find_element("xpath", '//*[@id="navbarHeaderContent"]/ul/li[7]/div/a[2]')
                sign_out.click()
                print('User ' + rownum + ' signed out successfully')


            #wb.save('New_excel.xls')
            self.driver.close()
            break


    def teardown(self):
        self.driver.close()

if __name__ == "__main__":
    unittest.main(testRunner=HtmlTestRunner.HTMLTestRunner(output='C:\\Users\\Rakeshkumar\\PycharmProjects\\MyPythonProject\\Reports'))