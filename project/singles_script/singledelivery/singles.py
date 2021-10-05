#import the datetime in order to get the current time
from datetime import datetime

#import the csv in order to write the data to csv
import csv

#Import the Selenium 
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


from selenium.webdriver.chrome.options import Options
import glob

import time
import random


#Import the os in order to get the current directory
import os
ROOT_PATH = os.path.realpath(os.path.dirname(os.path.abspath(__file__)))

class SinglesApp():
    def __init__(self, name=""):
        #self.name: str = name
        self.driver = self.make_driver()
        
        self.inputpath = ''
        self.outputpath = ''
        self.errorsymbols = []


    def make_driver(self):
        options = Options()
        if os.name == "nt":
            options.add_argument('--headless')
            options.add_argument("--no-sandbox")
            driver = webdriver.Chrome(options=options)
        else:
            options.add_argument('--headless')
            options.add_argument("--no-sandbox")
            driver = webdriver.Chrome(chrome_options=options, executable_path='./chromedriver')


        driver.maximize_window()
        driver.implicitly_wait(10)

        return driver


    def close(self):
    	#Close the selenium driver
        self.driver.quit()


    def time_sleep(self, type):
        if type == 1:
            sleeptime = random.randrange(10,100)/100
        elif type == 2:
            sleeptime = random.randrange(70, 200)/100
        elif type == 3:
            sleeptime = random.randrange(100, 300)/100
        elif type == 4:
            sleeptime = random.randrange(150, 400)/100
        elif type == 5:
            sleeptime = random.randrange(400, 500)/100
        elif type == 401:
            sleeptime = random.randrange(60, 100)
        time.sleep(sleeptime)

    def getInputHtmlsPath(self):
        lstFiles = []

        import configparser

        parser = configparser.ConfigParser()
        parser.read('setting.ini')

        self.inputpath = parser.get('global', 'input')
        self.outputpath = parser.get('global',  'output')

        for (root, dirs, file) in os.walk(self.inputpath):
            for f in file:
                if '.html' in f:
                    lstFiles.append(os.path.abspath(os.path.join(root, f)))

        return lstFiles

    def processing(self):
        lstInputPaths = self.getInputHtmlsPath()

        result = []
        header_data = ['Debut Date', 'Peak Date', 'Peak Pos', 'Wks at Peak', 'Weeks Charted', 'Chart Title', 'Artist', 'A-Side', 'B-Side', 'Label & Number']

        result.append(header_data)
        isfirst = True
        index = 1
        for path in lstInputPaths:
            if isfirst == True:
                isfirst = False
                continue

            if '\\admin.html' in path or '\\record_research.html' in path or '__MACOSX' in path:
                continue

            #print(index)
            index += 1

            driver = self.driver
            try:
                driver.get(path)
            except Exception as e:
                print("driver error:{}".format(e))
                continue
            print("-------------------------------------------------------")
            print(path)
            rows = []
            try:
                rows = driver.find_elements_by_xpath("//*[@id='search_results']/table/tbody/tr")
            except Exception as e:
                print("getting rows error:{}".format(e))
            
            isrowfirst = True
            for row in rows:
                if isrowfirst == True:
                    isrowfirst = False
                    continue
                debutdate = row.find_element_by_xpath("./td[2]").text
                peakdate = row.find_element_by_xpath("./td[3]").text
                peakpos = row.find_element_by_xpath("./td[4]").text
                wks = row.find_element_by_xpath("./td[5]").text
                weeks = row.find_element_by_xpath("./td[6]").text
                chart = row.find_element_by_xpath("./td[7]").text
                artist = row.find_element_by_xpath("./td[8]").text
                aside = row.find_element_by_xpath("./td[9]").text
                bside = row.find_element_by_xpath("./td[10]").text
                label = row.find_element_by_xpath("./td[11]").text
                
                onedata = [debutdate, peakdate, peakpos, wks, weeks, chart, artist, aside, bside, label]
                print(onedata)
                result.append(onedata)
            
        self.writeDatatoExcel(result)

        #print(self.errorsymbols)
        return

    def writeDatatoExcel(self, result):
        outputfilename = self.outputpath + '/' + datetime.now().strftime("%Y%m%d%H%M%S") + '.xlsx'
        import xlsxwriter

        workbook = xlsxwriter.Workbook(outputfilename)
        worksheet = workbook.add_worksheet()

        for row_num, row_data in enumerate(result):
            for col_num, col_data in enumerate(row_data):
                worksheet.write(row_num, col_num, col_data)

        workbook.close()


if __name__ == "__main__":
    app = SinglesApp()
    app.processing()
    app.close()