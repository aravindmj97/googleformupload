# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import time
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import openpyxl


def main():
    current = " "
    wb = openpyxl.load_workbook('abc.xlsx')
    sheet = wb['BACK UP']
    
    driver = webdriver.Chrome()
    driver.get("https://docs.google.com/forms/d/e/1FAIpQLSeTThy67d3USQEbzdxhoknyNbBtEr2wU_W0Qo-0pZdgSePeLA/viewform")

    for j in range(443,465):
        if sheet['E'+str(j)].value:
            driver.find_element_by_class_name("quantumWizMenuPaperselectOptionList").click()
            options=driver.find_element_by_class_name("exportSelectPopup")
            time.sleep(3)
            print(options)
            contents = options.find_elements_by_tag_name('content')
            
            [i.click() for i in contents if i.text == "PARAVOR"]
            time.sleep(2)

            village = driver.find_element_by_name("entry.1279182967")
            village.clear()
            if sheet['F'+str(j)].value:
                village.send_keys(sheet['F'+str(j)].value)
            else:
                village.send_keys("Paravor")
            camp = driver.find_element_by_name("entry.775513399")
            camp.clear()
            camp.send_keys('CUSAT')


            name = driver.find_element_by_name("entry.282201977")
            name.clear()
            name.send_keys(sheet['B'+str(j)].value)

            age = driver.find_element_by_name("entry.433850487")
            
            age.clear()
            age.send_keys(str(int(sheet['C'+str(j)].value)))

            contents = driver.find_elements_by_class_name('docssharedWizToggleLabeledContainer')
            if "F" in str(sheet['D'+str(j)].value):
                contents[0].click()
            else:
                contents[1].click()
                
            


            addr = driver.find_element_by_name("entry.224155327")
            addr.clear()
            addr.send_keys(sheet['E'+str(j)].value)

            contact = driver.find_element_by_name("entry.1399226084")
            contact.clear()
            if sheet['H'+str(j)].value:

                if type(sheet['H'+str(j)].value) == str:
                    contact.send_keys(str(sheet['H'+str(j)].value)[:10])
                else:
                    contact.send_keys(str(int(sheet['H'+str(j)].value))[:10])
                    current = str(int(sheet['H'+str(j)].value))
            else:
                contact.send_keys(current)
            #adhar = driver.find_element_by_name("entry.1563307912")
            #adhar.clear()
            #adhar.send_keys('1022')

            submit = driver.find_elements_by_class_name('quantumWizButtonPaperbuttonEl')
            time.sleep(2)
            submit[0].click()
            driver.back()

    print("completed")


main()

input()
