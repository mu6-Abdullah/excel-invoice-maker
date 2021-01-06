# This is a sample Python script.
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import load_workbook
from os import path

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
def xPathReturner(rowNum, id):
    webPageAccessXPath = '/html/body/div[3]/div/div[3]/div/div/div[1]/div[2]/div[2]/div[4]/table/tbody/'
    if id == 'date':
        return webPageAccessXPath + 'tr['+ str(rowNum) + ']/td[2]/span[2]/textarea'
    elif id == 'rate':
        return webPageAccessXPath + 'tr['+ str(rowNum) + ']/td[3]/span/input'
    elif id == 'description':
        return webPageAccessXPath + 'tr['+ str(rowNum) + ']/td[2]/span[1]/div/div/input'
    elif id == 'quantity':
        return webPageAccessXPath + 'tr['+ str(rowNum) + ']/td[4]/input'

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    try:
        wb = load_workbook('C:/Users/m27ab/OneDrive/Desktop/python/decemberInvoice.xlsx')
    except FileNotFoundError:
        print('incorrect name of file')
    ws = wb.active



    web = webdriver.Chrome('C:/Users/m27ab/OneDrive/Desktop/python/chromedriver.exe')
    web.get('https://www.invoicesimple.com/invoice-generator?utm_source=website&utm_medium=organic&utm_content=main_nav')
    time.sleep(2)
    name = ['Jin Lee', 'Shawn']
    email = ['mu6@ualberta.ca', 'shawn.du@gongchacanada.ca']
    address = ['1117 23B Ave NW', '1356 Windermere Way SW']
    phone = ['5879885823','5879876988']
    postalCode = ['T6J 4P3', 'T6W 2J3']

    pageFromName = web.find_element_by_xpath('//*[@id="invoice-company-name"]')
    pageFromName.send_keys(name[0])

    pageToName = web.find_element_by_xpath('//*[@id="invoice-client-name"]')
    pageToName.send_keys(name[1])

    pageFromEmail = web.find_element_by_xpath('//*[@id="invoice-company-email"]')
    pageFromEmail.send_keys(email[0])

    pageToEmail = web.find_element_by_xpath('//*[@id="invoice-client-email"]')
    pageToEmail.send_keys(email[1])

    pageFromAddress = web.find_element_by_xpath('//*[@id="invoice-company-address1"]')
    pageFromAddress.send_keys(address[0])

    pageToAddress = web.find_element_by_xpath('//*[@id="invoice-client-address1"]')
    pageToAddress.send_keys(email[1])

    pageFromPhone = web.find_element_by_xpath('//*[@id="invoice-company-phone"]')
    pageFromPhone.send_keys(phone[0])

    pageToPhone = web.find_element_by_xpath('//*[@id="invoice-client-phone"]')
    pageToPhone.send_keys(phone[1])

    pageFromCity = web.find_element_by_xpath('//*[@id="invoice-company-address2"]')
    pageFromCity.send_keys('Edmonton, AB')

    pageToCity = web.find_element_by_xpath('//*[@id="invoice-client-address2"]')
    pageToCity.send_keys('Edmonton, AB')

    pageFromPostalCode = web.find_element_by_xpath('//*[@id="invoice-company-address3"]')
    pageFromPostalCode.send_keys(postalCode[0])

    pageToPostalCode = web.find_element_by_xpath('//*[@id="invoice-client-address3"]')
    pageToPostalCode.send_keys(postalCode[1])

    sectionAddButton = web.find_element_by_xpath('//*[@id="invoice-item-add"]')

    pageTaxSection = web.find_element_by_xpath('/html/body/div[3]/div/div[3]/div/div/div[2]/div/div/div[2]/table[1]/tbody/tr[3]/td[2]/span/span/input')
    for i in range(7):
        pageTaxSection.send_keys(Keys.BACK_SPACE)
    pageTaxSection.send_keys('0')

    rowCount = 0
    firstRow = True
    for row in ws.values:
        if firstRow == True:
            firstRow = False
            continue
        else:
            rowCount += 1
            print(row)
            date = str(row[1])
            hours = str(row[2])
            description = str(row[3])
            pageDescription = web.find_element_by_xpath(xPathReturner(rowCount, 'description'))
            pageDescription.send_keys(description)

            pageDate = web.find_element_by_xpath(xPathReturner(rowCount, 'date'))
            pageDate.send_keys(date)

            pageRate = web.find_element_by_xpath(xPathReturner(rowCount, 'rate'))
            pageRate.send_keys('30')

            pageAmount = web.find_element_by_xpath(xPathReturner(rowCount, 'quantity'))
            print(hours)
            pageAmount.send_keys(Keys.BACK_SPACE)
            pageAmount.send_keys(hours)

            sectionAddButton.click()




