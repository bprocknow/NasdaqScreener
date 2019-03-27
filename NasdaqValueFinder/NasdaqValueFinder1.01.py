import os
import sys
#For reading websites
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
#For reading Google Sheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
#For waiting for pages to load if they are causing issues
import time
import re
import signal
import xlrd
import openpyxl
import heapq

'''
Overview:
This is a stock screener
Takes input from an excel that was downloaded from www.nasdaq.com and it contains all of the stocks on the nasdaq.
The input is the ticker and the market cap of the stock.  The balance sheet is found from the ticker, and all
of the values that are used as part of the screener are calculated as a function of the market cap and balance sheet.
There are nodes that store the tickers, market cap, balance sheet, and all of the computed values.
The nodes are written to an excel after the balance sheet and computed values are found.
After the new excel is created with the necessary information, it can be maximized and printed out, or used in
any other way (AI??? ;))
'''

'''
Problems:
'''

'''
Node object holds the ticker, marketCap, balanceSheetList.
BVMC4Y = Book value to market cap last four years
BVMC1Y = Book value to market cap last year
CAMC4Y = Current Assets to Market Cap last four years
CAMC1Y = Current Assets to Market Cap last year
CEMC4Y = Cash and Cash Equivalent to Market Cap last four years
CACL4Y = Current Assets to Current Liabilities last four years
CACL1Y = Current Assets to Current Liabilities last year
'''
class Node:
    def __init__(self, browser, ticker, marketCap):
        self.ticker = ticker
        self.marketCap = findMarketCap(marketCap)
        self.balanceSheet = getBalanceSheet(browser, ticker)
        self.BVMC4Y = computeBookValueToMarketCapLastFourYears(self.balanceSheet, self.marketCap)
        self.BVMC1Y = computeBookValueToMarketCapLastYear(self.balanceSheet, self.marketCap)
        self.CAMC4Y = computeCurrentAssetsToMarketCapLastFourYears(self.balanceSheet, self.marketCap)
        self.CAMC1Y = computeCurrentAssetsToMarketCapLastYear(self.balanceSheet, self.marketCap)
        self.CEMC4Y = computeCashAssetsToMarketCapLastFourYears(self.balanceSheet, self.marketCap)
        self.CACL4Y = currentAssetsToCurrentLiabilitiesFourYears(self.balanceSheet)
        self.CACL1Y = currentAssetsToCurrentLiabilitiesLastYear(self.balanceSheet)
    def getBalanceSheet(self):
        return self.balanceSheet
    def getMarketCap(self):
        return self.marketCap
    def getTicker(self):
        return self.ticker
    def getBVMC4Y(self):
        return self.BVMC4Y
    def getBVMC1Y(self):
        return self.BVMC1Y
    def getCAMC4Y(self):
        return self.CAMC4Y
    def getCAMC1Y(self):
        return self.CAMC1Y
    def getCEMC4Y(self):
        return self.CEMC4Y
    def getCACL4Y(self):
        return self.CACL4Y
    def getCACL1Y(self):
        return self.CACL1Y

#Returns the market cap in an int from the inputted string
def findMarketCap(marketCap):
    try:
        #Returns the market caps that are below one million dollars
        return int(marketCap)
    except ValueError:
        pass
    if marketCap[0] is "n":
        return None
    marketCapStr=''
    multiplierFactor=0
    indexOfPeriod=0
    #Get list of the digits in the marketCap & set multiplierFactor
    i=0
    while(i<len(marketCap)):
        if (marketCap[i]=='M'):
            multiplierFactor=1000000
        if (marketCap[i]=='B'):
            multiplierFactor=1000000000
        if (marketCap[i]=='T'):
            multiplierFactor=1000000000000
        if (marketCap[i]=='.'):
            indexOfPeriod=i-1
        #Remove everything but numbers
        try:
            int(marketCap[i])
            marketCapStr+=marketCap[i]
        except ValueError:
            pass
        i+=1
    try:
        marketCap=int(marketCapStr)
    except ValueError:
        print("fuckedMarketCapException")
        return None
    return marketCap*multiplierFactor/(10**(len(marketCapStr)-indexOfPeriod))

#Opens the nasdaq website and clicks on the cookie box so that the rest of the program isn't fucked.
def clickOnCookieBox(browser):
    browser.get("http://www.nasdaq.com")
    cookieClickLocation = ("//div[@id='cookieConsent']/a[@id='cookieConsentOK']")
    timeout = 15
    openedPage = True
    while(openedPage):
        try:
            WebDriverWait(browser, timeout).until(EC.visibility_of_element_located((By.XPATH, cookieClickLocation)))
            openedPage=False
        except TimeoutException:
            print("cookieClickerTimeoutException")
            pass
    #Click on accept cookies
    cookieClickLink = browser.find_element_by_xpath(cookieClickLocation)
    cookieClickLink.click()

'''
Returns the balance sheet of the ticker from nasdaq.com
Wait for the balance sheet numbers to be visible.
If the ticker is not available or doesn't have any information on the website,
return None
'''
def getBalanceSheet(browser, ticker):
    URL = "https://www.nasdaq.com/symbol/"+ ticker+ '/financials?query=balance-sheet/'
    browser.get(URL)
    #Wait until the top of the Annual Income Statement is loaded.  The rest of the page is hopefully loaded by then. :)
    timeout = 5
    try:
        WebDriverWait(browser, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="financials-iframe-wrap"]/div[1]/h3')))
    except TimeoutException:
        print ("Timeout error 0")
        #There is an instance of nasdaq's website having nothing on the screen except a warning sign.
        try:
            nasdaqRuntimeError = browser.find_element_by_xpath('/html/body/span/h2/i')
            if (nasdaqRuntimeError.text == "Runtime Error"):
                print("nasdaqRuntimeException")
                return None
        except Exception:
            pass
    #Test if there is any data for this symbol
    try:
        #Location of element that checks if there is any information on the nasdaq website about the ticker.
        noDataAvailableElement = browser.find_element_by_xpath('//*[@id="quotes_content_left_nodatatext"]/span/b')
        if (noDataAvailableElement.text == "There is currently no data for this symbol."):
            return None
    except Exception:
        pass
    try:
        #Location of element that checks if there is any information on the nasdaq website about the ticker.
        #Sometimes there is the element above, but sometimes it is this element.  IDK...
        noDataAvailableElement = browser.find_element_by_xpath('//*[@id="left-column-div"]/div[1]')
        if (len(noDataAvailableElement.text)>3):
            return None
        #Element was found, that means there is no data for the ticker.
    except Exception:
        #Element was not found
        pass
    try:
        #Tests the current URL, if the page doesn't have the Cash and Cash Equivalents
        #Text, it is not on the right page.  Search for the link to the balance sheet and click
        testingCorrectElement = browser.find_element_by_xpath("//div[@class='genTable']/table/tbody/tr[2]/th")
        if (testingCorrectElement.text != "Cash and Cash Equivalents"):
            balanceSheetLink = browser.find_element_by_xpath('//*[@id="tab2"]/span')
            balanceSheetLink.click()
    #Sometimes the noDataAvailableElement (^^^) doesn't work, and it passes to this point.  Try again
    except NoSuchElementException:
        time.sleep(2)
        return getBalanceSheet(browser, ticker)
    #Sometimes the page loads to quickly and the element hasn't loaded yet.  Wait 1 sec and try again so it loaded
    except Exception:
        time.sleep(2)
        testingCorrectElement = browser.find_element_by_xpath("//div[@class='genTable']/table/tbody/tr[2]/th")
        if (testingCorrectElement.text != "Cash and Cash Equivalents"):
            balanceSheetLink = browser.find_element_by_xpath('//*[@id="tab2"]/span')
            balanceSheetLink.click()
    #Wait for the balance sheet to load.
    try:
        timeout = 5
        #Location of last element on the balance sheet
        WebDriverWait(browser, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="financials-iframe-wrap"]/div[1]/table/tbody/tr[34]/td[2]')))
    except TimeoutException:
        print ("Timeout error 1")
        time.sleep(5)
        return getBalanceSheet(browser, ticker)
    #Contains the int of all asset prices
    BalanceSheet=[]
    row=2
    while(row<35):
        if (row == 8):
            pass
        elif (row == 16):
            pass
        elif (row == 27):
            pass
        else:
            index = 2
            #tmpList adds all the assets in each row, then is added to the Balance sheet list.
            tmpList=[]
            while(index<6):
                #HTMLObject is the type pulled from the HTML website.  Needs to be turned into an int
                try:
                    location = "//div[@class='genTable']/table/tbody/tr["+str(row)+"]/td["+str(index)+"]"
                    HTMLObject=browser.find_elements_by_xpath(location)
                    #Int of asset price
                    AssetInt = [x.text for x in HTMLObject]
                    tmpList.append(AssetInt[0])
                    index = index+1
                #Try this thing again, the website was loaded too quickly and elements were missed
                except IndexError:
                    print("indexError")
                    return getBalanceSheet(browser, ticker)
                #Sometime the websites get mad for sending too many requests too fast, take a break try again
                except Exception:
                    print("elementNotFoundException")
                    time.sleep(3)
                    location = "//div[@class='genTable']/table/tbody/tr["+str(row)+"]/td["+str(index)+"]"
                    HTMLObject=browser.find_elements_by_xpath(location)
                    #Int of asset price
                    AssetInt = [x.text for x in HTMLObject]
                    tmpList.append(AssetInt[0])
                    index = index+1
            BalanceSheet.append(tmpList)
        row = row+1

    return BalanceSheet

#Take list and find the numbers in each index.  Put the numbers in the index in a string
#and sum the string with the numbers of the other strings.  Return the total value
def getIntFromList(numYears, list):
    sumOfList=0
    i=0
    while(i<numYears):
        tmpStr=""
        for x in list[i]:
            try:
                int(x)
                tmpStr+=x
            except ValueError:
                pass
        i+=1
        try:
            sumOfList+=int(tmpStr)*1000
        except ValueError:
            pass
    return sumOfList

'''
Uses getIntFromList ^^^ Returns the int value from the balance sheet
Used for maxHeap
Finds the value of the balance sheet over the specified time frame
If the balance sheet or market cap are not available, return -1
'''
def computeBookValueToMarketCapLastFourYears(balanceSheet, marketCap):
    if balanceSheet is None:
        return -1
    else:#Gets the total Assets from the list inside of the list.
        totalAssetsList = balanceSheet[12]
        totalAssets = getIntFromList(4,totalAssetsList)/len(totalAssetsList)
        goodwillList =  balanceSheet[8]
        goodwill =  getIntFromList(4,goodwillList)/len(goodwillList)
        intangibleList = balanceSheet[9]
        intangible = getIntFromList(4,intangibleList)/len(intangibleList)
        liabilityList = balanceSheet[22]
        liability =  getIntFromList(4,liabilityList)/len(liabilityList)
        bookValueOverMeanLast4Years = totalAssets-goodwill-intangible-liability
        if (marketCap==None):
            return -1
        BVMK = bookValueOverMeanLast4Years/marketCap
        return BVMK
def computeBookValueToMarketCapLastYear(balanceSheet, marketCap):
    if balanceSheet is None:
        return -1
    totalAssetsList =  balanceSheet[12]
    totalAssets =  getIntFromList(1, totalAssetsList)
    goodwillList =  balanceSheet[8]
    goodwill =  getIntFromList(1, goodwillList)
    intangibleList =  balanceSheet[9]
    intangible =  getIntFromList(1, intangibleList)
    liabilityList =  balanceSheet[22]
    liability =  getIntFromList(1, liabilityList)
    bookValueLastYear = totalAssets-goodwill-intangible-liability
    if (marketCap==None):
        return -1
    BVMK = bookValueLastYear/marketCap
    return BVMK
    #Takes into account 0.5*Inventory
def computeCurrentAssetsToMarketCapLastFourYears(balanceSheet, marketCap):
    if balanceSheet is None:
        return -1
    cashAndCashEquivList = balanceSheet[0]
    cashAndCashEquiv =  getIntFromList(4, cashAndCashEquivList)/len(cashAndCashEquivList)
    shortTermInvestList = balanceSheet[1]
    shortTermInvest = getIntFromList(4,shortTermInvestList)/len(shortTermInvestList)
    netRecievablesList = balanceSheet[2]
    netRevievables = getIntFromList(4,netRecievablesList)/len(netRecievablesList)
    inventoryList =  balanceSheet[3]
    inventory =  getIntFromList(4, inventoryList)/len(inventoryList)
    liabilityList =  balanceSheet[22]
    liability =  getIntFromList(4, liabilityList)/len(liabilityList)
    netCurrentAssets = cashAndCashEquiv+shortTermInvest+netRevievables+0.5*inventory
    currentAssetsLastFourYears = netCurrentAssets-liability
    if (marketCap==None):
        return -1
    CAMK = currentAssetsLastFourYears/marketCap
    return CAMK
def computeCurrentAssetsToMarketCapLastYear(balanceSheet, marketCap):
    if balanceSheet is None:
        return -1
    cashAndCashEquivList = balanceSheet[0]
    cashAndCashEquiv =  getIntFromList(1, cashAndCashEquivList)
    shortTermInvestList = balanceSheet[1]
    shortTermInvest = getIntFromList(1,shortTermInvestList)
    netRecievablesList = balanceSheet[2]
    netRevievables = getIntFromList(1,netRecievablesList)
    inventoryList =  balanceSheet[3]
    inventory =  getIntFromList(1, inventoryList)
    liabilityList =  balanceSheet[22]
    liability =  getIntFromList(1, liabilityList)
    netCurrentAssets = cashAndCashEquiv+shortTermInvest+netRevievables+0.5*inventory
    currentAssetsLastYear = netCurrentAssets-liability
    if (marketCap==None):
        return -1
    CAMK = currentAssetsLastYear/marketCap
    return CAMK
def computeCashAssetsToMarketCapLastFourYears(balanceSheet, marketCap):
    if balanceSheet is None:
        return -1
    cashAndCashEquivList =  balanceSheet[0]
    cashAndCashEquiv = getIntFromList(4,cashAndCashEquivList)/len(cashAndCashEquivList)
    shortTermList =  balanceSheet[1]
    shortTerm =  getIntFromList(1, shortTermList)/len(shortTermList)
    accountPayableList = balanceSheet[13]
    accountPayable = getIntFromList(4,accountPayableList)/len(accountPayableList)
    shortTermDebtList = balanceSheet[14]
    shortTermDebt = getIntFromList(4,shortTermDebtList)/len(shortTermDebtList)
    otherCurrentLiabilitiesList = balanceSheet[15]
    otherCurrentLiabilities = getIntFromList(4,otherCurrentLiabilitiesList)/len(otherCurrentLiabilitiesList)
    currentLiabilities = accountPayable+shortTermDebt+otherCurrentLiabilities
    cashAssetsFourYears = cashAndCashEquiv+shortTerm-currentLiabilities
    if (marketCap==None):
        return -1
    CAMK = cashAssetsFourYears/marketCap
    return CAMK
def currentAssetsToCurrentLiabilitiesFourYears(balanceSheet):
    if balanceSheet is None:
        return -1
    cashAndCashEquivList = balanceSheet[0]
    cashAndCashEquiv =  getIntFromList(4, cashAndCashEquivList)/len(cashAndCashEquivList)
    shortTermInvestList = balanceSheet[1]
    shortTermInvest = getIntFromList(4,shortTermInvestList)/len(shortTermInvestList)
    netRecievablesList = balanceSheet[2]
    netRevievables = getIntFromList(4,netRecievablesList)/len(netRecievablesList)
    inventoryList =  balanceSheet[3]
    inventory =  getIntFromList(4, inventoryList)/len(inventoryList)
    accountPayableList = balanceSheet[13]
    accountPayable = getIntFromList(4,accountPayableList)/len(accountPayableList)
    shortTermDebtList = balanceSheet[14]
    shortTermDebt = getIntFromList(4,shortTermDebtList)/len(shortTermDebtList)
    otherCurrentLiabilitiesList = balanceSheet[15]
    otherCurrentLiabilities = getIntFromList(4,otherCurrentLiabilitiesList)/len(otherCurrentLiabilitiesList)
    currentLiabilities = accountPayable+shortTermDebt+otherCurrentLiabilities
    netCurrentAssets = cashAndCashEquiv+shortTermInvest+netRevievables+0.5*inventory
    if (currentLiabilities==0):
        currentLiabilities=1
    currentAssetsToLiabilitiesFourYears = netCurrentAssets/currentLiabilities
    return currentAssetsToLiabilitiesFourYears
def currentAssetsToCurrentLiabilitiesLastYear(balanceSheet):
    if balanceSheet is None:
        return -1
    cashAndCashEquivList = balanceSheet[0]
    cashAndCashEquiv =  getIntFromList(1, cashAndCashEquivList)
    shortTermInvestList = balanceSheet[1]
    shortTermInvest = getIntFromList(1,shortTermInvestList)
    netRecievablesList = balanceSheet[2]
    netRevievables = getIntFromList(1,netRecievablesList)
    inventoryList =  balanceSheet[3]
    inventory =  getIntFromList(1, inventoryList)
    accountPayableList = balanceSheet[13]
    accountPayable = getIntFromList(1,accountPayableList)
    shortTermDebtList = balanceSheet[14]
    shortTermDebt = getIntFromList(1,shortTermDebtList)
    otherCurrentLiabilitiesList = balanceSheet[15]
    otherCurrentLiabilities = getIntFromList(1,otherCurrentLiabilitiesList)
    currentLiabilities = accountPayable+shortTermDebt+otherCurrentLiabilities
    netCurrentAssets = cashAndCashEquiv+shortTermInvest+netRevievables+0.5*inventory
    if (currentLiabilities==0):
        currentLiabilities=1
    currentAssetsToLiabilitiesFourYears = netCurrentAssets/currentLiabilities
    return currentAssetsToLiabilitiesFourYears

#Sets up the browser, incognito, doesn't wait for the page to fully load.
def configureBrowser():
    #Initialize path to chromedriver & broswer
    path_to_chromedriver = '/Users/benprocknow/Downloads/chromedriver'
    chrome_options = webdriver.ChromeOptions()
    #Opens an incognito tab
    chrome_options.add_argument('--incognito')
    chrome_options.add_argument('--no-sandbox')
    #With these capablities, when the page is .get(URL), the full page doesn't have
    #To fully load, this is in combination with a wait for element to be visible
    #In the getBalanceSheet method.
    caps = DesiredCapabilities.CHROME
    caps["pageLoadStrategy"] = "none"
    browser = webdriver.Chrome(desired_capabilities=caps, executable_path=path_to_chromedriver,options=chrome_options)
    return browser

def openGoogleSpreadSheet(googleSheetName):
    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name("NasdaqValueFinder.json", scope)
    client = gspread.authorize(creds)
    #Opens the google sheet
    sheet = client.open(googleSheetName).sheet1
    return sheet

def findRowOfGoogle(sheet):
    stringCellLeftOff = str(sheet.find("ENDED HERE"))
    listStringCellLeftOff = stringCellLeftOff.split(" ")
    stringCellLeftOff=listStringCellLeftOff[1]
    stringCellLeftOffRow = ""
    i=0
    while(i<len(stringCellLeftOff)-2):
        if stringCellLeftOff[i] is not "R":
            stringCellLeftOffRow+=stringCellLeftOff[i]
        i+=1
    rowOfGoogle = int(stringCellLeftOffRow)
    return rowOfGoogle

'''
Acts as a main method.  Launches all helper methods to return a list
of Nodes.
Inputs excel sheet with ticker in zero row and market cap in third row
This sheet must be on sheet index 0
Inputs number of rows that should be scanned on the excel.  Input of 'all'
means that all the rows should be checked
Returns a list of the balance sheets of all the tickers
'''
def writeToExcelCalculateValues(excelSpreadsheetName, numberOfRowsCompared, googleSheetName):
    #Operator of the browser
    browser = configureBrowser()
    #Opens the excel that contains the market cap and the ticker
    nasdaqExcel = xlrd.open_workbook(excelSpreadsheetName)
    nasdaqSheet = nasdaqExcel.sheet_by_index(0)
    #Operator of the google sheet that is used to push all of the calculated values
    sheet = openGoogleSpreadSheet(googleSheetName)
    #Finds the row that the program left off on after it was discontinued
    rowOfGoogle = findRowOfGoogle(sheet)
    rowOfExcel = rowOfGoogle-1
    #Open the browser and click on the box that asks if it is ok to use cookies
    clickOnCookieBox(browser)
    #Loop through the rows of the excel sheet and the google Sheet
    while (rowOfExcel<numberOfRowsCompared+1):
        #Used if the google api says that the program is writing to the google sheet too quickly.
        rowOfGoogleBackup = rowOfGoogle
        ticker = str(nasdaqSheet.row_values(rowOfExcel)[0])
        marketCap = nasdaqSheet.row_values(rowOfExcel)[3]
        #Restart the browser if getBalanceSheet threw an exception.  The program usually throws Timeout
        #Exceptions from retrieving the webpage, but sometimes google API kicks me out for writing over
        #500 cells/100 seconds, that's why there is the time.sleep
        createdNode = False
        while(createdNode is False):
            try:
                newNode = Node(browser, ticker, marketCap)
                createdNode = True
            except Exception:
                print("newNodeException")
                browser.quit()
                time.sleep(10)
                browser = configureBrowser()
                clickOnCookieBox(browser)
        try:
            if (newNode.getBalanceSheet() is None or newNode.getMarketCap() is None):
                sheet.update_cell(rowOfGoogle,1," ")
                rowOfExcel+=1
                rowOfGoogle+=1
            else:
                    sheet.update_cell(rowOfGoogle,1,str(ticker))
                    sheet.update_cell(rowOfGoogle,2,str(newNode.getMarketCap()))
                    sheet.update_cell(rowOfGoogle,3, str(newNode.getBVMC4Y()))
                    sheet.update_cell(rowOfGoogle,4,str(newNode.getBVMC1Y()))
                    sheet.update_cell(rowOfGoogle,5,str(newNode.getCAMC4Y()))
                    sheet.update_cell(rowOfGoogle,6,str(newNode.getCAMC1Y()))
                    sheet.update_cell(rowOfGoogle,7,str(newNode.getCEMC4Y()))
                    sheet.update_cell(rowOfGoogle,8,str(newNode.getCACL4Y()))
                    sheet.update_cell(rowOfGoogle,9,str(newNode.getCACL1Y()))
                    rowOfExcel+=1
                    rowOfGoogle+=1
                    #Used to find where the program left off when an exception is thrown
                    sheet.update_cell(rowOfGoogle,1,"ENDED HERE")
        #Once 500 cells are written, google will stop letting the program add more
        #information to the google sheet.  Take a break and start over.
        except Exception:
            print("googleException")
            time.sleep(45)
            if (newNode.getBalanceSheet() is None or newNode.getMarketCap() is None):
                sheet.update_cell(rowOfGoogleBackup,1," ")
                rowOfGoogleBackup+=1
            else:
                sheet.update_cell(rowOfGoogleBackup,1,str(ticker))
                sheet.update_cell(rowOfGoogleBackup,2,str(newNode.getMarketCap()))
                sheet.update_cell(rowOfGoogleBackup,3, str(newNode.getBVMC4Y()))
                sheet.update_cell(rowOfGoogleBackup,4,str(newNode.getBVMC1Y()))
                sheet.update_cell(rowOfGoogleBackup,5,str(newNode.getCAMC4Y()))
                sheet.update_cell(rowOfGoogleBackup,6,str(newNode.getCAMC1Y()))
                sheet.update_cell(rowOfGoogleBackup,7,str(newNode.getCEMC4Y()))
                sheet.update_cell(rowOfGoogleBackup,8,str(newNode.getCACL4Y()))
                sheet.update_cell(rowOfGoogleBackup,9,str(newNode.getCACL1Y()))
                rowOfGoogleBackup+=1
                #Used to find where the program left off when an exception is thrown
                sheet.update_cell(rowOfGoogle,1,"ENDED HERE")
            rowOfGoogle = rowOfGoogleBackup
            rowOfExcel = rowOfGoogle-1
    #Read from book, put into maxHeap
    browser.quit()

def main():
    #Name of the excel with ticker of stocks in row 0 & market Cap in row 3
    #MUST BE A .XLS FILE!!!
    excelName = "./All_stocks_Excel/March24_2019Nasdaq.xls"
    #Name of google sheet
    googleSheetName = "NasdaqValueFinder"
    '''
    NOTE: THE INDEX OF THE EXCEL STARTS AT 0, AND THE INDEX OF THE GOOGLE SHEET STARTS AT 1
    '''
    #COUNT FOR HOW MANY ROWS YOU WANT TO CHECK HERE.
    numberOfRowsCompared=3462

    writeToExcelCalculateValues(excelName, numberOfRowsCompared, googleSheetName)

if __name__ == "__main__":
    main()
