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
from selenium.webdriver.common.keys import Keys
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

It is connected to the google project called NasdaqValueFinder

Takes input from an excel that was downloaded from www.nasdaq.com and it contains all of the stocks on the nasdaq.
The input is the ticker and the market cap of the stock.  The balance sheet is found from the ticker from
www.rocketfinancial.com, and all of the values that are used as part of the screener are calculated as a function
of the market cap and balance sheet.  There are nodes that store the tickers, market cap, balance sheet, and all
of the computed values.
The nodes are written to an excel after the balance sheet and computed values are found.

WAYS TO IMPROVE:
Find if the company is in a bubble.  Find the growth of debt over the life of the company compared to
its earnings.  Also could be found by growth of earnings vs market cap.

Find a way to take all the computed values and spit out less numbers, but each number is more important.  Maybe I
only want companies that have current assets to current liabilities over 2.5 and a good book value etc.  Could use machine
learning???

Version 1.03?  Include Income Statement for each company and find the ROA, ROI, Return on fixed assets, depreciation costs
relative to assets, follow Security Analysis.  Includes Cash Flow Statement!


Version 2.00?  Saves all the balance sheets for each company in an array or something in the google sheet for the last decade+
and finds how the debt of all industrials/health companies/financials/ etc have been growing relative to earnings.
"BubbleFinder.py".  Is based off of the Pricipals of Navigating Debt Crisises.  Mostly macro
'''

'''
Problems:
'''

'''
Node object holds the ticker, marketCap, balanceSheetList.
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
    #Returns the market caps that are below one million dollars
    try:
        outputMC = ""
        for i in marketCap[1:]:
            if i!=",":
                outputMC+=i
        return int(outputMC)
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
    if indexOfPeriod==0:
        return marketCap*multiplierFactor
    return marketCap*multiplierFactor/(10**(len(marketCapStr)-indexOfPeriod))

def clickOnNoThanks(browser):
    pass

'''
Opens the nasdaq website and clicks on the cookie box so that the rest of the program isn't fucked.

URL: https://www.sec.gov/search/search.htm
'''
def openSecTickerAllYearsBalanceSheet(browser, ticker):
    URL = "https://www.sec.gov/search/search.htm"
    browser.get(URL)
    timeout = 5
    loadedPage = False
    makeSurePageIsOpened(browser, '//*[@id="cik"]', timeout, "Location of the text box was not found")
    searchBox = browser.find_element_by_xpath('//*[@id="cik"]')
    searchBox.send_keys(ticker)
    searchBox.send_keys(Keys.RETURN)
    #See if there is no matching ticker symbol
    try:
        #Location of the first interactive data button
        WebDriverWait(browser, timeout).until(EC.visibility_of_element_located((By.XPATH, '/html/body/div/center/h1')))
    #There is data on this ticker
    except Exception:
        pass
    return True

'''
Goes row by row through the first page with all of the filings for the current company
and finds the first row that contains a Interactive Data button.
Returns the type of filing that is in that row
'''
def getFileType(browser):
    timeout = 2
    makeSurePageIsOpened(browser, '//*[@id="count"]/option[5]', timeout, "Location of the limit results per page not found.")
    #Make the results per page 100
    limitResultsPerPageElement = browser.find_element_by_xpath('//*[@id="count"]/option[5]')
    limitResultsPerPageElement.click()
    currentRow = 0
    while(currentRow<100):
        try:
            #This is the element for the interactive button
            browser.find_element_by_xpath('//*[@id="seriesDiv"]/table/tbody/tr['+str(currentRow)+']/td[2]/a[2]')
            filingType = browser.find_element_by_xpath('//*[@id="seriesDiv"]/table/tbody/tr['+str(currentRow)+']/td[1]')
            return filingType.text
        #The current row doesn't contain an interactive data button
        except Exception:
            currentRow+=1
    #There were no interactive links in the first page, return None
    return None

'''
Helper method for getBalanceSheet
Make sure the page is fully opened
'''
def makeSurePageIsOpened(browser, element, timeout, exceptionMessage):
    while(True):
        try:
            #Location of the first interactive data button
            WebDriverWait(browser, timeout).until(EC.visibility_of_element_located((By.XPATH, element)))
            break
        except TimeoutException:
            print("Make sure page is opened timeout Exception")
            pass
        except Exception:
            print(exceptionMessage)
            #return None

def getFilingDateYear(browser, currentRow):
    filingYearElement = browser.find_element_by_xpath('//*[@id="seriesDiv"]/table/tbody/tr['+str(currentRow)+']/td[4]')
    return filingYearElement.text[:4]

'''
Returns the balance sheet of the ticker from www.sec.gov
Wait for the balance sheet numbers to be visible.
If the ticker is not available or doesn't have any information on the website,
return None
'''
def getBalanceSheet(browser, ticker):
    #Open the website and get the browser to the page where all of the filings are.
    if openSecTickerAllYearsBalanceSheet(browser, ticker) is False:
        return None
    #Find the type of filing that the current ticker files under.
    fileType = getFileType(browser)
    if fileType is None:
        return None
    #Enter in the fileType to the Filing Type text box.
    filingTypeBoxElement = browser.find_element_by_xpath('//*[@id="type"]')
    filingTypeBoxElement.send_keys(str(fileType))
    filingTypeBoxElement.send_keys(Keys.RETURN)
    #Loop through all of the interactive data links and get the balance sheet from each
    currentWhiteRow = 2
    currentBlueRow = 1
    boolWhiteRow = True
    #This dictionary will contain all of the balance sheet infomation for the current ticker
    balanceSheetDictionary ={}
    timeout=2
    while(True):
        #Click on the current row interactive data button
        #The white and blue rows switch off
        if (boolWhiteRow):
            #Wait for new page to load
            makeSurePageIsOpened(browser, '//*[@id="seriesDiv"]/table/tbody/tr['+str(currentWhiteRow)+']/td[2]/a[2]' , timeout, "Interactive Data Link is broken for "+ticker)
            try:
                interactiveDataElement = browser.find_element_by_xpath('//*[@id="seriesDiv"]/table/tbody/tr['+str(currentWhiteRow)+']/td[2]/a[2]')
                #There is a delay between this element being found and being able to be clicked on.
                while(True):
                    try:
                        balanceSheetDictionary[str(getFilingDateYear(browser, currentWhiteRow))] = {}
                        interactiveDataElement.click()
                        currentWhiteRow+=2
                        boolWhiteRow = False
                        break
                    except Exception:
                        interactiveDataElement = browser.find_element_by_xpath('//*[@id="seriesDiv"]/table/tbody/tr['+str(currentWhiteRow)+']/td[2]/a[2]')
                        print("Interactive Balance Sheet Exception, WR")
                        time.sleep(0.05)
                        pass
            #There are no more interactive data links
            except Exception:
                print("No more interactive data links Exception")
                break
        #Blue row
        else:
            makeSurePageIsOpened(browser, '//*[@id="seriesDiv"]/table/tbody/tr['+str(currentBlueRow)+']/td[2]/a[2]', timeout, "Location of the interactive data button for the blue row didn't load for "+ticker)
            try:
                interactiveDataElement = browser.find_element_by_xpath('//*[@id="seriesDiv"]/table/tbody/tr['+str(currentBlueRow)+']/td[2]/a[2]')
                boolWhiteRow = True
                currentBlueRow+=1
                #There is a delay between this element being found and being able to be clicked on.
                while(True):
                    try:
                        balanceSheetDictionary[getFilingDateYear(browser, currentBlueRow)] = {}
                        interactiveDataElement.click()
                        break
                    except Exception:
                        interactiveDataElement = browser.find_element_by_xpath('//*[@id="seriesDiv"]/table/tbody/tr['+str(currentBlueRow)+']/td[2]/a[2]')
                        print("Interactive Balance Sheet Exception, BR")
                        time.sleep(0.05)
                        pass
            #There are no more interactive data links
            except Exception:
                print("No mroe interactive data links Exception")
                break
        makeSurePageIsOpened(browser, '//*[@id="menu_cat2"]', timeout, "Finincial Statements Button didn't load")
        financialStatementsElement = browser.find_element_by_xpath('//*[@id="menu_cat2"]')
        financialStatementsElement.click()
        makeSurePageIsOpened(browser, '//*[@id="r2"]/a', timeout, "There are no selections for the balance sheet for "+ticker)
        #Finding the Consolidated Balance Sheets button
        currentRowFinancial = 2
        while(True):
            try:
                balanceSheetElement = browser.find_element_by_xpath('//*[@id="r'+str(currentRowFinancial)+'"]/a')
            #There is no consolidated balance sheet for this ticker, try the next block
            except Exception:
                print("Consolidated balance sheet element not found for "+ticker)
                break
            if balanceSheetElement.text == "CONSOLIDATED BALANCE SHEETS" or balanceSheetElement.text == "Consolidated Balance Sheets":
                balanceSheetElement.click()
                getConsolidatedBalanceSheet(browser, balanceSheetDictionary)
                break
            elif balanceSheetElement.text == "OTHER BALANCE SHEET NAMES":
                balanceSheetElement.click()
                getOTHERBALANCESHEET(browser, balanceSheetDictionary)
            else:
                currentRowFinancial+=1
        #Try to find the other types of balance sheets - see General Electric

    return balanceSheetDictionary

'''
Helper method for computeXXXXLastYear
#Take list and find the numbers in each index.  Put the numbers in the index in a string
#and sum the string with the numbers of the other strings.  Return the total value
'''
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
Goes through every row in the excel sheet
Inputs number of rows that should be scanned on the excel.  Input of 'all'
means that all the rows should be checked
Returns a list of the balance sheets of all the tickers
'''
def writeToExcelCalculateValues(excelSpreadsheetName, googleSheetName):
    #Operator of the browser
    browser = configureBrowser()
    #Opens the excel that contains the market cap and the ticker
    nasdaqExcel = xlrd.open_workbook(excelSpreadsheetName)
    nasdaqSheet = nasdaqExcel.sheet_by_index(0)
    #Find the number of rows in the EXCEL
    numberOfRowsCompared = nasdaqSheet.nrows
    #Operator of the google sheet that is used to push all of the calculated values
    sheet = openGoogleSpreadSheet(googleSheetName)
    #Finds the row that the program left off on after it was discontinued
    rowOfGoogle = findRowOfGoogle(sheet)
    rowOfExcel = rowOfGoogle-1
    #SOMETIMES THERE IS A POPUP THAT APPEARS, NOT DONE YET BUT IF IT IS A PROPLEM TAKE CARE OF IT HERE
    clickOnNoThanks(browser)
    #Loop through the rows of the excel sheet and the google Sheet
    while (rowOfExcel<numberOfRowsCompared):
        #Used if the google api says that the program is writing to the google sheet too quickly.
        rowOfGoogleBackup = rowOfGoogle
        ticker = str(nasdaqSheet.row_values(rowOfExcel)[0])
        marketCap = nasdaqSheet.row_values(rowOfExcel)[3]
        sector = nasdaqSheet.row_values(rowOfExcel)[5]
        name = nasdaqSheet.row_values(rowOfExcel)[1]
        #Restart the browser if getBalanceSheet threw an exception.  The program usually throws Timeout
        #Exceptions from retrieving the webpage, but sometimes google API kicks me out for writing over
        #500 cells/100 seconds, that's why there is the time.sleep
        createdNode = False
        while(createdNode is False):
            newNode = Node(browser,ticker,marketCap)
            try:
                newNode = Node(browser, ticker, marketCap)
                createdNode = True
            except Exception:
                print("newNodeException")
                browser.quit()
                time.sleep(10)
                browser = configureBrowser()
        try:
            if (newNode.getBalanceSheet() is None or newNode.getMarketCap() is None):
                    #sheet.update_cell(rowOfGoogle,1,"")
                    rowOfExcel+=1
                    rowOfGoogle+=1
                    #sheet.update_cell(rowOfGoogle,1,"ENDED HERE")
            else:
                sheet.update_cell(rowOfGoogle,2,str(newNode.getMarketCap()))
                sheet.update_cell(rowOfGoogle,3,str(sector))
                sheet.update_cell(rowOfGoogle,4,str(name))
                sheet.update_cell(rowOfGoogle,5, str(newNode.getBVMC4Y()))
                sheet.update_cell(rowOfGoogle,6,str(newNode.getBVMC1Y()))
                sheet.update_cell(rowOfGoogle,7,str(newNode.getCAMC4Y()))
                sheet.update_cell(rowOfGoogle,8,str(newNode.getCAMC1Y()))
                sheet.update_cell(rowOfGoogle,9,str(newNode.getCEMC4Y()))
                sheet.update_cell(rowOfGoogle,10,str(newNode.getCACL4Y()))
                sheet.update_cell(rowOfGoogle,11,str(newNode.getCACL1Y()))
                #Used to find where the program left off when an exception is thrown
                sheet.update_cell(rowOfGoogle+1,1,"ENDED HERE")
                sheet.update_cell(rowOfGoogle,1,str(ticker))
                rowOfExcel+=1
                rowOfGoogle+=1
        #Sometimes an exception is thrown that requires the program to be restarted.
        except gspread.exceptions.APIError:
            print("Unauthenticated Exception")
            browser.quit()
            time.sleep(10)
            browser = configureBrowser()
            clickOnCookieBox(browser)
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
    excelName = "./All_stocks_Excel/March27_2019Nasdaq.xls"
    #Name of google sheet
    googleSheetName = "NasdaqValueFinderMar29"
    '''
    NOTE: THE INDEX OF THE EXCEL STARTS AT 0, AND THE INDEX OF THE GOOGLE SHEET STARTS AT 1
    '''

    writeToExcelCalculateValues(excelName, googleSheetName)

if __name__ == "__main__":
    main()
