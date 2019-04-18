import os
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
import requests
import time
import re
import signal
import xlrd
import heapq

'''
Overview:  This is an implementation of a stock screener.
Prints out the stocks in order from best to worst based on the computations from the balance sheets
of all the stocks.
Takes a excel file from nasdaq website with all stocks on it.  Goes through
the list and uses the ticker and market cap from the excel to use webscraping to find the balance
sheet of the ticker on www.nasdaq.com.  Takes the balance sheet and computes certain values that
I want as a part of the stock screener.  Takes these values and puts them in a max heap.  After the
specifed number of tickers has been looked through, pop these values off the top of the heap until the
heap is empty.
'''

'''
Problems:  getBalanceSheet has the most problems.  It sometimes scapes too fast and causes the servers on nasdaq
to kill the connection.  There are a lot of try catch blocks to restart the scape again, but I have not been successful
at scraping more than 50 tickers with no exception.  Functionally everything works but I think that the program
gets ahead of itself sometimes when working with the internet and nasdaq and stuff gets fucked up.
'''

'''
Node object holds the ticker, marketCap, balanceSheetList.
'''
class Node:
    def __init__(self, balanceSheet, ticker, marketCap):
        self.ticker = ticker
        self.marketCap = findMarketCap(marketCap)
        self.balanceSheetList = balanceSheet
        self.next=None
    def getBalanceSheet(self):
        return self.balanceSheetList
    def getMarketCap(self):
        return self.marketCap
    def getTicker(self):
        return self.ticker
    def getNext(self):
        return self.next
    def setNext(self,node):
        self.next=node

'''
Implemented as a linked list.  Stores nodes in order based on what attribute to look for
ie: currentAssets, the first link will be the node with the lowest current assets.
Returns None if there is no data available for the ticker
If the new node doesn't have a balance sheet or market cap, nothing is inserted.
Check the docs for getBalanceSheet and getMarketCap to find when these cases occur.
Possible inputs:
computeBookValueToMarketCapLastFourYears
computeBookValueToMarketCapLastYear
computeCurrentAssetsToMarketCapLastFourYears
computeCurrentAssetsToMarketCapLastYear
computeCashAssetsToMarketCapLastFourYears
currentAssetsToCurrentLiabilitiesFourYears
currentAssetsToCurrentLiabilitiesLastYears
'''
class MaxStack:
    def __init__(self, methodToSortBy):
        self.rootNode = None
        self.methodToSortBy = methodToSortBy
        self.size = 0

    def getMethodToSortBy(self):
        return self.methodToSortBy

    def getSize(self):
        return self.size

    def getRootNode(self):
        return self.rootNode

    def pushRecursion(self, currentNode, newNode):
        #Compare to the next Node, find the location of the new node
        if (self.methodToSortBy=='computeBookValueToMarketCapLastFourYears'):
            if (computeBookValueToMarketCapLastFourYears(currentNode)<computeBookValueToMarketCapLastFourYears(newNode)):
                newNode.setNext(currentNode)
                return newNode
            elif (currentNode.getNext()==None):
                currentNode.setNext(newNode)
                return currentNode
            else:
                currentNode.setNext(self.pushRecursion(currentNode.getNext(), newNode))
                return currentNode
        elif (self.methodToSortBy=='computeBookValueToMarketCapLastYear'):
            if (computeBookValueToMarketCapLastYear(currentNode)<computeBookValueToMarketCapLastYear(newNode)):
                newNode.setNext(currentNode)
                return newNode
            elif (currentNode.getNext()==None):
                currentNode.setNext(newNode)
                return currentNode
            else:
                currentNode.setNext(self.pushRecursion(currentNode.getNext(), newNode))
                return currentNode
        elif (self.methodToSortBy=='computeCurrentAssetsToMarketCapLastFourYears'):
            if (computeCurrentAssetsToMarketCapLastFourYears(currentNode)<computeCurrentAssetsToMarketCapLastFourYears(newNode)):
                newNode.setNext(currentNode)
                return newNode
            elif (currentNode.getNext()==None):
                currentNode.setNext(newNode)
                return currentNode
            else:
                currentNode.setNext(self.pushRecursion(currentNode.getNext(), newNode))
                return currentNode
        elif (self.methodToSortBy=='computeCurrentAssetsToMarketCapLastYear'):
            if (computeCurrentAssetsToMarketCapLastYear(currentNode)<computeCurrentAssetsToMarketCapLastYear(newNode)):
                newNode.setNext(currentNode)
                return newNode
            elif (currentNode.getNext()==None):
                currentNode.setNext(newNode)
                return currentNode
            else:
                currentNode.setNext(self.pushRecursion(currentNode.getNext(), newNode))
                return currentNode
        elif (self.methodToSortBy=='computeCashAssetsToMarketCapLastFourYears'):
            if (computeCashAssetsToMarketCapLastFourYears(currentNode)<computeCashAssetsToMarketCapLastFourYears(newNode)):
                newNode.setNext(currentNode)
                return newNode
            elif (currentNode.getNext()==None):
                currentNode.setNext(newNode)
                return currentNode
            else:
                currentNode.setNext(self.pushRecursion(currentNode.getNext(), newNode))
                return currentNode
        elif (self.methodToSortBy=='currentAssetsToCurrentLiabilitiesFourYears'):
            if (currentAssetsToCurrentLiabilitiesFourYears(currentNode)<currentAssetsToCurrentLiabilitiesFourYears(newNode)):
                newNode.setNext(currentNode)
                return newNode
            elif (currentNode.getNext()==None):
                currentNode.setNext(newNode)
                return currentNode
            else:
                currentNode.setNext(self.pushRecursion(currentNode.getNext(), newNode))
                return currentNode
        elif (self.methodToSortBy=='currentAssetsToCurrentLiabilitiesLastYear'):
            if (currentAssetsToCurrentLiabilitiesLastYear(currentNode)<currentAssetsToCurrentLiabilitiesLastYear(newNode)):
                newNode.setNext(currentNode)
                return newNode
            elif (currentNode.getNext()==None):
                currentNode.setNext(newNode)
                return currentNode
            else:
                currentNode.setNext(self.pushRecursion(currentNode.getNext(), newNode))
                return currentNode
    #If node doesn't have a balance sheet or market cap, return
    def push(self, node):
        if (node.getBalanceSheet() == None or node.getMarketCap() == None):
            return
        #Increase size because the new node will be added to the max heap
        self.size+=1
        #First Node in the list
        if (self.rootNode==None):
            self.rootNode = node
        #Check if the second node is null, then compare the first node to the new node.
        elif (self.rootNode.getNext()==None):
            if (self.methodToSortBy=='computeBookValueToMarketCapLastFourYears'):
                if (computeBookValueToMarketCapLastFourYears(self.rootNode)<computeBookValueToMarketCapLastFourYears(node)):
                    tmpNode = self.rootNode
                    self.rootNode = node
                    self.rootNode.setNext(tmpNode)
                else:
                    self.rootNode.setNext(node)
            elif (self.methodToSortBy=='computeBookValueToMarketCapLastYear'):
                if (computeBookValueToMarketCapLastYear(self.rootNode)<computeBookValueToMarketCapLastYear(node)):
                    tmpNode = self.rootNode
                    self.rootNode = node
                    self.rootNode.setNext(tmpNode)
                else:
                    self.rootNode.setNext(node)
            elif (self.methodToSortBy=='computeCurrentAssetsToMarketCapLastFourYears'):
                if (computeCurrentAssetsToMarketCapLastFourYears(self.rootNode)<computeCurrentAssetsToMarketCapLastFourYears(node)):
                    tmpNode = self.rootNode
                    self.rootNode = node
                    self.rootNode.setNext(tmpNode)
                else:
                    self.rootNode.setNext(node)
            elif (self.methodToSortBy=='computeCurrentAssetsToMarketCapLastYear'):
                if (computeCurrentAssetsToMarketCapLastYear(self.rootNode)<computeCurrentAssetsToMarketCapLastYear(node)):
                    tmpNode = self.rootNode
                    self.rootNode = node
                    self.rootNode.setNext(tmpNode)
                else:
                    self.rootNode.setNext(node)
            elif (self.methodToSortBy=='computeCashAssetsToMarketCapLastFourYears'):
                if (computeCashAssetsToMarketCapLastFourYears(self.rootNode)<computeCashAssetsToMarketCapLastFourYears(node)):
                    tmpNode = self.rootNode
                    self.rootNode = node
                    self.rootNode.setNext(tmpNode)
                else:
                    self.rootNode.setNext(node)
            elif (self.methodToSortBy=='currentAssetsToCurrentLiabilitiesFourYears'):
                if (currentAssetsToCurrentLiabilitiesFourYears(self.rootNode)<currentAssetsToCurrentLiabilitiesFourYears(node)):
                    tmpNode = self.rootNode
                    self.rootNode = node
                    self.rootNode.setNext(tmpNode)
                else:
                    self.rootNode.setNext(node)
            elif (self.methodToSortBy=='currentAssetsToCurrentLiabilitiesLastYear'):
                if (currentAssetsToCurrentLiabilitiesLastYear(self.rootNode)<currentAssetsToCurrentLiabilitiesLastYear(node)):
                    tmpNode = self.rootNode
                    self.rootNode = node
                    self.rootNode.setNext(tmpNode)
                else:
                    self.rootNode.setNext(node)
        else:
            self.rootNode = self.pushRecursion(self.rootNode,node)

    def pop(self):
        oldRoot = self.rootNode
        nodeAfterRoot = self.rootNode.getNext()
        self.rootNode = nodeAfterRoot
        return oldRoot

#Returns the market cap in an int from the inputted string
def findMarketCap(marketCap):
    marketCapStr=''
    multiplierFactor=0
    indexOfPeriod=0
    i=0
    #Get list of the digits in the marketCap & set multiplierFactor
    while(i<len(marketCap)):
        if (marketCap[i]=='n'):
            return None
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
        pass
    return marketCap*multiplierFactor/(10**(len(marketCapStr)-indexOfPeriod))

#Opens the nasdaq website and clicks on the cookie box so that the rest of the program isn't fucked.
def clickOnCookieBox(browser):
    browser.get("http://www.nasdaq.com")
    cookieClickLocation = ("//div[@id='cookieConsent']/a[@id='cookieConsentOK']")
    timeout = 25
    try:
        WebDriverWait(browser, timeout).until(EC.visibility_of_element_located((By.XPATH, cookieClickLocation)))
    except TimeoutException:
        print ("Timeout error")
        browser.quit()
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
        return getBalanceSheet(browser, ticker)
    #Test if there is any data for this symbol
    try:
        #Location of element that checks if there is any information on the nasdaq website about the ticker.
        noDataAvailableElement = browser.find_element_by_xpath('//*[@id="quotes_content_left_nodatatext"]/span/b')
        if (noDataAvailableElement.text=="There is currently no data for this symbol."):
            return None
    except Exception:
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
                    print("Index Error")
                    return getBalanceSheet(browser, ticker)
                #Sometime the websites get mad for sending too many requests too fast, take a break try again
                except Exception:
                    print("Exception")
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
Used for MaxStack
Finds the value of the balance sheet over the specified time frame
If the balance sheet or market cap are not available, return -1
'''
def computeBookValueToMarketCapLastFourYears(node):
    if node.getBalanceSheet() is None:
        return -1
    else:#Gets the total Assets from the list inside of the list.
        totalAssetsList = node.getBalanceSheet()[12]
        totalAssets = getIntFromList(4,totalAssetsList)/4
        goodwillList =  node.getBalanceSheet()[8]
        goodwill =  getIntFromList(4,goodwillList)/4
        intangibleList = node.getBalanceSheet()[9]
        intangible = getIntFromList(4,intangibleList)/4
        liabilityList = node.getBalanceSheet()[22]
        liability =  getIntFromList(4,liabilityList)/4
        bookValueOverMeanLast4Years = totalAssets-goodwill-intangible-liability
        if (node.getMarketCap()==None):
            return -1
        BVMK = bookValueOverMeanLast4Years/node.getMarketCap()
        return BVMK
def computeBookValueToMarketCapLastYear(node):
    if node.getBalanceSheet() is None:
        return -1
    totalAssetsList =  node.getBalanceSheet()[12]
    totalAssets =  getIntFromList(1, totalAssetsList)
    goodwillList =  node.getBalanceSheet()[8]
    goodwill =  getIntFromList(1, goodwillList)
    intangibleList =  node.getBalanceSheet()[9]
    intangible =  getIntFromList(1, intangibleList)
    liabilityList =  node.getBalanceSheet()[22]
    liability =  getIntFromList(1, liabilityList)
    bookValueLastYear = totalAssets-goodwill-intangible-liability
    if (node.getMarketCap()==None):
        return -1
    BVMK = bookValueLastYear/node.getMarketCap()
    return BVMK
    #Takes into account 0.5*Inventory
def computeCurrentAssetsToMarketCapLastFourYears(node):
    if node.getBalanceSheet() is None:
        return -1
    cashAndCashEquivList = node.getBalanceSheet()[0]
    cashAndCashEquiv =  getIntFromList(4, cashAndCashEquivList)
    shortTermInvestList = node.getBalanceSheet()[1]
    shortTermInvest = getIntFromList(4,shortTermInvestList)
    netRecievablesList = node.getBalanceSheet()[2]
    netRevievables = getIntFromList(4,netRecievablesList)
    inventoryList =  node.getBalanceSheet()[3]
    inventory =  getIntFromList(4, inventoryList)
    liabilityList =  node.getBalanceSheet()[22]
    liability =  getIntFromList(4, liabilityList)/4
    netCurrentAssets = cashAndCashEquiv+shortTermInvest+netRevievables+0.5*inventory
    currentAssetsLastFourYears = netCurrentAssets-liability
    if (node.getMarketCap()==None):
        return -1
    CAMK = currentAssetsLastFourYears/node.getMarketCap()
    return CAMK
def computeCurrentAssetsToMarketCapLastYear(node):
    if node.getBalanceSheet() is None:
        return -1
    cashAndCashEquivList = node.getBalanceSheet()[0]
    cashAndCashEquiv =  getIntFromList(1, cashAndCashEquivList)
    shortTermInvestList = node.getBalanceSheet()[1]
    shortTermInvest = getIntFromList(1,shortTermInvestList)
    netRecievablesList = node.getBalanceSheet()[2]
    netRevievables = getIntFromList(1,netRecievablesList)
    inventoryList =  node.getBalanceSheet()[3]
    inventory =  getIntFromList(1, inventoryList)
    liabilityList =  node.getBalanceSheet()[22]
    liability =  getIntFromList(1, liabilityList)
    netCurrentAssets = cashAndCashEquiv+shortTermInvest+netRevievables+0.5*inventory
    currentAssetsLastYear = netCurrentAssets-liability
    if (node.getMarketCap()==None):
        return -1
    CAMK = currentAssetsLastYear/node.getMarketCap()
    return CAMK
def computeCashAssetsToMarketCapLastFourYears(node):
    if node.getBalanceSheet() is None:
        return -1
    cashAndCashEquivList =  node.getBalanceSheet()[0]
    cashAndCashEquiv = getIntFromList(4,cashAndCashEquivList)
    shortTermList =  node.getBalanceSheet()[1]
    shortTerm =  getIntFromList(1, shortTermList)
    accountPayableList = node.getBalanceSheet()[13]
    accountPayable = getIntFromList(4,accountPayableList)
    shortTermDebtList = node.getBalanceSheet()[14]
    shortTermDebt = getIntFromList(4,shortTermDebtList)
    otherCurrentLiabilitiesList = node.getBalanceSheet()[15]
    otherCurrentLiabilities = getIntFromList(4,otherCurrentLiabilitiesList)
    currentLiabilities = accountPayable+shortTermDebt+otherCurrentLiabilities
    cashAssetsFourYears = cashAndCashEquiv+shortTerm-currentLiabilities
    if (node.getMarketCap()==None):
        return -1
    CAMK = cashAssetsFourYears/node.getMarketCap()
    return CAMK
def currentAssetsToCurrentLiabilitiesFourYears(node):
    if node.getBalanceSheet() is None:
        return -1
    cashAndCashEquivList = node.getBalanceSheet()[0]
    cashAndCashEquiv =  getIntFromList(4, cashAndCashEquivList)
    shortTermInvestList = node.getBalanceSheet()[1]
    shortTermInvest = getIntFromList(4,shortTermInvestList)
    netRecievablesList = node.getBalanceSheet()[2]
    netRevievables = getIntFromList(4,netRecievablesList)
    inventoryList =  node.getBalanceSheet()[3]
    inventory =  getIntFromList(4, inventoryList)
    accountPayableList = node.getBalanceSheet()[13]
    accountPayable = getIntFromList(4,accountPayableList)
    shortTermDebtList = node.getBalanceSheet()[14]
    shortTermDebt = getIntFromList(4,shortTermDebtList)
    otherCurrentLiabilitiesList = node.getBalanceSheet()[15]
    otherCurrentLiabilities = getIntFromList(4,otherCurrentLiabilitiesList)
    currentLiabilities = accountPayable+shortTermDebt+otherCurrentLiabilities
    netCurrentAssets = cashAndCashEquiv+shortTermInvest+netRevievables+0.5*inventory
    if (currentLiabilities==0):
        currentLiabilities=1
    currentAssetsToLiabilitiesFourYears = netCurrentAssets/currentLiabilities
    return currentAssetsToLiabilitiesFourYears
def currentAssetsToCurrentLiabilitiesLastYear(node):
    if node.getBalanceSheet() is None:
        return -1
    cashAndCashEquivList = node.getBalanceSheet()[0]
    cashAndCashEquiv =  getIntFromList(1, cashAndCashEquivList)
    shortTermInvestList = node.getBalanceSheet()[1]
    shortTermInvest = getIntFromList(1,shortTermInvestList)
    netRecievablesList = node.getBalanceSheet()[2]
    netRevievables = getIntFromList(1,netRecievablesList)
    inventoryList =  node.getBalanceSheet()[3]
    inventory =  getIntFromList(1, inventoryList)
    accountPayableList = node.getBalanceSheet()[13]
    accountPayable = getIntFromList(1,accountPayableList)
    shortTermDebtList = node.getBalanceSheet()[14]
    shortTermDebt = getIntFromList(1,shortTermDebtList)
    otherCurrentLiabilitiesList = node.getBalanceSheet()[15]
    otherCurrentLiabilities = getIntFromList(1,otherCurrentLiabilitiesList)
    currentLiabilities = accountPayable+shortTermDebt+otherCurrentLiabilities
    netCurrentAssets = cashAndCashEquiv+shortTermInvest+netRevievables+0.5*inventory
    if (currentLiabilities==0):
        currentLiabilities=1
    currentAssetsToLiabilitiesFourYears = netCurrentAssets/currentLiabilities
    return currentAssetsToLiabilitiesFourYears

'''
Acts as a main method.  Launches all helper methods to return a list
of Nodes.
Inputs excel sheet with ticker in zero row and market cap in third row
This sheet must be on sheet index 0
Inputs number of rows that should be scanned on the excel.  Input of 'all'
means that all the rows should be checked
Returns a list of the balance sheets of all the tickers
'''
def printMinHeapCalculateValues(excelSpreadsheetName, numberOfRowsCompared):
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
    nasdaqExcel = xlrd.open_workbook(excelSpreadsheetName)
    nasdaqSheet = nasdaqExcel.sheet_by_index(0)
    #Second row of Excel = index 1
    i=1
    BVMC4Y = MaxStack('computeBookValueToMarketCapLastFourYears')
    BVMC1Y = MaxStack('computeBookValueToMarketCapLastYear')
    CAMC4Y = MaxStack('computeCurrentAssetsToMarketCapLastFourYears')
    CAMC1Y = MaxStack('computeCurrentAssetsToMarketCapLastFourYears')
    CEMC4Y = MaxStack('computeCashAssetsToMarketCapLastFourYears')
    CAL4Y = MaxStack('currentAssetsToCurrentLiabilitiesFourYears')
    CAL1Y = MaxStack('currentAssetsToCurrentLiabilitiesLastYear')
    clickOnCookieBox(browser)
    while (i<numberOfRowsCompared+1):
        ticker = nasdaqSheet.row_values(i)[0]
        marketCap = nasdaqSheet.row_values(i)[3]
        balanceSheet = getBalanceSheet(browser, ticker)
        newNode = Node(balanceSheet,ticker,marketCap)
        BVMC4Y.push(newNode)
        newNode = Node(balanceSheet,ticker,marketCap)
        BVMC1Y.push(newNode)
        newNode = Node(balanceSheet, ticker, marketCap)
        CAMC4Y.push(newNode)
        newNode = Node(balanceSheet, ticker, marketCap)
        CAMC1Y.push(newNode)
        newNode = Node(balanceSheet, ticker, marketCap)
        CEMC4Y.push(newNode)
        newNode = Node(balanceSheet, ticker, marketCap)
        CAL4Y.push(newNode)
        newNode = Node(balanceSheet, ticker, marketCap)
        CAL1Y.push(newNode)
        i+=1
    browser.quit()
    #Print minHeap
    i=0
    while(i<BVMC4Y.getSize()):
        firstNode = BVMC4Y.pop()
        print("Ticker: "+firstNode.getTicker()+" Book Value to Market Cap 4 years: "+str(computeBookValueToMarketCapLastFourYears(firstNode)))
        i+=1
    print("\n")
    i=0
    while(i<BVMC1Y.getSize()):
        firstNode = BVMC1Y.pop()
        print ("Ticker: "+firstNode.getTicker()+" Book Value to Market Cap Last year: "+str(computeBookValueToMarketCapLastYear(firstNode)))
        i+=1
    print("\n")
    i=0
    while(i<CAMC4Y.getSize()):
        firstNode = CAMC4Y.pop()
        print("Ticker: "+firstNode.getTicker()+" Current Assets to Market Cap 4 Years: "+str(computeCurrentAssetsToMarketCapLastFourYears(firstNode)))
        i+=1
    print("\n")
    i=0
    while(i<CAMC1Y.getSize()):
        firstNode = CAMC1Y.pop()
        print("Ticker: "+firstNode.getTicker()+" Current Assets To Market Cap Last Year: "+str(computeCurrentAssetsToMarketCapLastYear(firstNode)))
        i+=1
    print("\n")
    i=0
    while(i<CEMC4Y.getSize()):
        firstNode = CEMC4Y.pop()
        print("Ticker: "+firstNode.getTicker()+" Cash Equivalent To Market Cap 4 Years: "+str(computeCashAssetsToMarketCapLastFourYears(firstNode)))
        i+=1
    print("\n")
    i=0
    while(i<CAL4Y.getSize()):
        firstNode = CAL4Y.pop()
        print("Ticker: "+firstNode.getTicker()+" Current Assets To Liabilities 4 years: "+str(currentAssetsToCurrentLiabilitiesFourYears(firstNode)))
        i+=1
    print("\n")
    i=0
    while(i<CAL1Y.getSize()):
        firstNode = CAL1Y.pop()
        print("Ticker: "+firstNode.getTicker()+" Current Assets To Liabilities Last Year: "+str(currentAssetsToCurrentLiabilitiesLastYear(firstNode)))
        i+=1
    print("\n")

def main():
    #Name of the excel with ticker of stocks in row 0 & market Cap in row 3
    #MUST BE A .XLS FILE!!!
    excelName = "./All_stocks_Excel/March20_2019Nasdaq.xls"
    #COUNT FOR HOW MANY ROWS YOU WANT TO CHECK HERE.
    numberOfRowsCompared=200

    printMinHeapCalculateValues(excelName, numberOfRowsCompared)

if __name__ == "__main__":
    main()
