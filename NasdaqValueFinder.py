import os
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import re
import signal

class searchCalculateBS:

    #TODO: Connect to excel, ask which indecies it should run?
    def __init__(self):
        print("What is the ticker?: ")
        tmpStr = sys.stdin.readline()
        self.ticker=""
        for char in tmpStr:
            if char.isalpha():
                self.ticker+=char
        #tmpStr removes the extra line from the input.
        path_to_chromedriver = '/Users/benprocknow/Downloads/chromedriver'
        self.browser = webdriver.Chrome(executable_path=path_to_chromedriver)
        self.balanceSheet = self.getBalanceSheet()

    def getBalanceSheet(self):
        URL = "https://www.nasdaq.com/symbol/"+ self.ticker+ '/financials?query=balance-sheet'
        #If self.browser doesn't load in 20 seconds, quit.
        signal.alarm(20)
        self.browser.get(URL)
        signal.alarm(0)
        #See if the page has loaded.
        timeout = 10
        try:
            WebDriverWait(self.browser, timeout).until(EC.visibility_of_element_located((By.XPATH, "//tr[th/@bgcolor='#E6E6E6']")))
        #Page didn't load
        except TimeoutException:
            print ("Timeout error")
            self.browser.quit()
        row=2
        #Contains the int of all asset prices
        BalanceSheet=[]
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
                    location = "//div[@class='genTable']/table/tbody/tr["+str(row)+"]/td["+str(index)+"]"
                    HTMLObject=self.browser.find_elements_by_xpath(location)
                    #Int of asset price
                    AssetInt = [x.text for x in HTMLObject]
                    tmpList.append(AssetInt[0])
                    index = index+1
                BalanceSheet.append(tmpList)
            row = row+1
        return BalanceSheet
        self.browser.quit()

    #Take list and find the numbers in each index.  Put the numbers in the index in a string
    #and sum the string with the numbers of the other strings.  Return the total value
    def getIntFromList(self,numYears, list):
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
                sumOfList+=int(tmpStr)
            except ValueError:
                pass
        return sumOfList

    #Find the sum of the assets over the last 4 years, find the mean
    def computeBookValueLastFourYears(self):
        #Gets the total Assets from the list inside of the list.
        totalAssetsList = self.balanceSheet[12]
        totalAssets = self.getIntFromList(4,totalAssetsList)/4
        goodwillList = self.balanceSheet[8]
        goodwill = self.getIntFromList(4,goodwillList)/4
        intangibleList = self.balanceSheet[9]
        intangible = self.getIntFromList(4,intangibleList)/4
        liabilityList = self.balanceSheet[22]
        liability = self.getIntFromList(4,liabilityList)/4
        bookValueOverMeanLast4Years = totalAssets-goodwill-intangible-liability
        return bookValueOverMeanLast4Years

    def computeBookValueThisYear(self):
        totalAssetsList = self.balanceSheet[12]
        totalAssets = self.getIntFromList(1, totalAssetsList)
        goodwillList = self.balanceSheet[8]
        goodwill = self.getIntFromList(1, goodwillList)
        intangibleList = self.balanceSheet[9]
        intangible = self.getIntFromList(1, intangibleList)
        liabilityList = self.balanceSheet[22]
        liability = self.getIntFromList(1, liabilityList)
        bookValueLastYear = totalAssets-goodwill-intangible-liability
        return bookValueLastYear

    #Takes into account 0.33*Inventory
    def computeCurrentAssetsOverFourYears(self):
        netCurrentAssetsList = self.balanceSheet[5]
        netCurrentAssets = self.getIntFromList(4, netCurrentAssetsList)/4
        inventoryList = self.balanceSheet[3]
        inventory = self.getIntFromList(4, inventoryList)/4
        liabilityList = self.balanceSheet[22]
        liability = self.getIntFromList(4, liabilityList)/4
        currentAssetsLastFourYears = netCurrentAssets-0.5*inventory-liability
        return currentAssetsLastFourYears

    def computeCurrentAssetsLastYear(self):
        netCurrentAssetsList = self.balanceSheet[5]
        netCurrentAssets = self.getIntFromList(1, netCurrentAssetsList)
        inventoryList = self.balanceSheet[3]
        inventory = self.getIntFromList(1, inventoryList)
        liabilityList = self.balanceSheet[22]
        liability = self.getIntFromList(1, liabilityList)
        currentAssetsLastYear = netCurrentAssets-0.5*inventory-liability
        return currentAssetsLastYear

    def computeCashAssetsFourYears(self):
        cashList = self.balanceSheet[0]
        cash = self.getIntFromList(cashList)
        shortTermList = self.balanceSheet[1]
        shortTerm = self.getIntFromList(shortTermList)
        liabilityList = self.balanceSheet[22]
        liability = self.getIntFromList(1, liabilityList)
        cashAssetsFourYears = cash+shortTerm-liability
        return cashAssetsFourYears

    def currentAssetsToCurrentLiabilitiesFourYears(self):
        netCurrentAssetsList = self.balanceSheet[5]
        netCurrentAssets = self.getIntFromList(4, netCurrentAssetsList)/4
        inventoryList = self.balanceSheet[3]
        inventory = self.getIntFromList(4, inventoryList)/4
        currentLiabilitiesList = self.balanceSheet[16]
        currentLiabilities = self.getIntFromList(4,currentLiabilitiesList)/4
        currentAssetsToLiabilitiesFourYears = (netCurrentAssets-0.5*inventory)/currentLiabilities
        return currentAssetsToLiabilitiesFourYears

    def currentAssetsToCurrentLiabilitiesLastYear(self):
        netCurrentAssetsList = self.balanceSheet[5]
        netCurrentAssets = self.getIntFromList(1, netCurrentAssetsList)
        inventoryList = self.balanceSheet[3]
        inventory = self.getIntFromList(1, inventoryList)
        currentLiabilitiesList = self.balanceSheet[16]
        currentLiabilities = self.getIntFromList(1,currentLiabilitiesList)
        currentAssetsToLiabilitiesFourYears = (netCurrentAssets-0.5*inventory)/currentLiabilities
        return currentAssetsToLiabilitiesFourYears

    def quitBrowser(self):
        self.browser.quit()

#TODO:  Connect to excel here, add searchCalculateBS.computeFromBalanceSheet to
#AVL tree/ max heap.  When done can grab whatever top results are desirable.
def main():
    WantToConnectExcelHere = searchCalculateBS()
    print("Book Value Ave 4 years: ",WantToConnectExcelHere.computeBookValueLastFourYears())
    print("Book Value for Last Year: ",WantToConnectExcelHere.computeBookValueThisYear())
    print("Current Assets Ave 4 years: ",WantToConnectExcelHere.computeCurrentAssetsOverFourYears())
    print("Current Assets Last Year: ",WantToConnectExcelHere.computeCurrentAssetsLastYear())
    print("Cash Assets Ave 4 years: ",WantToConnectExcelHere.computeCashAssetsFourYears())
    print("Current Assets to Current Liabilities Last Four Years: ", WantToConnectExcelHere.currentAssetsToCurrentLiabilitiesFourYears())
    print("Current Assets to Current Liabilities Last Year: ", WantToConnectExcelHere.currentAssetsToCurrentLiabilitiesLastYear())
    WantToConnectExcelHere.quitBrowser()


if __name__ == "__main__":
    main()
