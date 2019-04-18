#For reading Google Sheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

class Node:
    def __init__(self, ticker, sector,name, numberFromGoogleSheet):
        self.ticker = ticker
        self.sector = sector
        self.name = name
        self.numberFromGoogleSheet = numberFromGoogleSheet
        self.next=None
    def getTicker(self):
        return self.ticker
    def getSector(self):
        return self.sector
    def getName(self):
        return self.name
    def getNumberFromSheet(self):
        return self.numberFromGoogleSheet
    def getNext(self):
        return self.next
    def setNext(self,node):
        self.next=node

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
        if (float(currentNode.getNumberFromSheet())<=float(newNode.getNumberFromSheet())):
            newNode.setNext(currentNode)
            return newNode
        elif currentNode.getNext()==None:
            currentNode.setNext(newNode)
            return currentNode
        else:
            currentNode.setNext(self.pushRecursion(currentNode.getNext(),newNode))
            return currentNode

    def push(self, node):
        #First Node in the list
        if (self.rootNode==None):
            self.rootNode = node
            self.size+=1
        elif(self.rootNode.getNext()==None):
            if self.rootNode.getNumberFromSheet()<node.getNumberFromSheet():
                tmpNode = self.rootNode
                self.rootNode = node
                self.rootNode.setNext(tmpNode)
                self.size+=1
            else:
                self.rootNode.setNext(node)
                self.size+=1
        else:
            self.rootNode = self.pushRecursion(self.rootNode,node)
            self.size+=1

    def pop(self):
        oldRoot = self.rootNode
        nodeAfterRoot = self.rootNode.getNext()
        self.rootNode = nodeAfterRoot
        self.size-=1
        return oldRoot

def openGoogleSpreadSheet(googleSheetName):
    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name("NYSEValueFinder.json", scope)
    client = gspread.authorize(creds)
    #Opens the google sheet
    sheet = client.open(googleSheetName).sheet1
    return sheet

def findBestCompaniesFromGoogleSheet(googleSheetName):
    sheet = openGoogleSpreadSheet(googleSheetName)
    #List of all the rows in the google Sheet.  It has to be done this way because
    #Reading each cell individually will cause the google api to kick me out every 500 cells read.
    allValues = sheet.get_all_values()
    maxStackBVMC4Y = MaxStack("BVMC4Y")
    maxStackBVMC1Y = MaxStack("BVMC1Y")
    maxStackCAMC4Y = MaxStack("CAMC4Y")
    maxStackCAMC1Y = MaxStack("CAMC1Y")
    maxStackCEMC4Y = MaxStack("CEMC4Y")
    maxStackCACL4Y = MaxStack("CACL4Y")
    maxStackCACL1Y = MaxStack("CACL1Y")
    googleRow = 1
    while googleRow < len(allValues)-1:
        ticker = allValues[googleRow][0]
        marketCap = allValues[googleRow][1]
        sector = allValues[googleRow][2]
        name = allValues[googleRow][3]
        if (len(name)>30):
            name = name[:30]
        BVMC4Y = allValues[googleRow][4]
        BVMC1Y = allValues[googleRow][5]
        CAMC4Y = allValues[googleRow][6]
        CAMC1Y = allValues[googleRow][7]
        CEMC4Y = allValues[googleRow][8]
        CACL4Y = allValues[googleRow][9]
        CACL1Y = allValues[googleRow][10]
        if ticker != " ":
            if float(BVMC4Y)>1.0:
                newNodeBVMC4Y = Node(ticker,sector,name,BVMC4Y)
                maxStackBVMC4Y.push(newNodeBVMC4Y)
            if float(BVMC1Y)>1.0:
                newNodeBVMC1Y = Node(ticker,sector,name,BVMC1Y)
                maxStackBVMC1Y.push(newNodeBVMC1Y)
            if float(CAMC4Y)>1.0:
                newNodeCAMC4Y = Node(ticker,sector,name,CAMC4Y)
                maxStackCAMC4Y.push(newNodeCAMC4Y)
            if float(CAMC1Y)>1.0:
                newNodeCAMC1Y = Node(ticker,sector,name,CAMC1Y)
                maxStackCAMC1Y.push(newNodeCAMC1Y)
            if float(CEMC4Y)>0.5:
                newNodeCEMC4Y = Node(ticker,sector,name,CEMC4Y)
                maxStackCEMC4Y.push(newNodeCEMC4Y)
            if float(CACL4Y)>2.5:
                newNodeCACL4Y = Node(ticker,sector,name,CACL4Y)
                maxStackCACL4Y.push(newNodeCACL4Y)
            if float(CACL1Y)>2.5:
                newNodeCACL1Y = Node(ticker,sector,name,CACL1Y)
                maxStackCACL1Y.push(newNodeCACL1Y)
        googleRow+=1
    print()
    print("Book Value to Market Cap last Four Years:")
    print('{:<8s}{:<32s}{:<24s}{:<12s}'.format("Ticker:","Name:","Sector: ","BVMC4Y:"))
    while maxStackBVMC4Y.getSize() > 0:
        topNode = maxStackBVMC4Y.pop()
        print('{:<8s}{:<32s}{:<24s}{:<12s}'.format(topNode.getTicker(),topNode.getName(),topNode.getSector(),topNode.getNumberFromSheet()))
    print()
    print()
    print("Book Value to Market Cap Last Year:")
    print('{:<8s}{:<32s}{:<24s}{:<12s}'.format("Ticker:","Name:","Sector: ","BVMC1Y:"))
    while maxStackBVMC1Y.getSize() > 0:
        topNode = maxStackBVMC1Y.pop()
        print('{:<8s}{:<32s}{:<24s}{:<12s}'.format(topNode.getTicker(),topNode.getName(),topNode.getSector(),topNode.getNumberFromSheet()))
    print()
    print()
    print("Current Assets To Market Cap Last Four Years:")
    print('{:<8s}{:<32s}{:<24s}{:<12s}'.format("Ticker:","Name:","Sector: ","CAMC4Y:"))
    while maxStackCAMC4Y.getSize() > 0:
        topNode = maxStackCAMC4Y.pop()
        print('{:<8s}{:<32s}{:<24s}{:<12s}'.format(topNode.getTicker(),topNode.getName(),topNode.getSector(),topNode.getNumberFromSheet()))
    print()
    print()
    print("Current Assets to Market Cap Last Year:")
    print('{:<8s}{:<32s}{:<24s}{:<12s}'.format("Ticker:","Name:","Sector: ","CAMC1Y:"))
    while maxStackCAMC1Y.getSize() > 0:
        topNode = maxStackCAMC1Y.pop()
        print('{:<8s}{:<32s}{:<24s}{:<12s}'.format(topNode.getTicker(),topNode.getName(),topNode.getSector(),topNode.getNumberFromSheet()))
    print()
    print()
    print("Cash and Cash Equivalents To Market Cap Last Four Years:")
    print('{:<8s}{:32s}{:<24s}{:<12s}'.format("Ticker:","Name","Sector: ","CEMC4Y:"))
    while maxStackCEMC4Y.getSize() > 0:
        topNode = maxStackCEMC4Y.pop()
        print('{:<8s}{:<32s}{:<24s}{:<12s}'.format(topNode.getTicker(),topNode.getName(),topNode.getSector(),topNode.getNumberFromSheet()))
    print()
    print()
    print("Current Assets to Current Liabilities Last Four Years:")
    print('{:<8s}{:32s}{:<24s}{:<12s}'.format("Ticker:","Name","Sector: ","CACL4Y:"))
    while maxStackCACL4Y.getSize() > 0:
        topNode = maxStackCACL4Y.pop()
        print('{:<8s}{:<32s}{:<24s}{:<12s}'.format(topNode.getTicker(),topNode.getName(),topNode.getSector(),topNode.getNumberFromSheet()))
    print()
    print()
    print("Current Assets to Current Liabilities Last Year:")
    print('{:<8s}{:32s}{:<24s}{:<12s}'.format("Ticker:","Name","Sector: ","CACL1Y:"))
    while maxStackCACL1Y.getSize() > 0:
        topNode = maxStackCACL1Y.pop()
        print('{:<8s}{:<32s}{:<24s}{:<12s}'.format(topNode.getTicker(),topNode.getName(),topNode.getSector(),topNode.getNumberFromSheet()))
def main():
    googleSheetName = "NYSEValueFinderMar31"
    findBestCompaniesFromGoogleSheet(googleSheetName)

if __name__ == "__main__":
    main()
