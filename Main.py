import re
from openpyxl import Workbook

#Open .ctx file and read contents
readFile = open("Regex_Test", "r")
readList = readFile.read()
readFile.close()

#Define regex patterns
regexGroup = r"GmmiLinkContextFrameObject\s+.+\s+.+\s+.+\s+.+\s+.+\s+.+\s+.+\s+.+\s+.+\s+.+"
regexBit = r"%?[I|Q]\d{4}"
regexDevice = r"GmmiObject\s\"\D+\d{5}\""
regexDescription = r"\"Desc2\"\s\".+\""

#Declare and initialize variable for list of regex matches
matchList = re.findall(regexGroup, readList)

#Create Excel workbook and assign variable for active sheet
wb = Workbook()
ws1 = wb.active

listLength = len(matchList)

#Loop through the list of matches and find device info
for _ in range(listLength):
    rowNumber = _
    columnLetter = 'A'
    strRowNumber = str(rowNumber)
    matchItem1 = re.findall(regexBit, matchList[_])
    if len(matchItem1) != 0:
        valueColumnA = matchItem1[0]
        ws1["%s%s" % (columnLetter, strRowNumber)] = valueColumnA

    columnLetter = 'B'
    matchItem2 = re.findall(regexDevice, matchList[_])
    if len(matchItem2) != 0:
        print(matchItem2)
        valueColumnB = matchItem2[0]
        ws1["%s%s" % (columnLetter, strRowNumber)] = valueColumnB

    columnLetter = 'D'
    matchItem3 = re.findall(regexDescription, matchList[_])
    print(len(matchItem3))
    if len(matchItem3) != 0:
        valueColumnC = matchItem3[0]
        ws1["%s%s" % (columnLetter, strRowNumber)] = valueColumnC

#Save excel workbook
wb.save('Practice2.xlsx')