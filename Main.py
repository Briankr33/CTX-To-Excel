import re
from openpyxl import Workbook
import Test

file_to_read = 'Regex_Test'

#Open .ctx file and read contents
with open(file_to_read, 'r') as f:
    readList = f.read()

#Define regex patterns
regexGroup = r"GmmiLinkContextFrameObject\s+.+\s+.+\s+.+\s+.+\s+.+\s+.+\s+.+\s+.+\s+.+\s+.+"
regexBit = r"%?[I|Q]\d{4}"
regexDevice = r"GmmiObject\s\"\D+\d{5}\""
regexDescription = r"\"Desc2\"\s\".+\""

#Declare and initialize variable for list of regex matches
matchList = re.findall(regexGroup, readList)
print(matchList)

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
    print(matchItem1)
    if matchItem1:
        if len(matchItem1) > 1:
            valueColumnA = ", ".join(matchItem1)
        else:
            valueColumnA = matchItem1[0]
        print(valueColumnA)
        ws1["%s%s" % (columnLetter, strRowNumber)] = valueColumnA

    columnLetter = 'B'
    matchItem2 = re.findall(regexDevice, matchList[_])
    if matchItem2:
        valueColumnB = matchItem2[0]
        ws1["%s%s" % (columnLetter, strRowNumber)] = valueColumnB

    columnLetter = 'D'
    matchItem3 = re.findall(regexDescription, matchList[_])
    if matchItem3:
        valueColumnC = matchItem3[0]
        ws1["%s%s" % (columnLetter, strRowNumber)] = valueColumnC

#Save excel workbook
print("Code finished successfully.")
wb.save('Practice2.xlsx')
print("Workbook saved.")