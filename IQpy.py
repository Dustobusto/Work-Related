''' scrapped script '''

import csv
import os
import re

# assign directory to specific location
os.chdir('Z:\\AIQ2\\122018\\Excel Files\\')

# function to eliminate unnecessary characters for GLAmount
def cleanAmount(glamount):
    glamount = re.sub(r'\$((\d+))(\.\d+)\W', r'\2\3', glamount)
    glamount = re.sub(r'\$((\d+),)*(\d+)(\.\d+)\W', r'\2\3\4', glamount)
    return glamount

# header row
headerDict = {
    'Level0': 'Level0',
    'batnbr': '',
    'jrnltype': 'IQ',
    'battype': 'N',
    'perpost': '',
    'batchhandle': 'R',
    'Actual': 'ACTUAL',
    'blank1': '',
    'blank2': '',
    'blank3': '',
    'total': '',
}

# Open CSV file, assign main fields to dictionary values
for filename in os.listdir():
    if filename[-4:] == '.csv':
        with open(filename, "r") as file:

            CSVFILE = csv.DictReader(file, delimiter=',')

            headerRow = []
            IQList = []
            x = []
            newList = []

            for row in CSVFILE:
                IQDict = {
                    'Level1': 'Level1',
                    'Comp': row['Company'],
                    'GLAcct': row['GLAcct'],
                    'cprojectid': '',
                    'ctaskid': '',
                    'GLSubAcct': row['GLSubAcct'],
                    'RefNbr': row['RefNbr'],
                    'TranDate': row['TranDate'],
                    'GLAmount1': '',
                    'GLAmount2': '',
                    'Desc': row['Desc'],
                    'IsCredit': row['IsCredit']
                }

        # If not credit, place GLAmount in the debit field placeholder

                if IQDict['IsCredit'] == 'false':
                    IQDict['GLAmount1'] = cleanAmount(row['GLAmount'])
                    IQDict['GLAmount2'] = ''

        # If credit, place GLAmount in the credit field placeholder

                elif IQDict['IsCredit'] == 'true':
                    IQDict['GLAmount1'] = ''
                    IQDict['GLAmount2'] = cleanAmount(row['GLAmount'])

        # If Company field is blank, then row is offset row with total $ amount.
        # Assign amount to headerRow and append to headerRow list
        # Append row to newList with other rows

                if not IQDict['Comp']:
                    IQDict['Comp'] = row['ChkCompany']
                    headerDict['total'] = cleanAmount(row['GLAmount'])
                    headerRow.append(list(headerDict.values()))
                    x = list(IQDict.values())
                    newList.append(x[0:-1])

        # Code to add row values to newList. Includes main body. Uses if-statement to
        # prevent repeat instances of the offset row being added. If offset row is detected,
        # append newList to IQList and reset newList for next batch

                x = list(IQDict.values())
                if x[0:-1] in newList:
                    IQList.append(newList)
                    newList = []
                else:
                    newList.append(x[0:-1])

        for i in range(0, len(headerDict)):
            file = open('IQFILE' + str(i) + '.csv', 'w')
            for a, b in zip(headerRow, IQList):
                    file.write(','.join(a) + '\n')
                    for i in b:
                        file.write(','.join(i) + '\n')
            file.close()

