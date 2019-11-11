import csv
import os
import datetime
import logging
import shutil

# Assign the current datetime
today = datetime.date.today()
currentDate = today.strftime('%m%Y')
perPost = today.strftime('%Y%m')
currentDateFull = today.strftime('%Y%m%d')
# Setting up the accounting file directories
zdrive = '\\\\otofileserv\\accounting'
aiq2_folder = 'AIQ2'
exFolder = 'Excel Files'
archiveFolder = 'Archive'

# create dir that points to excel/archive folder
excelDir = os.path.join(zdrive, os.sep, aiq2_folder, currentDate, exFolder)
archDir = os.path.join(zdrive, os.sep, aiq2_folder, currentDate, archiveFolder)

# assign directory to specific location
os.chdir(excelDir)

# initiate logging
logging.basicConfig(filename='Logfile.log', level=logging.INFO,format='%(asctime)s:%(message)s')

# create empty lists for later use
headerDict = {}
headerList = []
IQDict = {}
IQList = []

# for each file in directory, look at extension, if its a csv, execute rest of script
for filename in os.listdir():
    file, file_ext = os.path.splitext(filename)
    if file_ext == '.csv':

        # Open CSV file, assign main fields to dictionary values
        with open(filename, 'r') as cfile:
            CSVFILE = csv.DictReader(cfile, delimiter=',')

            for row in CSVFILE:

                headerDict = {

                    'Level0': 'Level0',
                    'batnbr': '',
                    'jrnltype': 'IQ',
                    'battype': 'N',
                    'perpost': perPost,
                    'batchhandle': 'B',
                    'Actual': 'ACTUAL',
                    'blank1': '',
                    'blank2': '',
                    'blank3': '',
                    'total': '',
                }

                IQDict = {

                    'Level1': 'Level1',
                    'Company': row['Company'],
                    'ChkCompany': row['ChkCompany'],
                    'ActualCompany': row['ChkCompany'],
                    'GLAcct': row['GLAcct'],
                    'blank1': '',
                    'blank2': '',
                    'GLSubAcct': row['GLSubAcct'],
                    'RefNbr': row['RefNbr'],
                    'TranDate': row['TranDate'],
                    'GLAmount1': '',
                    'GLAmount2': '',
                    'Desc': row['Desc']

                }

                y = float(row['GLAmount'])
                y = ("%.2f" % y)
                row['GLAmount'] = str(y)

                x = list(headerDict.values())
                x[10] = row['GLAmount']
                headerList.append(x)

                f = open(excelDir + '\\' + row['ChkCompany'] + '.dta', 'w', newline='')
                f.write(str(','.join(headerList[0])) + '\n')
                f.close()
                headerList = []

                if row['IsCredit'] == 'false':
                    IQDict['GLAmount1'] = row['GLAmount']
                elif row['IsCredit'] == 'true':
                    if row['GLAmount'][0] == '-':
                        IQDict['GLAmount1'] = row['GLAmount']
                    else:
                        IQDict['GLAmount2'] = row['GLAmount']

                IQList.append(list(IQDict.values()))

        cfile.close()

for filename in os.listdir():
    file, file_ext = os.path.splitext(filename)
    if file_ext == '.dta':
        for item in IQList:
            if item[1] and item[3] == file:
                f = open(filename, 'a', newline='')
                f.write(str(','.join(item[0:2])) + ',' + str(','.join(item[4:])) + '\n')
                f.close()
            elif not item[1] and item[3] == file:
                f = open(filename, 'a', newline='')
                f.write(str(''.join(item[0])) + ',' + str(','.join(item[3:])) + '\n')
                f.close()
        os.rename(filename, 'IJ-' + file + '-' + currentDateFull + '_IQ' + file_ext)

for filename in os.listdir():
    file, file_ext = os.path.splitext(filename)
    if file_ext == '.dta':
        shutil.move(filename, "S:\\Transaction Import\\Inbound Folder\\" + filename)
        logging.info(' Transported ' + filename +
                    ' to otofileserv\Transaction Import\Inbound Folder')

for filename in os.listdir():
    file, file_ext = os.path.splitext(filename)
    if file_ext == '.csv':
        shutil.move(filename, archDir + '\\' + filename)
        logging.info(' Transported ' + filename + ' to archive')