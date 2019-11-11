import re, datetime, os, time, shutil, pyperclip, logging, csv

# Assign the current datetime
today = datetime.date.today()
currentDate = today.strftime('%m%Y')
perPost = today.strftime('%Y%m')
currentDateFull = today.strftime('%Y%m%d')
newTranDate = today.strftime('%m/%d/%Y')

zdrive = 'Z:'
sdrive = 'S:'
CAfolder = 'CashAccrual'
excelFolder = 'Excel Files'
archiveFolder = 'Archive'
transactionFolder = 'Transaction Import'
inboundFolder = 'Inbound Folder'

inboundDir = os.path.join(sdrive, os.sep, transactionFolder, inboundFolder)

excelDir = os.path.join(zdrive, os.sep, CAfolder, currentDate, excelFolder)
if not os.path.exists(excelDir):
    os.makedirs(excelDir)

archiveDir = os.path.join(zdrive, os.sep, CAfolder, currentDate, archiveFolder)
if not os.path.exists(archiveDir):
    os.makedirs(archiveDir)

os.chdir(excelDir)
logging.basicConfig(filename='Logfile.log', level=logging.INFO,format='%(asctime)s:%(message)s')

headerList = []
GJList = []


# If CSV exists, begin conversion to DTA
for filename in os.listdir():
    file, file_ext = os.path.splitext(filename)
    if file_ext == '.csv':
        with open(filename, 'r') as csvfile:
            CSVFILE = csv.DictReader(csvfile, delimiter=',')
            for row in CSVFILE:

                headerDict = {

                    'Level0': 'Level0',
                    'batnbr': '',
                    'jrnltype': 'GJ',
                    'battype': 'N',
                    'perpost': perPost,
                    'batchhandle': 'B',
                    'Actual': 'ACTUAL',
                    'blank1': '',
                    'blank2': '',
                    'blank3': '',
                    'total': '',
                }


                CADict = {

                    "BatNbr": "Level1",
                    "ChkCompany": row['ChkCompany'],
                    "GLAcct": row['GLAcct'],
                    "GLSite": row['GLSite'],
                    "GLActivity": row['GLActivity'],
                    "GLSubAcct": row['GLSubAcct'],
                    "RefNbr": row['RefNbr'],
                    "TranDate": row['TranDate'],
                    "Debit": row['Debit'],
                    "Credit": row['Credit'],
                    "Description": row['Description']
                }

                glSub = row['GLSubAcct']
                if row['GLSubAcct'] == '0.0' or row['GLSubAcct'] == '0':
                    glSub = '000000'

                glAcct = row['GLAcct']

                hList = list(headerDict.values())

                if list(CADict.values())[8] == '':
                    pass
                else:
                    if row['ChkCompany']:
                        x = list(CADict.values())[8]
                        x = float(x)
                        x = ('%.2f' % x)

                        hList[10] = str(x)
                        headerList.append(hList)
                        f = open(excelDir + '\\' + row['ChkCompany'] + '.dta', 'w', newline='')
                        f.write(str(','.join(headerList[0])) + '\n')
                        f.close()
                        headerList = []

                caList = list(CADict.values())
                caList[5] = glSub
                caList[2] = glAcct
                caList[7] = newTranDate

                if caList[8] != '':
                    x = caList[8]
                    x = float(x)
                    x = ('%.2f' % x)
                    caList[8] = str(x)
                GJList.append(caList)

        csvfile.close()

# Create DTA file based on company name
for filename in os.listdir():
    file, file_ext = os.path.splitext(filename)
    if file_ext == '.dta':
        for item in GJList:
            if item[1] == file:
                f = open(excelDir + '\\' + filename, 'a', newline='')
                f.write(str(','.join(item)) + '\n')
                f.close()
        os.rename(filename, 'IJ-' + file + '-' + currentDateFull + '_GJ' + file_ext)

# If file is dta, move to inbound folder. move csv to archive
for filename in os.listdir():
    file, file_ext = os.path.splitext(filename)
    if file_ext == '.dta':
        shutil.move(filename, "S:\\Transaction Import\\Inbound Folder\\" + filename)
        logging.info(' Transported ' + filename + ' to otofileserv\Transaction Import\Inbound Folder')
    elif file_ext == '.csv':
        shutil.move(filename, archiveDir + '\\' + filename)
        logging.info(' Transported ' + filename + ' to archive')
