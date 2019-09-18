import re, datetime, os, time, shutil, pyperclip, logging

today = datetime.date.today()
folderString = today.strftime('%m%Y')
folderString2 = today.strftime('%Y%m')

zdrive = 'Z:'
sdrive = 'S:'
CAfolder = 'CashAccrual'
excelFolder = 'Excel Files'
archiveFolder = 'Archive'
transactionFolder = 'Transaction Import'
inboundFolder = 'Inbound Folder'

inboundDir = os.path.join(sdrive, os.sep, transactionFolder, inboundFolder)

excelDir = os.path.join(zdrive, os.sep, CAfolder, folderString, excelFolder)
if not os.path.exists(excelDir):
    os.makedirs(excelDir)

archiveDir = os.path.join(zdrive, os.sep, CAfolder, folderString, archiveFolder)
if not os.path.exists(archiveDir):
    os.makedirs(archiveDir)

os.chdir(excelDir)
logging.basicConfig(filename='Logfile.log', level=logging.INFO,format='%(asctime)s:%(message)s')

nameList = []
newList = []
# create list of files with .csv extension
for filename in os.listdir():
    if filename[-4:] == '.csv':
        nameList.append(filename)

# read the current csv file's data
for n in range(0, len(nameList)):
    with open(nameList[n], "r") as file:
        data = file.readlines()

    firstList = []

# uses regex to eliminate unnecessary symbols/characters
    for i in data:
        i = re.sub(r'"*\W((\d+),)*(\d+)(\.\d+)\W"*', r'\2\3\4', i)
        i = re.sub(r'\W-\W\W\W','', i)
        firstList.append(i.split(','))
    file.close()

#constructs a batch file line depending on the characteristics of that line

    sList = []
    for i in firstList[1:]:
        if i[1]:
            if i[2] == '21000':
                x = 'Level1' + ','.join(i)
                y = x.split('\n')
                sList.append(y[0].split(','))
                d = re.sub(r'(\d+)\/(\d+)\/(\d+)', r'\3', i[7])
                g = re.sub(r'(\d+)\/(\d+)\/(\d+)', r'\1', i[7])
                e = re.sub(r'(\d+)\/(\d+)\/(\d+)', r'\2', i[7])
                if int(g) < 10 and int(g) > 0:
                    g = '0' + g
                if int(e) < 10 and int(e) > 0:
                    e = '0' + e
                z = "Level0,,GJ,N," + d + g + ",B,ACTUAL,,,," + i[-3]
                sList.insert(0, z.split(','))
                newList.append(sList)
                sList = []
            else:
                x = 'Level1' + ','.join(i)
                y = x.split('\n')
                sList.append(y[0].split(','))
        else:
            pass
for i in newList:
    f = open("IJ-" + i[1][1] + '-' + d + g + e + "_GJ.dta", "w")
    for j in i:
        f.write(','.join(j) + '\n')
    f.close()

for filename in os.listdir():
    if filename[-4:] == '.csv':
        shutil.move(excelDir + '\\' + filename, archiveDir + '\\' + filename)
        logging.info("Moved CSV file to Archive folder")
    elif filename[-4:] == '.dta':
        shutil.move(excelDir + '\\' + filename, inboundDir + '\\' + filename)
        logging.info("Moved " + filename + " to Transaction Import/Inbound folder")
