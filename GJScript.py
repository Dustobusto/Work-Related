import datetime, os, time, pyperclip, logging, shutil

today = datetime.date.today()
folderString = today.strftime('%m%Y')

zdrive = 'Z:'
sdrive = 'S:'
CAfolder = 'CashAccrual'
excelFolder = 'Excel Files'
archiveFolder = 'Archive'
transactionFolder = 'Transaction Import'
inboundFolder = 'Inbound Folder'

#inboundDir = os.path.join(sdrive, os.sep, transactionFolder, inboundFolder)

excelDir = os.path.join(zdrive, os.sep, CAfolder, folderString, excelFolder)
if not os.path.exists(excelDir):
    os.makedirs(excelDir)

archiveDir = os.path.join(zdrive, os.sep, CAfolder, folderString, archiveFolder)
if not os.path.exists(archiveDir):
    os.makedirs(archiveDir)

os.chdir(excelDir)
logging.basicConfig(filename='Logfile.log', level=logging.INFO,format='%(asctime)s:%(message)s')

# if file with .xls extension exists in dir, load Excel
for filename in os.listdir():
    file, file_ext = os.path.splitext(filename)
    if file_ext == '.xls' and file + '.csv' in os.listdir():
        pass
    elif file_ext == '.xls' and file + '.csv' not in os.listdir():
        pyperclip.copy(filename[0:-4])
        time.sleep(5)
        try:
            os.startfile(filename)
            logging.info(' Starting Excel...')
        except IOError as e:
            raise e
        time.sleep(10)

# Launch AutoHotKey script to convert Excel files into CSV files.
        try:
            os.system('C:\\Users\\tiimport\\Desktop\\IQScript.exe')
        except IOError as e:
            raise e
        time.sleep(5)
        shutil.move(excelDir + '\\' + filename, archiveDir + '\\' + filename)
        if filename[0:-4] + '.csv' in os.listdir():
            logging.info(' ' + filename[0:-4] + '.csv was created and ' + filename + ' was moved to the archive folder.')
        else:
            logging.info(' Error: CSV file not found.')
        time.sleep(5)
    else:
        logging.info(" No Excel files remaining in folder. Script will be stopped.")

