import datetime, os, time, pyperclip, logging, shutil

today = datetime.date.today()
folderString = today.strftime('%m%Y')
#os.chdir("Z:\\AIQ2\\" + folderString + "\\Excel Files")
# create Excel sub-folder within current date folder

zdrive = 'Z:'
sdrive = 'S:'
AIQ2folder = 'AIQ2'
excelFolder = 'Excel Files'
iqFolder = 'IQ Files'
archiveFolder = 'Archive'

excelDir = os.path.join(zdrive, os.sep, AIQ2folder, folderString, excelFolder)

if not os.path.exists(excelDir):
    os.makedirs(excelDir)

# create sub-folders within current date folder
iqfolderDir = os.path.join(zdrive, os.sep, AIQ2folder, folderString, iqFolder)


# newpath = 'Z:\\AIQ2\\' + folderString + '\\IQ Files'
if not os.path.exists(iqfolderDir):
    os.makedirs(iqfolderDir)

aDir = os.path.join(zdrive, os.sep, AIQ2folder, folderString, archiveFolder)
# newpath = 'Z:\\AIQ2\\' + folderString + '\\Archive'
if not os.path.exists(aDir):
    os.makedirs(aDir)


os.chdir(excelDir)
logging.basicConfig(filename='Logfile.log', level=logging.INFO, format='%(asctime)s:%(message)s')

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
        shutil.move('Z:\\AIQ2\\' + folderString + '\\Excel Files\\' + filename,
                    'Z:\\AIQ2\\' + folderString + '\\Archive\\' + filename)
        if filename[0:-4] + '.csv' in os.listdir():
            logging.info(' ' + filename[0:-4] + '.csv was created and ' + filename + ' was moved to the archive folder.')
        else:
            logging.info(' Error: CSV file not found.')
        time.sleep(5)
    else:
        logging.info(" No Excel files remaining in folder. Script will be stopped.")

