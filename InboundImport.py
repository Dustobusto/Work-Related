import os, re, pyperclip, time, logging, shutil, datetime

#date and time
today = datetime.date.today()
folderString = today.strftime('%m%Y')

# lists to create folders for each accounting file type and their destination
logList = ['IJ Log', 'IQ Log', 'GJ Log', 'TX Log']
evalList = ['DTA Files', 'Good', 'Bad']

oldpath = 'S:\\Transaction Import\\Processed Folder\\'

# create path if doesn't exist
newpath = oldpath + folderString
if not os.path.exists(newpath):
    os.makedirs(newpath)

logpath = newpath + '\\Log Archive'
if not os.path.exists(logpath):
    os.makedirs(logpath)

def createPath(newpath, acctType, fileEval):
    createPath = newpath + '\\' + acctType + '\\' + fileEval
    return createPath
    
# function to create accounting file type folders if they don't exist
def newPathFunc(acctType, fileEval):
    createdPath = createPath(newpath, acctType, fileEval)
    if not os.path.exists(createdPath):
        os.makedirs(createdPath)
        
# execute function every time script is called
for i in logList
    for j in evalList:
        newPathFunc(i, j)
        
       
sDir = 'S:/Transaction Import/Inbound Folder'
inDir = os.path.join(sDir)

os.chdir(inDir)
logging.basicConfig(filename='Logfile.log', level=logging.INFO, format='%(asctime)s %(message)s', datefmt='[%I:%M %p %A, %B %d, %Y]')

# big automation function (hang on)
def autohotkey():

    logging.info('                                                       ')
    
    # Launch AutoHotKey script // Will use TIAUT00.exe to import DTA files into DynamicsSL
    try:
        os.system("C:/Users/TIImport/desktop/TI_Script.exe")
    except IOError as e:
        raise e
    os.chdir(r'S:\Transaction Import\Processed Folder')

    # For each file cycle, move the file to its respective destination based on file type and content
    for file in os.listdir():
        if file[-4:] == '.log'.upper():
            f = open(file, 'r')
            data_file = f.read()
            f.close()
            if 'Level0,,IJ' in data_file:
                if 'Batch is out of balance' not in data_file and 'The Number of Errors detected was 0' in data_file:
                    shutil.move(oldpath + file,
                                createPath(newpath, logList[0], evalList[1]) + '\\' + file)
                    logging.info('     ** IMPORT SUCCESSFUL **')
                else:
                    shutil.move(oldpath + file,
                                createPath(newpath, logList[0], evalList[2]) + '\\' + file)
                    logging.info('     @@ IMPORT FAILED @@')
            elif 'Level0,,IQ' in data_file:
                if 'Batch is out of balance' not in data_file and 'The Number of Errors detected was 0' in data_file:
                    shutil.move(oldpath + file,
                                createPath(newpath, logList[1], evalList[1]) + '\\' + file)
                    logging.info('     ** IMPORT SUCCESSFUL **')
                else:
                    shutil.move(oldpath + file,
                                createPath(newpath, logList[1], evalList[2]) + '\\' + file)
                    logging.info('     @@ IMPORT FAILED @@')
            elif 'Level0,,GJ' in data_file:
                if 'Batch is out of balance' not in data_file and 'The Number of Errors detected was 0' in data_file:
                    shutil.move(oldpath + file,
                                createPath(newpath, logList[2], evalList[1]) + '\\' + file)
                    logging.info('     ** IMPORT SUCCESSFUL **')
                else:
                    shutil.move(oldpath + file,
                                createPath(newpath, logList[2], evalList[2]) + '\\' + file)
                    logging.info('     @@ IMPORT FAILED @@')
            elif 'Level0,,TX' in data_file:
                if 'Batch is out of balance' not in data_file and 'The Number of Errors detected was 0' in data_file:
                    shutil.move(oldpath + file,
                                createPath(newpath, logList[3], evalList[1]) + '\\' + file)
                    logging.info('     ** IMPORT SUCCESSFUL **')
                else:
                    shutil.move(oldpath + file,
                                createPath(newpath, logList[3], evalList[2]) + '\\' + file)
                    logging.info('     @@ IMPORT FAILED @@')
        elif file[0:13] == 'TI Automation':
            f = open(file, 'r')
            data_file = f.readlines()
            f.close()
            if "Processing Error" in data_file:
                logging.info('     ## IMPORT FAILED ## Processing Error. Check logs for file status')
            shutil.move(oldpath + file,
                        logpath + file)
        elif file[-4:] == '.dta':
            f = open(file, 'r')
            data_file = f.readline()
            data_file = data_file.split(',')
            f.close()
            if data_file[2] == 'IQ':
                shutil.move(oldpath + file,
                            createPath(newpath, logList[1], evalList[0]) + '\\' + file)
            elif data_file[2] == 'IJ':
                shutil.move(oldpath + file,
                            createPath(newpath, logList[0], evalList[0]) + '\\' + file)
            elif data_file[2] == 'GJ':
                shutil.move(oldpath + file,
                            createPath(newpath, logList[2], evalList[0]) + '\\' + file)
            elif data_file[2] == 'TX':
                shutil.move(oldpath + file,
                            createPath(newpath, logList[3], evalList[0]) + '\\' + file)

    logging.info("                                             ")

companyList = []
logging.info(' [ BEGIN SCRIPT ]')

# If file has .dta extension, map out its name and copy to clipboard for script to use
for file in os.listdir():
    if file[-4:] == '.dta':
        completeFileName = re.sub(r'(\D{2}\-)(\w+\-)(\d{8})(\_\D)*.dta', r'\1\2\3', file)
        hotelID = re.sub(r'(\D{2})\-(\w+)\-(\d{8})(\_)*(\D)*.dta', r'\2', file)
        logging.info('   Attempting to import: {}'.format(completeFileName))
        if hotelID not in companyList:
            print(hotelID)
            companyList.append(hotelID)
            pyperclip.copy(companyList[len(companyList) - 1])
            autohotkey()
        else:
            pass

print("No files to import")
time.sleep(2)
print("Moving log files to their respective destinations...")
logging.info('   No files to process. Now exiting..')
time.sleep(1)
print("\nEnding automation script in 5 seconds...")
logging.info(' [ END SCRIPT ]')
time.sleep(5)
