import os, re, pyperclip, time, logging, shutil, datetime

today = datetime.date.today()
folderString = today.strftime('%m%Y')

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IJ Log\\DTA Files'
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IJ Log\\Good'
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IJ Log\\Bad'
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IQ Log\\DTA Files'
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IQ Log\\Good'
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IQ Log\\Bad'
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\GJ Log\\DTA Files'
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\GJ Log\\Good'
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\GJ Log\\Bad'
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\TX Log\\DTA Files'
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\TX Log\\Good'
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\TX Log\\Bad'
if not os.path.exists(newpath):
    os.makedirs(newpath)

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\Log Archive'
if not os.path.exists(newpath):
    os.makedirs(newpath)

sDir = 'S:/Transaction Import/Inbound Folder'
inDir = os.path.join(sDir)

os.chdir(inDir)
logging.basicConfig(filename='Logfile.log', level=logging.INFO, format='%(asctime)s %(message)s', datefmt='[%I:%M %p %A, %B %d, %Y]')

def autohotkey():

   # print(time.strftime("\n[%I:%M %p %A, %B %d, %Y]") + " Attempting to execute AHK script...")
   # logging.info('     Attempting to execute import automation script...')
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
            x = f.read()
            f.close()
            if 'Level0,,IJ' in x:
                if 'Batch is out of balance' not in x and 'The Number of Errors detected was 0' in x:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IJ Log\\Good\\' + file)
                    logging.info('     ** IMPORT SUCCESSFUL **')
                else:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IJ Log\\Bad\\' + file)
                    logging.info('     @@ IMPORT FAILED @@')
            elif 'Level0,,IQ' in x:
                if 'Batch is out of balance' not in x and 'The Number of Errors detected was 0' in x:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IQ Log\\Good\\' + file)
                    logging.info('     ** IMPORT SUCCESSFUL **')
                else:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IQ Log\\Bad\\' + file)
                    logging.info('     @@ IMPORT FAILED @@')
            elif 'Level0,,GJ' in x:
                if 'Batch is out of balance' not in x and 'The Number of Errors detected was 0' in x:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\GJ Log\\Good\\' + file)
                    logging.info('     ** IMPORT SUCCESSFUL **')
                else:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\GJ Log\\Bad\\' + file)
                    logging.info('     @@ IMPORT FAILED @@')
            elif 'Level0,,TX' in x:
                if 'Batch is out of balance' not in x and 'The Number of Errors detected was 0' in x:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\TX Log\\Good\\' + file)
                    logging.info('     ** IMPORT SUCCESSFUL **')
                else:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\TX Log\\Bad\\' + file)
                    logging.info('     @@ IMPORT FAILED @@')
        elif file[0:13] == 'TI Automation':
            f = open(file, 'r')
            x = f.readlines()
            f.close()
            if "Processing Error" in x:
                logging.info('     ## IMPORT FAILED ## Processing Error. Check logs for file status')
            shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                        'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\Log Archive\\' + file)
        elif file[-4:] == '.dta':
            f = open(file, 'r')
            x = f.readline()
            x = x.split(',')
            f.close()
            if x[2] == 'IQ':
                shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                            'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IQ Log\\DTA Files\\' + file)
            elif x[2] == 'IJ':
                shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                            'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IJ Log\\DTA Files\\' + file)
            elif x[2] == 'GJ':
                shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                            'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\GJ Log\\DTA Files\\' + file)
            elif x[2] == 'TX':
                shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                            'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\TX Log\\DTA Files\\' + file)

    #print(time.strftime("\nChecking for more DTA files to import..."))
    logging.info("                                             ")
    # logging.info("   Checking for more data files to process..")

compList = []
logging.info(' [ BEGIN SCRIPT ]')

# If file has .dta extension, map out its name and copy to clipboard for script to use
for file in os.listdir():
    if file[-4:] == '.dta':
        k = re.sub(r'(\D{2}\-)(\w+\-)(\d{8})(\_\D)*.dta', r'\1\2\3', file)
        d = re.sub(r'(\D{2})\-(\w+)\-(\d{8})(\_)*(\D)*.dta', r'\2', file)
        logging.info('   Attempting to import: {}'.format(k))
        if d not in compList:
            print(d)
            compList.append(d)
            pyperclip.copy(compList[len(compList) - 1])
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
