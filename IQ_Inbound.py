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

newpath = 'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\Log Archive'
if not os.path.exists(newpath):
    os.makedirs(newpath)

sDir = 'S:/Transaction Import/Inbound Folder'
inDir = os.path.join(sDir)

os.chdir(inDir)
logging.basicConfig(filename='Logfile.log', level=logging.INFO,format='%(asctime)s:%(message)s')

def autohotkey():

    print(time.strftime("\n[%m/%d/%Y][%H:%M]") + " Attempting to execute AHK script...")
    logging.info('     Attempting to execute AHK script...')

    # Launch AutoHotKey script // Will use TIAUT00.exe to import DTA files into DynamicsSL
    try:
        os.system("C:/Users/TIImport/desktop/TI_Script.exe")
    except IOError as e:
        raise e
    print(time.strftime("\n[%m/%d/%Y][%H:%M]") + " Imported successfully!")
    logging.info('        ***   Imported Successfully !!   ***')
    os.chdir(r'S:\Transaction Import\Processed Folder')
    logging.info('     Moving log file...')

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
                else:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IJ Log\\Bad\\' + file)
            elif 'Level0,,IQ' in x:
                if 'Batch is out of balance' not in x and 'The Number of Errors detected was 0' in x:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IQ Log\\Good\\' + file)
                else:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\IQ Log\\Bad\\' + file)
            elif 'Level0,,GJ' in x:
                if 'Batch is out of balance' not in x and 'The Number of Errors detected was 0' in x:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\GJ Log\\Good\\' + file)
                else:
                    shutil.move('S:\\Transaction Import\\Processed Folder\\' + file,
                                'S:\\Transaction Import\\Processed Folder\\' + folderString + '\\GJ Log\\Bad\\' + file)
        elif file[0:13] == 'TI Automation':
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

    logging.info('     Log file moved.')
    print(time.strftime("\nChecking for more DTA files to import..."))
    logging.info("     Checking for more DTA files to import")

compList = []
logging.info(' [ BEGIN SCRIPT ]')

# If file has .dta extension, map out its name and copy to clipboard for script to use
for file in os.listdir():
    if file[-4:] == '.dta':
        k = re.sub(r'(\D{2}\-)(\w+\-)(\d{8})(\_\D)*.dta', r'\1\2\3', file)
        d = re.sub(r'(\D{2})\-(\w+)\-(\d{8})\_(\D)*.dta', r'\2', file)
        logging.info('   Now importing {}'.format(k))
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
logging.info('   No files to import. Exiting.')
time.sleep(1)
print("\nEnding automation script in 5 seconds...")
logging.info(' [ END SCRIPT ]')
time.sleep(5)
