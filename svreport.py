'''
Original names of vendor associated variables have been changed for confidentiality
'''

import pandas as pd
import csv, os, pysftp, shutil, datetime, logging
import smtplib

# date code
today = datetime.date.today()
folderString = today.strftime('%m-%d-%Y')

# archive code updates each day
newpath = '\\\\FileServ\\Payroll\\VENDOR\\Archive\\' + folderString
if not os.path.exists(newpath):
    os.makedirs(newpath)

svDir = '\\\\FileServ\\Payroll\\VENDOR'  # supervisor report dir
filenameDir = svDir + '\\' + 'Archive' + '\\' + folderString + '\\'  # points to archive folder for specific file names
talentDir = '/outbound/pde'  # dir in VENDOR's SFTP folder for new employee csv

# counter for changes made in CSV file; changes = upload new file to SFTP
changeCount = 0

os.chdir(svDir)
supervisorList = []

# log (it's better than bad - it's good)
logging.basicConfig(
    filename='Logfile.log',
    level=logging.INFO,
    format='%(asctime)s %(message)s',
    datefmt='[%I:%M %p | %B %d, %Y]'
)


# Email VENDOR representative of Supervisor update (if needed)
EMAIL_PASSWORD = os.environ.get('tiimport_PASS')
EMAIL_ADDRESS = os.environ.get('tiimport_USER')

subject = 'Changes made in Employee CSV file'
msg = 'Changes in the supervisor column were made to the CSV file. \
       Please update the supervisor information to reflect current data.'

def send_email(subject, msg):
    try:
        server = smtplib.SMTP('smtp.SERVERDOMAIN.com:587')
        server.ehlo()
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        message = 'Subject: {}\n\n{}'.format(subject, msg)
        server.sendmail(EMAIL_ADDRESS, 'TEST@VENDORDOMAIN.com', message)
        server.quit()
        print('Email successfully delivered')
    except:
        print('Email failed to deliver')

def attrFunc():
    for attr in dirStruct:
        if attr.filename[0:7] == 'NewEmp_':
            sftp.get(VENDORdir + '/' + attr.filename,
                     localpath=svDir + '\\' +
                               attr.filename,
                     preserve_mtime=True)
            print(' • Retrieved [' + str(attr.filename) + '] from SFTP server\n')
            logging.info('----                                      ')
            logging.info('» RETRIEVED              [ ' + str(attr.filename) + ' ]')
        else:
            pass


# function to create future file to be uploaded to SFTP (if needed) - writes file, creates headers, closes file.
def f3open(headers):
    f3 = open(svDir + '\\' + filename + '_New' + file_ext, 'w')
    f3.write(headers)
    f3.close()


# for each employee row in csv file, check for inconsistencies in supervisor code
# no inconsistencies = write row
# inconsistencies = change supervisor code to blank, write row, increment change counter
def rowIter():
    global changeCount
    if fRow[32] in supervisorList or fRow[32] == '':  # fRow[32] points to the supervisor code
        f3 = open(svDir + '\\' + filename + '_New' + file_ext, 'a')
        f3.write(','.join(fRow) + '\n')
        f3.close()
    else:
        fRow[32] = ''
        f3 = open(svDir + '\\' + filename + '_New' + file_ext, 'a')
        f3.write(','.join(fRow) + '\n')
        f3.close()
        changeCount += 1
    return (changeCount)


# look for supervisor report, convert file to dataframe, convert df to list, format list
# to contain proper supervisor codes, append codes to new list for future comparison with employee file

if "Supervisor.xlsx" in os.listdir():
    print("Supervisor report detected")
    # filename, file_ext = os.path.splitext(file)
    file = "Supervisor.xlsx"
    df = pd.read_excel(file, index_col=0)
    for id in df['Employee Number (Supervisor)'].dropna().to_list():
        id = int(id)
        id = str(id)
        if len(id) == 5:
            id = '0' + id
        elif len(id) == 4:
            id = '00' + id
        elif len(id) == 3:
            id = '000' + id
        supervisorList.append(id)
    shutil.move(file, filenameDir + file)
else:
    print("Supervisor report not detected (╯°□°)╯")
    logging.info("----                           ")
    logging.info("» Supervisor report not detected")
    logging.info("» Quitting program")
    logging.info("----                           ")
    quit()

# connect to VENDOR's SFTP server
hostName = 'sftp.VENDOR.com'
myuserName = ## censored ##
mypw = ##censored##

cnopts = pysftp.CnOpts()
cnopts.hostkeys = None

with pysftp.Connection(
        host=hostName,
        username=myuserName,
        password=mypw,
        port=22,
        cnopts=cnopts
) as sftp:
    print("\n*** Connection established ***\n")
    sftp.cwd(VENDORdir)
    dirStruct = sftp.listdir_attr()

    # scan SFTP folder for new employee file, download to local directory
    attrList = []
    for attr in dirStruct:
        attrList.append(attr.filename[0:7])
    if 'NewEmp_' in attrList:
        attrFunc()
    else:
        print('New Employee file not found. Quitting program.')
        logging.info('----                                      ')
        logging.info('New Employee file not found. Quitting program.')
        logging.info('----                                      ')
        quit()

    # read downloaded file using CSV reader, execute new file function, read each row in file
    os.chdir(svDir)
    for file in os.listdir():
        filename, file_ext = os.path.splitext(file)
        if file_ext == '.csv':
            with open(file, 'r') as f:
                reader = csv.reader(f)
                headers = next(f, '')
                f3open(headers)
                for fRow in reader:
                    rowIter()
                f.close()
                shutil.move(file, filenameDir + filename + '_nc' + file_ext)

                # if changes were made, upload new file to SFTP server to replace old file
                if changeCount > 0:
                    print(' • Changes detected in CSV file. Initiating re-upload process!\n')
                    logging.info('» CHANGES DETECTED ')
                    logging.info('» Initiating re-upload..')
                    os.rename(svDir + '\\' + filename + '_New' + file_ext, svDir + '\\' + file)
                    sftp.put(svDir + '\\' + file,
                             remotepath=talentDir + '/' + file,
                             preserve_mtime=True)
                    logging.info('» Upload Complete!')
                    logging.info('----                                      ')
                    shutil.move(file, filenameDir + filename + '_c' + file_ext)
                    send_email(subject, msg)
                else:
                    print('\n - No changes detected. File will not be uploaded\n')
                    logging.info('» No changes detected. File will not be uploaded.')
                    logging.info('----                                      ')
                    os.rename(svDir + '\\' + filename + '_New' + file_ext, svDir + '\\' + file)
                    shutil.move(file, filenameDir + filename + '_nc' + file_ext)
