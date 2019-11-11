import pandas as pd
import datetime, os
import logging
import shutil

# Current month and year will name the folder
today = datetime.date.today()
folderString = today.strftime('%m%Y')

# Setting up the accounting file directories
zdrive = '\\\\otofileserv\\accounting'
caFolder = 'CashAccrual'
exFolder = 'Excel Files'
archiveFolder = 'Archive'

excelDir = os.path.join(zdrive, os.sep, caFolder, folderString, exFolder)
archDir = os.path.join(zdrive, os.sep, caFolder, folderString, archiveFolder)

# If excel file and archive dir don't exist, create them
if not os.path.exists(excelDir):
    os.makedirs(excelDir)
if not os.path.exists(archDir):
    os.makedirs(archDir)

# Points to Excel Files as current folder directory
os.chdir(excelDir)

# If .xls file exists in excel directory without a .csv, execute code to convert file
for filename in os.listdir():
    file, file_ext = os.path.splitext(filename)
    if file_ext == '.xls' and file + '.csv' in os.listdir():
        pass
    elif file_ext == '.xls' and file + '.csv' not in os.listdir():
        logging.info(' Converting from xls to csv')
        xls_file = pd.read_excel(filename, index_col=None)
        xls_file.to_csv(file + '.csv', encoding='utf-8')
        shutil.move(filename, archDir + '\\' + filename)
        logging.info(' Transported ' + filename +
                     ' to archive')