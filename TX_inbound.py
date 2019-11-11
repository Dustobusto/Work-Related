import os, datetime

today = datetime.date.today()
folderString = today.strftime('%m%Y')
folderString2 = today.strftime('%Y%m')
folderString3 = today.strftime('%Y%m%d')

sdrive = 'S:'
transactionFolder = 'Transaction Import'
inboundFolder = 'Inbound Folder'

inboundDir = os.path.join(sdrive, os.sep, transactionFolder, inboundFolder)
os.chdir(inboundDir)

# If file is AnyBill, read the data, split the extension and name, create new file with new data, and remove old file
for i in os.listdir():
    if i[0:7].upper() == 'ANYBILL':

        with open(i, 'r') as file:
            data = file.readlines()

        file, file_ext = os.path.splitext(i)
        x = file.split('_')
        file = "IJ-" + x[1] + '-' + folderString3 + "_TX"
        file_ext = '.dta'
        filename = file + file_ext
        with open(filename, 'w') as f:
            for line in data:
                if 'Level0' in line:
                    f.write(line)
                elif 'Level1' in line:
                    z = line.split(',')
                    x = ','.join(z[0:11])
                    f.write(x + '\n')
        f.close()
        os.remove(i)
