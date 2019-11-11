headers = ['BatNbr', 'Company', 'ChkCompany', 'GLAcct', 'GLSubAcct', 'RefNbr', 'TranDate', 'GLAmount', 'Desc',
           'IsCredit', 'Site', 'Fund', 'VendorID', 'VendorName', 'BillCharges', 'InvoiceDate', 'DueDate',
           'ServiceBeginDate', 'ServiceEndDate', 'DaysOfService', 'VendorAddr', 'VendorAccount', 'BillMonth',
           'FileTotalAmt']

# Load up XML parser and other necessary modules
import xml.etree.ElementTree as ET
import datetime, os, csv, shutil, logging

# Current month and year will name the folder
today = datetime.date.today()
folderString = today.strftime('%m%Y')

# Setting up the accounting file directories
zdrive = '\\\\otofileserv\\accounting'
# caFolder = 'Cash Accrual'
aiq2_folder = 'AIQ2'
exFolder = 'Excel Files'
archiveFolder = 'Archive'

excelDir = os.path.join(zdrive, os.sep, aiq2_folder, folderString, exFolder)
archDir = os.path.join(zdrive, os.sep, aiq2_folder, folderString, archiveFolder)

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
        logging.info('Generating csv file')
        # Create blank CSV file named after xls file
        IQ_data = open(excelDir + '\\' + file + '.csv', 'w', newline='')

        csvwriter = csv.writer(IQ_data)

        # Begin parsing xls file as xml file
        tree = ET.parse(excelDir + '\\' + file + file_ext)
        root = tree.getroot()

        # xml encoding assigned to variables
        worksheet = '{urn:schemas-microsoft-com:office:spreadsheet}Worksheet'
        table = '{urn:schemas-microsoft-com:office:spreadsheet}Table'
        row = '{urn:schemas-microsoft-com:office:spreadsheet}Row'
        cell = '{urn:schemas-microsoft-com:office:spreadsheet}Cell'
        data = '{urn:schemas-microsoft-com:office:spreadsheet}Data'

        rowList = root.findall(worksheet + '/' + table + '/' + row)

        # Lists to hold xml field values (masterList holds itemList)
        masterList = []
        itemList = []

        csvwriter.writerow(headers)

#       Seeks out data to load into csv and assigns it by row to itemList
        count = 0
        while count != len(rowList):
            for row_item in rowList[count]:
                for cell_item in row_item:
                    if cell_item.tag == data:
                        if cell_item.text in headers:
                            continue
                        else:
                            itemList.append(cell_item.text)

            # Put itemList row data into masterList, reset itemList for next row
            masterList.append(itemList)
            count += 1
            itemList = []

        # Write each row in masterList to the csv file and close the file
        for item in masterList[1:]:
            item[0] = ''

        for item in masterList[1:len(masterList)]:
            csvwriter.writerow(item)

        IQ_data.close()

# If file is dta, move to inbound folder. move csv to archive
for filename in os.listdir():
    file, file_ext = os.path.splitext(filename)
    if file_ext == '.xls':
        shutil.move(filename, archDir)
        logging.info(' Transported ' + filename + ' to archive')
