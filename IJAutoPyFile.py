import os, time, pyperclip, logging

tDir = 'S:/Transaction Import'
pDir = 'S:/Transaction Import/Inbound Folder'
mydir1 = os.path.join(tDir)
mydir2 = os.path.join(pDir)

os.chdir(mydir1)
logging.basicConfig(filename='Logfile.log', level=logging.INFO, format='%(asctime)s:%(message)s')

# If files with an extension of .dta exist in folder, clean up data in files
for filename in os.listdir():
    if filename[-4:] == ".dta":

        print(time.strftime("[%m/%d/%Y][%H:%M]") + " Begin IJ import script")

        newList = []
        firstList = []

        try:
            with open(filename, "r") as file:
                data = file.readlines()
        except IOError as e:
            raise e
        for i in data:
            firstList.append(i.split('\n'))

        for i in firstList:
            newList.append(i[0].split(','))

        # Rearrange column data to proper location so it can be accepted into Dynamics SL
        for i in newList:
            print(i)
            if i[0] == "Level1":
                i[10] = i[9]
                i[9] = i[8]
                i[8] = i[7]
                i[7] = i[5]
                i[6] = i[4]
                i[5] = i[3]
                i[4] = ''
                i[3] = ''
                del i[-1]
                pyperclip.copy(i[1])

        file.close()

        # Write new data into a fresh data file
        os.chdir(pDir)
        try:
            with open(filename, "w") as file:
                for i in newList:
                    file.write(','.join(i) + '\n')
        except IOError as e:
            raise e


        os.chdir(tDir)
        if filename[-4:] == ".dta":
            logging.info(' Cleaned ' + filename + ' and moved to Inbound Folder')
            os.remove(filename)
        file.close()