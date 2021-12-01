import os
import win32com.client as win32
from win32com.client import constants
import re
import sys
import time 

"Finds SEM data files in the form of word documents, updates them"

projectNumber = input("Insert your project number here: ")
projectNumber = projectNumber.strip(' ')
analysisTablePath = ""
inital_path = r'C:\Users\jcornelius\Desktop\\'

def edsDataFinder(projectNumber, start_path):

    print("Running the EDS Data Finder")
    print()
          
    with os.scandir(start_path) as entries:
        projectPath = " "
        for entry in entries:
            if entry.name.startswith(projectNumber) and os.path.isdir(entry) == True:
                projectPath = start_path + entry.name
                sys.stdout.write("The project path is " + projectPath)
                print()

    if projectNumber not in projectPath:
        print("Check you have entered the project number correctly. You entered: " + projectNumber)
        print("Check the correct folder is on your desktop")
        print("Try these changes and then run again")
        print()

        enter = 0
        while enter != "":
            enter = input("Hit enter to close this window, "
                          "then run the program again with the changes made")
        else:
            exit()


    with os.scandir(projectPath) as folders:
        semPath = " "
        for folder in folders:
            if folder.name.startswith("SEM") and folder.name.endswith("SEM") and os.path.isdir(folder) == True:
                semPath = projectPath + "\\" + folder.name
                print("I found the SEM folder")
                print()

    if semPath == " ":
        print("Make sure there is an SEM folder in the project folder titled: SEM")
        print("Try these changes and then run again")

        enter = 0
        while enter != "":
            enter = input("Hit enter to close this window, "
                          "then run the program again with the changes made")
        else:
            exit()

    with os.scandir(semPath) as sheets:
        edsDataPath = " "
        edsFilePaths = []
        for sheet in sheets:
            if sheet.name.startswith('Analysis Table1') and sheet.name.endswith('.xlsm'):
                edsDataPath = semPath + "\\" + sheet.name
                edsFilePaths.append(edsDataPath)

    if edsDataPath == " ":
        print("Check there is an excel document called: Analysis Table1")
        print("Check the Analysis Table1 file type is an .xlsm file")
        print("Try these changes and then run again")

        enter = 0
        while enter != "":
            enter = input("Hit enter to close this window, "
                          "then run the program again with the changes made")
        else:
            exit()

    with os.scandir(semPath) as files:
        for file in files:
            if file.name.endswith('.doc') and file.name.startswith(projectNumber):
                edsDataPath = semPath + "\\" + file.name
                edsFilePaths.append(edsDataPath)

    if edsFilePaths == []:
        print("Make sure the word documents are saved with " + projectNumber + " at the start")
        print("Make sure the word documents are .doc files")
        print("Try these changes and then run again")

        enter = 0
        while enter != "":
            enter = input("Hit enter to close this window, "
                          "then run the program again with the changes made")
        else:
            exit()
    else:
        return edsFilePaths


edsFilePathss = edsDataFinder(projectNumber, inital_path)


def docConverter(submittalPath):
    
    print("Converting the document")
    # Opening MS Word

    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(submittalPath)
    time.sleep(0.2)

    # Rename path with .docx
    new_file_abs = os.path.abspath(submittalPath)
    new_file_abs = re.sub('\.doc', '.docx', new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(True)
    print("Done converting the document")
    print()


#Splits off the excel file at the beginning of the array,
# so the .doc files are the only ones processed.
splitEDSFilePaths = edsFilePathss[1:]


for i in range(len(splitEDSFilePaths)):
    docConverter(splitEDSFilePaths[i])
    splitEDSFilePaths[i] = splitEDSFilePaths[i] + 'x'
    

def TableParser2(edsFilePath, counter):

    path = edsFilePath

    print("I'm parsing the document for the EDS data table..." + path)

    step1 = path.split("SEM\\" + projectNumber, 1)
    dataLabel = step1[1].split(".docx", 1)
    dataLabel = dataLabel[0]
    #print("Data Label: " + dataLabel)
    #print(step1)
    step2 = step1[0]
    #print(step2)
    step3 = step2 + "\SEM"
    #print(step3)

    output_path = step3
    #opening excel file

    XL = win32.Dispatch('Excel.Application')
    XL.Visible = True
    
    XLBook = XL.Workbooks.Open(os.path.join(output_path, 'Analysis Table1.xlsm'))
    XLSheet = XLBook.Worksheets(1)
    #counter to track excel row
    

    #opening word file
    word = win32.Dispatch('Word.Application')
    word.Visible = True
    
    word.Documents.Open(path)
    doc = word.ActiveDocument

    table = doc.Tables(1)

    for x in range(1, (table.Columns.Count - 1)):
        
        for y in range(1, (table.Rows.Count)):
            # print(counter)
            content = table.Cell(Row=y, Column=x).Range.Text
            # print("The row content is: " + content)
            # print()
            excelRowPosition = y + 1 
            excelColumnPosition = x + (2 * counter)
            
            XLSheet.Cells(excelRowPosition, excelColumnPosition).Value = table.Cell(Row=y, Column=x).Range.Text

            if XLSheet.Cells(y, excelColumnPosition).Value is None:
                # print("There was nothing")
                XLSheet.Cells(y, excelColumnPosition).Value = dataLabel
            else:
                print(XLSheet.Cells(y + (2 * counter), excelColumnPosition).Value)
                
    doc.Save()
    doc.Close()
    # word.Quit()
    del word

    XLBook.Save()
    # XLBook.Close()
    # XL.Quit()
    del XL


counter = 0
for i in range(len(splitEDSFilePaths)):

    TableParser2(splitEDSFilePaths[i], counter)
    counter = counter + 1
