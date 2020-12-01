#Author Ryan Lampe
# Python Program to send CSELAP Evaluations to LAs

#Directions for windows
# 1. Install Python: https://www.python.org/downloads/windows/
# 2.  Go to commandline, navigate to the directory where this script is stored
#     Ensure that the LA_Roster is also in the same directory
# 3. Type pip install xlrd to install the xlrd module for reading from excel files
# 4. Run the program!




#Misc:
#
# Apparently, you need to know the file id to download a file from box, not just the filepath... so search API!
# https://github.com/box/box-node-sdk/issues/344
# https://stackoverflow.com/questions/45870835/how-to-obtain-file-id-from-file-name-in-box-com-api


#SETUP:
# 1. Download the Course LA Roster for the current semester from box
#    (Should be named Year Term.xlsx in the Course LA Roster Folder)
# 1b.  Update la_roster filepath below with the path to whereever the
#      file was downloaed to. In windows, if you find the file in file explorer,
#      you can hold shift and right click the file and choose "Copy as path"
#      the updated variable should look something like this:
#      la_roster_file_path = ("C:\\Users\\rjlam\\Documents\\LAP\\Evaluation_Automation\\LA_Roster.xlsx")

# 2. Download the evaluations subfolder to send emails from.
#    (Should be in box named as Evaluation Name eg: Final Evaluation )
#    You will probably need to unzip the folder!
# 2b.  Update the evaluation folder path to the filepath for wherever the
#      evals folder was downloaded. (This folder should then contain subfolders
#      LA's (and/or CLs) )


#Expected folder structure for evals:
# Evaluation Name -> LA Reports, CL Reports (Optional) -> 101, 156, 155E, etc.(opt) -> PDFs -> LastName_Final.pdf



import xlrd

#LA struct
class LA:
    def __init__(self, firstName, lastName, email, course, evalFound, downloadSuccess):
        self.firstName = firstName
        self.lastName = lastName
        self.email = email
        self.evalFound = evalFound
        self.course = course
        self.downloadSuccess = downloadSuccess



#HEY YOU! CHANGE THESE!
#TODO - MOVE TO CONFIG FILE!

la_roster_file_path = ("C:\\Users\\rjlam\\Documents\\LAP\\LA_Roster.xlsx")
evaluation_folder_path = ("C:\\Users\\rjlam\\Documents\\LAP\\Compiled Final Evals\\Compiled Final Evals")



#Open book:
wb = xlrd.open_workbook(la_roster_file_path)
sheet = wb.sheet_by_index(0)

#Map columns
columnNamesDict = dict()
for i in range(sheet.ncols):
    columnName = sheet.cell_value(0,i)
    columnNamesDict[columnName] = i #The column name is the key, and the i is its index.
    #TODO REMOVE PRINT DEBUG
    print(columnName)

#Make sure that the dictonary has the columns that we need:
requiredColumns = ['FirstName', 'LastName', 'Course', 'Position', 'Email']
for c in requiredColumns:
    if c not in columnNamesDict:
        print("Unable to find column: %s. Quitting." % c)
        exit()



print("Looking for mathing names and evaluation files")

#build a list of las = []
las = list()

for i in range(1, sheet.nrows): #skip the first row as it should be column headers
    fname = sheet.cell_value(i, 0)
    lname = sheet.cell_value(i, 1)
    email = sheet.cell_value(i, 2)
    course = sheet.cell_value(i, 3)
    la = LA(fname, lname, email, course, False, False)
    las.append(la)

for p in las:
    prettyName = "%-10s, %-10s --- %15s" % (p.firstName, p.lastName, p.email)
    print(prettyName)
#    print(p.firstName + ", " + p.lastName + " --- " + p.email)
