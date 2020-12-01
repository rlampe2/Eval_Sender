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

#TODO Connect to Box for files

#TODO Get names and emails from excel sheet in box


#Open file:
file_path = ("C:\\Users\\rjlam\\Documents\\LAP\\Evaluation_Automation\\LA_Roster.xlsx")
#Open book:
wb = xlrd.open_workbook(file_path)
sheet = wb.sheet_by_index(0)

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
