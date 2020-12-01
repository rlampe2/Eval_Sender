#Author Ryan Lampe
# Python Program to send CSELAP Evaluations to LAs

#Directions for windows
# 1. Install Python: https://www.python.org/downloads/windows/
# 2.  Go to commandline, navigate to the directory where this script is stored
#     Ensure that the LA_Roster is also in the same directory
# 3. Type pip install xlrd to install the xlrd module for reading from excel files
# 4. Run the program!


import xlrd

#TODO Connect to Box for files

#TODO Get names and emails from excel sheet in box


#Open file:
file_path = ("C:\\Users\\rjlam\\Documents\\LAP\\LA_Roster.xlsx")
#Open book:
wb = xlrd.open_workbook(file_path)
sheet = wb.sheet_by_index(0)

print("Looking for mathing names and evaluation files")

#build a dictionary of names that map to a list (contains firstname, lastname, email, course, download successful etc)
las = {}

for i in range(1, sheet.nrows): #skip the first row as it should be column headers
    las[sheet.cell_value(i,0) + "_" + sheet.cell_value(i,1)] = [sheet.cell_value(i,2)]

print(las)
