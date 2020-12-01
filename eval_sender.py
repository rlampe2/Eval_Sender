#Author Ryan Lampe
# Python Program to send CSELAP Evaluations to LAs

#Directions for windows
# 1. Install Python: https://www.python.org/downloads/windows/
# 2.  Go to commandline, navigate to the directory where this script is stored
#     Ensure that the LA_Roster is also in the same directory
# 3. Type pip install xlrd to install the xlrd module for reading from excel files
# 4. Run the program!



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

#Expected File naming convention: LastName_ShortEvalname.pdf.  Eg: Lampe_Mid.pdf



import xlrd
import os.path
import re
import eval_config

#LA struct
class LA:
    def __init__(self, firstName, lastName, email, course, position, eval_path, downloadSuccess):
        self.firstName = firstName
        self.lastName = lastName
        self.email = email
        self.course = course
        self.position = position
        self.eval_path = eval_path
        self.downloadSuccess = downloadSuccess #Unimplemented until Box API approved


la_roster_file_path = eval_config.la_roster_file_path
evaluation_folder_path = eval_config.evaluation_folder_path


#Open book:
if not os.path.isfile(la_roster_file_path):
    print("Couldn't find LA Roster File. Exiting...")
    exit()
wb = xlrd.open_workbook(la_roster_file_path)
sheet = wb.sheet_by_index(0)

#Map columns
columnNamesDict = dict()
for i in range(sheet.ncols):
    columnName = sheet.cell_value(0,i)
    columnNamesDict[columnName] = i #The column name is the key, and the i is its index.

#Make sure that the dictonary has the columns that we need:
requiredColumns = ['FirstName', 'LastName', 'Course', 'Position', 'Email']
for c in requiredColumns:
    if c not in columnNamesDict:
        print("Unable to find column: %s. Quitting." % c)
        exit()

#Ask which courses to do evals for:
courses = list()
n = int(input("Enter the number of courses this eval is for:"))
for i in range(0, n):
    course = input("Enter course number: ")
    #Remove CSCE- if included
    course = re.sub("^[^1]*1", "1", course ) #Replace all text leading up to 1.
    course = course.upper()
    if course == '156':
        h = input("Include 156H? (Yes/No) ")
        if h[0] == 'y' or h[0] == 'Y':
            courses.append('156H')
    courses.append(course)

positions = list()

p = input("Will some of these evals be sent to LAs? (Yes/No) ")
if p[0] == 'y' or p[0] == 'Y':
    positions.append('LA')

p = input("Will some of these evals be sent to CLs? (Yes/No) ")
if p[0] == 'y' or p[0] == 'Y':
    positions.append('CL')



print("\nWe need the last part of the file name for each evaluation.\n  Eg if the filenames are like: Lampe_Final.pdf, Hahn_Final.pfd then enter \"Final\"\n (without quotes, no extension!)")
short_eval_name = input("Please enter the last part of the eval name: ")


print("Looking for mathing names and evaluation files for\n  Position(s): %s \n  Course(s):   %s\n\n " % (str(positions), str(courses) ) )

#build a list of las = []
las = list()

for i in range(1, sheet.nrows): #skip the first row as it should be column headers
    fname = sheet.cell_value(i, columnNamesDict['FirstName'])
    lname = sheet.cell_value(i, columnNamesDict['LastName'])
    email = sheet.cell_value(i, columnNamesDict['Email'])
    course = str(sheet.cell_value(i, columnNamesDict['Course'])).upper()
    #Since reading from excel can have 101.0 instead of 101, but can't round
    #since 155N isn't an int:
    if '.' in course:
        course = str(int(float(course)))
    position = sheet.cell_value(i, columnNamesDict['Position'])
    la = LA(fname, lname, email, course, position, False, False)
    if la.course in courses and la.position in positions:
        las.append(la)


#Find the files for each LA

file = list()
no_file = list()
num_without_file = 0
num_with_file = 0

for p in las:
    indv_eval_file = f"{evaluation_folder_path}\\{p.position} Reports\\{p.course}\\PDFs\\{p.lastName}_{short_eval_name}.pdf"
    print(indv_eval_file)
    p.eval_path = indv_eval_file
    if not os.path.isfile(indv_eval_file):
        num_without_file += 1
        no_file.append(p)
    else:
        file.append(p)
        num_with_file += 1

print("----------------------------\nSearching Complete. \nResults: \nEvals found: %s. \nEvals not found: %s" % (num_with_file, num_without_file))
print("LAs/CLs with eval files: ")
for p in file:
    prettyName = "%-10s, %-10s --- %15s   %5s   %5s" % (p.firstName, p.lastName, p.email, p.position, p.course)
    print(prettyName)

if num_without_file != 0:
    print("LAs/CLs without evals found:-----------------------")
    for p in no_file:
       prettyName = "%-10s, %-10s --- %15s   %5s   %5s" % (p.firstName, p.lastName, p.email, p.position, p.course)
       print(prettyName)
cont = input("Continue to email LAs/CLs with found files?")
if not(cont[0] == 'y' or cont[0] == 'Y'):
    print("Exiting...")
    exit()

#load up all the emails:
print("------------------------\nLoading Emails!")
first = True
from string import Template
import smtplib
import ssl
from email.message import EmailMessage


user_email = eval_config.user_email
user_password = eval_config.user_email_password



for p in file:
    body = Template(eval_config.email_body).substitute(first_name = p.firstName, last_name = p.lastName, course = p.course, position = p.position)

    msg = EmailMessage()
    msg['Subject'] = Template(eval_config.email_subject).substitute(first_name = p.firstName, last_name = p.lastName, course = p.course, position = p.position)
    msg['From'] = user_email
    msg['To'] = p.email
    msg.set_content(body)
    #Print this first before attaching PDF - otherwise it'll flood stdout
    if first:
        print("Here is what an example looks like:")
        print(str(msg))
        first = False
        cont = input("Continue? (Yes/No) ")
        if not(cont[0] == 'y' or cont[0] == 'Y'):
            print("Exiting...")
            exit()

    with open(p.eval_path, "rb") as content_file:
        content = content_file.read()
        msg.add_attachment(content, maintype='application', subtype='pdf', filename= f"{p.lastName}_{short_eval_name}.pdf")

    print("Creating SSL Context")
    context = ssl.create_default_context()

    print("Logging In")
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context = context) as server:
        server.login(user_email, user_password)
        server.send_message(msg)
