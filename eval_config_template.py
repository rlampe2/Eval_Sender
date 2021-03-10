#Author: Ryan Lampe
# rlampe@huskers.unl.edu
# Configuration File for LAP Eval Sender

# This is the template configuration file for the eval sender
# First, make a COPY of this file in the same directory and name it as evaL_config.py
# Then, make the following changes in the eval_config.py file

# File Paths:
#   Update these to include the paths to the LA Roster spreadsheet
#   and the folder holding the evaluations (This folder should have one or two
#   subfolders LA Reports and CL Reports if setup to match the program.)
#   eg: la_roster_file_path = "C:\\Users\\rjlam\\Documents\\LAP\\LA_Roster.xlsx"
#   Note: The name of the roster can be different than LA_Roster.xlsx, just make sure that is the name given here.


la_roster_file_path = ""
evaluation_folder_path = ""

# Email Settings:
#   Add the user email address and password e.g. user_email = 'cselap@gmail.com'
#   Also, ensure that the email is configured (login on browser) to allow less
#   secure app access is "ON". (Google auto shuts this off. It may reject you, but they will send you
#   an email which you can turn it back on

# This is the cse...@gmail email and @gmail password
user_email = ''
user_email_password = ''


# Email Message:

# When writing the email subject and body use ${first_name}, ${last_name}
# ${course} and ${position} if you want these values to be tailored to each person
# that the email is sent to.
# eg: """Hello ${firstname},
#        How are you?"""   in the email_body in this file would become
#     """Hello Ryan,
#       How are you?"""   when the actual email is sent to someone named Ryan.
# The triple equals preserve newlines and you will see a preview of your message
# initial message before you send them all out!
# Examples:
#email_subject = "Final Evaluation for ${course} ${position}s"
# email_body = """Hello ${first_name},
# We are pleased to provide your LAP eval.
#
# Thanks,
# LAP TEAM.
#  """

#  A recently used message:
#mail_subject = "CSELAP - Midsemester Evaulations"
#email_body = """Hello ${first_name},
#Midsemester evaluations (round one) are completed. Please find your evaluation attached to this email. If you have questions, please reach out!
#
#
#Thanks,
#LAP TEAM. """
email_subject = ""
email_body = """ """
