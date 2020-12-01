#Author: Ryan Lampe
# This program is development for sending emails via python

import smtplib
import ssl
from email.message import EmailMessage

user_email = 'learningassistants@gmail.com'
user_password = 'TODO'

to = ['glampe99@gmail.com']
subject = 'Test of Python Mailer'
body = 'I sent this with python!'


msg = EmailMessage()
msg['Subject'] = subject
msg['From'] = user_email
msg['To'] = to

text = msg.as_string()


#send
context = ssl.create_default_context()
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context = context) as server:
    server.login(user_email, user_password)
    for addr in to:
        server.sendmail(user_email, addr, text )
