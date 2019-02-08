''' Usage 
python sendEmailUsingOutlook.py -e xxx@gmail.com -s "Subject of email" -m "Email Message"
'''

import win32com.client as win32
import argparse

def SendEmail(emailAddress, subject, message):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = emailAddress
    mail.Subject = subject
    mail.Body = message
    mail.HTMLBody = '<h2>' + message +'</h2>' #this field is optional
    # To attach a file to the email (optional):
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add(attachment)    
    mail.Send()


#Get arguments from user.
parser = argparse.ArgumentParser()
parser.add_argument(
    '--emailAddress',
    '-e',
    default='manjunath.k@siemens-healthineers.com',
    help='Enter email address of the recepient')
parser.add_argument(
    '--subject',
    '-s',
    default='Subject',
    help='Enter subject for the email')
parser.add_argument(
    '--message',
    '-m',    
    default='message',
    help='Enter message for the email')

args = parser.parse_args()

SendEmail(args.emailAddress, args.subject, args.message)


