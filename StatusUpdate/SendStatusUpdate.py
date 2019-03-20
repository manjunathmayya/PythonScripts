'''

This script does the below things

- Find the latest mail from sent items or inbox items based on the given subject
- Draft a reply email for the selected mail item.
- Extract excel sheet contents from the provided excel file and sheet name to be read.
- Attach the excel sheet contents to the draft email.
- Create a new file containing only the required excel sheet
- Attach the new file created to the email.
- Send this draft email to addresses mentioned in to and cc list.

'''

import win32com.client,sys
import pandas as pd
import logging
import config
import openpyxl

logging.basicConfig(format='%(message)s', level=logging.DEBUG)
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")                                   
logging.info("Subject: " + config.subject_to_find)


def ExtractSheetAndCreateNewFile(new_filename):
    '''
    Open the workbook, keep only the required sheet and remove others.
    Save the workbook as a new file wth new_filename.
    '''
    wb = openpyxl.load_workbook(config.filename)   
    
    for sheet in wb.worksheets:
        if config.sheetname not in sheet.title:
            wb.remove(sheet)
        
    wb.save(new_filename)

def GetMessageFromFolder(folder):
    '''
    Search email with required subject, in specified folder (passed as parameter)
    Return a tuple with values
    - whether email is found or not
    - found email object
    '''
    items = outlook.GetDefaultFolder(folder)                                    
    messages = items.Items
    
    found_message = False    
    message = None
    for message in reversed(messages):
        if config.subject_to_find in message.Subject:
            logging.debug ("Found in " + ("sent" if folder == 5 else "inbox") + " items")
            found_message = True
            break 
    
    return found_message, message
                                    
              
#Get latest emails from sent items (folder number 5) and inbox (folder number 6)                     
found_sent_message, sentmessage = GetMessageFromFolder(5)
found_inbox_message, inboxmessage = GetMessageFromFolder(6)

  
#Return if email is not found in both sent items and inbox
if not found_sent_message  and not found_inbox_message:
    logging.debug("Mail not found! returning!!!")
    sys.exit()

#Select latest email message of the two (sent item or inbox)
message = None     
if found_sent_message and found_inbox_message:
    logging.debug('Inbox message received time : ' + str(inboxmessage.ReceivedTime))
    logging.debug('Sent item sent time         : ' + str(sentmessage.SentOn))
    if inboxmessage.ReceivedTime > sentmessage.SentOn:
        message = inboxmessage
        logging.info("inbox message selected \n")
    else:
        message = sentmessage
        logging.info("sent message selected \n")
elif found_sent_message:
    message = sentmessage
    logging.info("sent message selected \n")
else:
    message = inboxmessage 
    logging.info("inbox message selected \n")

    
#Create a reply message with data from sheetname mentioned in configuration    
message = message.Reply() 

data = pd.read_excel(config.filename,sheet_name = config.sheetname)
data = data.fillna("")    
message.HTMLBody =  config.content  + '<br />  <br />'+  data.to_html()   + message.HTMLBody

#Create a new file with only required sheet from excel file.
ExtractSheetAndCreateNewFile(config.file_to_attach)

#Attach the newly created file to the email and send email to required addresses
message.Attachments.Add(config.file_to_attach)
message.To = config.to
message.Cc = config.cc  
message.Send()          
