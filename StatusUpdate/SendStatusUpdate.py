'''

This script does the below things

- Find the latest mail of sent item or inbox item based on the given subject
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
    wb = openpyxl.load_workbook(config.filename)   
    
    for sheet in wb.worksheets:
        if config.sheetname not in sheet.title:
            wb.remove(sheet)
        
    wb.save(new_filename)

def GetMessageFromFolder(folder):
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
                                    
                                    
found_sent_message, sentmessage = GetMessageFromFolder(5)
found_inbox_message, inboxmessage = GetMessageFromFolder(6)

  
if not found_sent_message  and not found_inbox_message:
    logging.debug("Mail not found! returning!!!")
    sys.exit()

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

    
message = message.Reply() 

data = pd.read_excel(config.filename,sheet_name = config.sheetname)
data = data.fillna("")    
message.HTMLBody =  config.content  + '<br />  <br />'+  data.to_html()   + message.HTMLBody


ExtractSheetAndCreateNewFile(config.file_to_attach)

message.Attachments.Add(config.file_to_attach)

message.To = config.to
message.Cc = config.cc  

message.Send()  
        
