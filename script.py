import win32com.client
import os
import datetime as dt
import rarfile



outputDir = "C:/Users/pc-asus/AppData/Local/Programs/Python/Python37/Read-email-with-python-V2/data/"



def restrictParams(senderEmail):
    # Config

    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")

    # Select main Inbox
    inbox = mapi.GetDefaultFolder(6) 
    messages = inbox.Items

    DateTime = dt.datetime.now() #- dt.timedelta(hours=24)
    DateTime = DateTime.strftime('%m/%d/%Y %H:%M %p')  #<--  format compatible avec "Restrict"

	# Only search emails in the day:
    messages = messages.Restrict("[ReceivedTime] >= '" + DateTime +"'")
    #messages = messages.Restrict("[Subject] = 'Log File'")
	# Configuration du mail du destinataire
    messages = messages.Restrict("[SenderEmailAddress] = '" + senderEmail +"'")

    return messages 



def extractEmailPJ(senderEmail, output_path= outputDir):
    message = restrictParams(senderEmail)
    DateTime = dt.datetime.now()
    try:
        for msg in list(message):
                try:
                    s = msg.sender
                    for attachment in msg.Attachments:
                        attachment.SaveASFile(os.path.join(output_path, attachment.FileName))

                        with open('./logFile.txt','a') as f:
                            f.write(f"\n==== date : {DateTime} ====\n attachment {attachment.FileName} from {s} saved")
                            print("=========== succes ==============")
                except Exception as e:
                    with open('./logFile.txt','a') as f:
                            f.write(f"\n==== date : {DateTime} ====\nerror when saving the attachment:" + str(e))
                            print("error when saving the attachment:" + str(e))
             
    except Exception as e:
        with open('./logFile.txt','a') as f:
                        f.write(f"\n==== date : {DateTime} ====\nerror when processing emails messages:" + str(e))
                        print("error when processing emails messages:" + str(e))
                
                
def extractFile(input_path:str):
    
    elements = os.listdir(path=input_path)
    for file in elements:
        if file.endswith('.rar'):
            file_path = os.path.join(input_path,file)
            rar = rarfile.RarFile(file_path)
            rar.extractall(path=input_path)
            del file
            del file_path
            del rar

extractEmailPJ('hermam225@gmail.com')
extractFile(input_path=outputDir)
