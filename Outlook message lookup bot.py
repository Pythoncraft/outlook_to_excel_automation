import win32com.client as client
import re

outlook = client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

inbox = namespace.GetDefaultFolder(6) # "6" refers to the inbox folder. returns an object

leads_folder = inbox.Folders['Leads'] # refering / accessing the 'Leads' subfolder
# print(leads_folder.Name)
message = leads_folder.items[0]
print(message.SenderName) 
print(message.CreationTime) 

