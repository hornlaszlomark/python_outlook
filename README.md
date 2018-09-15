Microsoft Outlook manipulation with Python

Let's start right away:

import win32com.client
outlook = win32.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")

There are 2 ways from here
1.) 
root_folder = outlook.GetDefaultFolder(6)

OR
2.)
root_folder = outlook.Folders(3).Items
messages = root_folder.Folders.Item(1)


Looping through folders:
for folder in root_folder.Folders:
  print(folder.Name)
  
 Reading messages:
 
 messages = root_folder.Items
 
 
 ... to be continued ... 
 #TODO:
 - accessing parts of messages (Sender.Name, SenderEmailAddress, Body, etc.)  
 - loop through messages
 - loop thorugh messages based on certain conditions
    -   
 - loop through messages and download attachments
 - loop through messsages and download certain attachments
 - writing and sending email
