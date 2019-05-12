

# Microsoft Outlook manipulation with win32com (Python)

This will always be the first thing to do when interacting with your Outlook on your computer:

```python
import win32com.client
outlook = win32.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
```
You can decide which path to take to continue.
1.) The standard path where you choose the default folder.
```python
root_folder = outlook.GetDefaultFolder(6)
```

<b>OR</b> the custom path where you decide and deliberately choose which email address to use.
2.)
```python
root_folder = outlook.Folders(1).Items
messages = root_folder.Folders.Item(1) # this should be the incoming messages folder
```
You can also you the exact name of the folder. For example:
3.)
```python
root_folder = outlook.Folders('Inbox').Items
```

Let's take number 2.) because it gives us more control over our script/program. 

#TODO:

Looping through folders:
```python
for folder in root_folder.Folders:
  print(folder.Name)
 ``` 
 
 FIND MORE DETAILED INFO IN THE WIKI: https://github.com/hornlaszlomark/python_outlook/wiki
 
 ... to be continued ... 
 #TODO:
 - accessing parts of messages (Sender.Name, SenderEmailAddress, Body, etc.)
 - loop through messages
 - loop thorugh messages based on certain conditions
 - loop through messages and download attachments
 - loop through messsages and download certain attachments
 - writing and sending email

All of the above is coming soon!
