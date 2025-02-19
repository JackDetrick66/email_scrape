#test
import win32com.client
import os
#connect to the local outlook
myOutlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")

#Told index 6 is usually the default inbox, plan to just use this
myInbox = myOutlook.GetDefaultFolder(6)

messages = myInbox.Items
# Sort by newest first
messages.Sort("[ReceivedTime]", True)

#test fetch 
for i, message in enumerate(messages[:5]):
    print(f"Subject: {message.Subject}")
    print(f"From: {message.SenderName} <{message.SenderEmailAddress}>")
    print(f"Received: {message.ReceivedTime}")
    print(f"Body:\n{message.Body[:500]}...")  # Print first 500 chars of the body
    print("-" * 50)