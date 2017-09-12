import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox_folder = outlook.GetDefaultFolder(6)  # "6" refers to the index of a folder - in this case,
# the inbox.
sent_folder = outlook.GetDefaultFolder(6)  # "5" refers to sent items.
cal_folder = outlook.GetDefaultFolder(12)

inbox = inbox_folder.Items
sent = sent_folder.Items
cal = cal_folder.Items
