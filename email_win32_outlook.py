import win32com.client as win32

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

mailItem = olApp.CreateItem(0)
mailItem.Subject = 'Hello 123'
mailItem.BodyFormat = 1
mailItem.Body = "Hello There"
mailItem.To = 'mail'
mailItem._oleobj_.Invole(*(64209, 0, 8, 0, olNS.Accounts.Item('mail')))

mailItem.Display()
mailItem.Save()
