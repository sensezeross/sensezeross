import win32com.client as win32

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

mailItem = olApp.CreateItem(0)
mailItem.Subject = 'Hello 123'
mailItem.BodyFormat = 1
mailItem.Body = "Hello There"
mailItem.To = 'kriskorn_ing@truecorp.co.th'
mailItem._oleobj_.Invole(*(64209, 0, 8, 0, olNS.Accounts.Item('kriskorn_ing@truecorp.co.th')))

mailItem.Display()
mailItem.Save()
