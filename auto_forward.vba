Private WithEvents Items As Outlook.Items
Private Sub Application_Startup()
    Dim outlookApp As Outlook.Application
    Dim outlookNamespace As Outlook.NameSpace
    Set outlookApp = Outlook.Application
    Set outlookNamespace = outlookApp.GetNamespace("MAPI")
    Set Items = outlookNamespace.Folders("ParkingandTransportationScheduling  (NBCUniversal, Orlando)").Folders("Inbox").Items
    Call NotifyForward
End Sub
Private Sub Items_ItemAdd(ByVal item As Object)
    'all them variables
    Dim app As Outlook.Application
    Dim my_space As NameSpace
    Dim inbox As MAPIFolder
    Dim filtered_items As Items
    Dim filter As String
    Dim mail As MailItem
    Dim mail2 As MailItem
    Dim regex As Object
    Dim regex2 As Object
    Dim to_addrs As String
   
    'just some setup so we can get this baby on the road
    Set app = New Outlook.Application
    Set my_space = GetNamespace("MAPI")
    'selecting, explicitly, the mailbox to perform actions on, as well as the folder inside of it to work on
    Set inbox = my_space.Folders("ParkingandTransportationScheduling  (NBCUniversal, Orlando)").Folders("Inbox")
   
    'some sql to search the mailbox specified above for emails fitting out criteria
    filter = "@SQL= urn:schemas:httpmail:subject LIKE '%[EXTERNAL] Your shift swap request has been%'"
    Set filtered_items = inbox.Items.Restrict(filter)
   
    'regular expression object to search the email body for TM email addresses
    'searches specifically for TM emails by looking for emails that are preceded by TM#1 or 2 Email
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .pattern = "(TM\#\d\sEmail(\s)*)([a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*)"
        .Global = True
        .IgnoreCase = True
    End With
   
    'regular expression object to search the first regex object for just the email
    Set regex2 = CreateObject("VBScript.RegExp")
    With regex2
        .pattern = "[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*"
        .Global = True
        .IgnoreCase = True
    End With
   
    'if we find no emails with a subject line meeting our criteria then clear out our variables
    If filtered_items.Count = 0 Then
        Set inbox = Nothing
        Set my_space = Nothing
    Else
        'looping through each mail item that was found meeting our criteria
        For Each item In filtered_items
            If TypeName(item) = "MailItem" Then
            Set mail = item
                'check if the item has been read, if so, skip it, if not then forward it to the tms and mark it as read
                If mail.UnRead = True Then
                    to_addrs = ""
                    Set matches = regex.Execute(mail.Body)
                    For Each Match In matches
                        Set matches2 = regex2.Execute(Match)
                        For Each Match2 In matches2
                            to_addrs = to_addrs + Match2 & "; "
                        Next Match2
                    Next Match
                    If to_addrs <> "" Then
                        Set mail2 = mail.Forward
                        mail2.To = to_addrs
                        mail2.Send
                        Set mail2 = Nothing
                    Else
                        'do nothing here
                    End If
                    'marking the mail item as read
                    With mail
                        .UnRead = False
                    End With
                Else
                    'skipping if the item has been read
                End If
            Else
            ' do nothing
            End If
        Next item
    End If
End Sub
