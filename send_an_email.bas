Attribute VB_Name = "send_an_email"
Sub send_an_email(a_path)
    
    Set mail_app = CreateObject("Outlook.Application")
    Set new_mail = mail_app.CreateItem(0)

    With new_mail
    
        .To = "i@rickhehe.com"
'        .CC = ""
'        .BCC = ""
        
        .Subject = "subject"
        
        .Body = "body"

        ' path of attachment
        .Attachments.Add a_path

        '.send
        .Display
    
    End With

End Sub

