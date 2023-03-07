Sub DelayedDelivery()
    Dim objMail As MailItem
    
  ' If you want to modify an existing email to use delayed delivery, you can add the following code after the Set objMail line to select the existing email:
    Set objMail = Application.ActiveInspector.CurrentItem

   
    
    ' Set the email details
    With objMail
        .Subject = "Delayed email"
      'Uncomment this line and edit to your needs .To = "recipient@example.com"
        .Body = "This is a test email."
        
        ' Set the delay delivery time to 6:00 AM
        .DeferredDeliveryTime = Date + TimeValue("06:00:00")
        
        ' Send the email
        .Send
    End With
    
    ' Release memory
    Set objMail = Nothing
End Sub
