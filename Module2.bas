Attribute VB_Name = "Module2"
Sub SendShiftEmail()
    Dim obApp As Object
    Dim NewMail As MailItem
    Dim MyDate
    Dim MyTime
    Dim signature As String
    
    MyDate = Date
    MyTime = Time

    Set obApp = Outlook.Application
    Set NewMail = obApp.CreateItem(olMailItem)
    
    With NewMail
        .Display
    End With
    
    signature = NewMail.Body
    
    If MyTime < TimeValue("13:00:00 am") Then
        With NewMail
            .Subject = "Shift Start " & (Date)
        .To = "exemplaryEmail@exemplary.com"
            .Body = "Dear Team," & vbCrLf & vbCrLf & "As of " & (MyDate) & " the shift has started at: " & MyTime & vbCrLf & vbCrLf & "Kind regards" & vbCrLf & signature
            .Display
        End With
    ElseIf MyTime > TimeValue("16:00:00 am") Then
        With NewMail
            .Subject = "Shift End " & (Date)
            .To = "awsho.pl@capgemini.com"
            .Body = "Dear Team," & vbCrLf & vbCrLf & "As of " & (MyDate) & " the shift has ënded at: " & MyTime & " am" & vbCrLf & "Kind regards" & vbCrLf & signature
            .Display
        End With
        
    End If
    Set obApp = Nothing
    Set NewMail = Nothing
End Sub
