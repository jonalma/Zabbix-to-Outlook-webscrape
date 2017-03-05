Sub Zalert()
'Note that this example opens and displays a new email message in Outlook, enters subject, body and attaches a file, but does not send a mail.
Dim outlook As outlook.Application
Dim email As outlook.MailItem

'Shell ("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe -url https://nms.j2noc.com/zab/tr_status.php")
Set outlook = New outlook.Application
'create and display a new email item:
Set email = outlook.CreateItem(olMailItem)

'Go through IE`
Set objShell = CreateObject("Shell.Application")
IE_count = objShell.Windows.Count
For X = 0 To (IE_count - 1)
    On Error Resume Next    ' sometimes more web pages are counted than are open
    my_url = objShell.Windows(X).Document.Location
    my_title = objShell.Windows(X).Document.Title
    
    'Check if Zabbix tab is open
    'If my_title Like "*Zabbix*"
    If InStr(my_title, "Zabbix") Then
        'MsgBox ("zabbix")
        Set ie = objShell.Windows(X)
        Exit For
    Else
        'MsgBox ("not here")
    End If
Next

'set properties of the new mail item:
With email
    .Display
    .CC = ""
End With

'if your signature populates the email body, save the existing HTML Body
Dim signature As String
signature = email.HTMLBody
Dim severity As String
Dim host As String
Dim alert As String
Dim lastChange As String
Dim age As String
Dim counter As Integer
counter = 0

If ie.ReadyState = ReadyState.READYSTATE_COMPLETE Then
    Set triggerTable = ie.Document.getElementsByClassName("tableinfo")(0)
    'Set triggerTable = ie.Document.all.tags("tbody").Item(12)
    'MsgBox (triggerTable.innerHTML)
    'It is pulling the wrong tbody when
    Do While ie.Busy: DoEvents: Loop
    Set triggers = triggerTable.getElementsByClassName("even_row selected")
    Do While ie.Busy: DoEvents: Loop

    For Each trig In triggers
        For Each trigInfo In trig.Children
            If StrComp(trigInfo.innerText, trig.Children(4).innerText) = 0 Then
                lastChange = trigInfo.innerText
                'email.Body = email.Body + Trim("Last change: " + trigInfo.innerText)
            ElseIf StrComp(trigInfo.innerText, trig.Children(1).innerText) = 0 Then
                severity = trigInfo.innerText
            ElseIf StrComp(trigInfo.innerText, trig.Children(5).innerText) = 0 Then
                age = trigInfo.innerText
                'email.Body = email.Body + Trim("Age: " + trigInfo.innerText)
            ElseIf StrComp(trigInfo.innerText, trig.Children(7).innerText) = 0 Then
                host = trigInfo.innerText
                'email.Body = email.Body + Trim("Host: " + trigInfo.innerText)
            ElseIf StrComp(trigInfo.innerText, trig.Children(8).innerText) = 0 Then
                alert = trigInfo.innerText
                'email.Body = email.Body + Trim("Name: " + trigInfo.innerText)
            Else
            End If
        Next
    
        If Not counter = 0 Then
            email.HTMLBody = "<BODY style=font-size:11pt;font-family:Calibri>" & email.HTMLBody & "<strong><u>Trigger Information</u></strong><br>" & severity & " - " & host & "<br><span style='background:yellow;mso-highlight:yellow'>" & alert & "</span><br>" & "Last Change: " & lastChange & "<br>" & "Age: " & age & "<br><br></BODY>"
        Else
           email.HTMLBody = "<BODY style=font-size:11pt;font-family:Calibri>" & "<br>" & "<strong><u>Trigger Information</u></strong><br>" & severity & " - " & host & "<br><span style='background:yellow;mso-highlight:yellow'>" & alert & "</span><br>" & "Last Change: " & lastChange & "<br>" & "Age: " & age & "<br><br></BODY>"
        End If
        
        counter = counter + 1
    Next
End If
'MsgBox (counter)


'Append your signature
email.HTMLBody = email.HTMLBody & signature

'Add the subject
email.Subject = "Z Alert: " + host + " " + alert + " (I-ticketnumber)"


''''''''''''''''''''''''''''''''''''''''''
'Put everthing in a maximo incident ticket
If ie.ReadyState = ReadyState.READYSTATE_COMPLETE Then
    For X = 0 To (IE_count - 1)
        On Error Resume Next    ' sometimes more web pages are counted than are open
        my_url = objShell.Windows(X).Document.Location
        my_title = objShell.Windows(X).Document.Title
    
        'Check if Zabbix tab is open
        'If my_title Like "*Zabbix*" Or my_url Like "*nms*" Then   'identify the existing web page
        If InStr(my_title, "Incidents") Or InStr(my_url, "maximo") Then
            'MsgBox ("maximo")
            Set ie = objShell.Windows(X)
            Exit For
        Else
            'MsgBox ("not here")
        End If
    Next
    
    'Fill in Summary field
    ie.Document.getElementById("mx628").Value = "Z Alert: " + host + " " + alert
    'Fill in Details field
    ie.Document.getElementById("dijitEditorBody").innerText = "hello"
    'MsgBox (ie.Document.getElementById("dijitEditorBody").innerHTML)
    'Fill classification path
    ie.Document.getElementById("mx646").Value = "PRODUCTION \ EFAX \ APPLICATION"
    'Fill Severity field (1, 2, 3, 4)
    ie.Document.getElementById("mx654").Value = "3"
    'Fill Service type (FBN, ESCALATE
    ie.Document.getElementById("mx708").Value = "ESCALATE"
    'Fill in Target Start field
    ie.Document.getElementById("mx803").Value = lastChange
    
    'Select an owner
    ie.Document.getElementById("mx2156_image").Click
    'Do While ie.ReadyState = 4: DoEvents: Loop
    
    'Click on owner group tab
    'ie.Document.getElementsByClassName("text tablabeloff off")(3).Click
    
    'Click filter
    'ie.Document.getElementById("mx3383").Click
    'Click on input field
    'ie.Document.getElementById("mx3716").Click
    
End If


End Sub
