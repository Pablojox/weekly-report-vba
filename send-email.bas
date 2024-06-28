Sub SendAllEmails()

    'Declare variables
    Dim wsEmployees As Worksheet
    Dim wsTemplate As Worksheet
    Dim rngStartEmployee As Range
    Dim rngEmails As Range
    Dim rowCount As Long
    Dim i As Long
    Dim folderPath As String
    Dim employeeName As String
    Dim emailAddress As String
    Dim pdfFilePath As String

    'Establish references for the worksheets
    Set wsTemplate = ActiveWorkbook.Sheets("EMAIL")
    Set wsEmployees = ActiveWorkbook.Sheets("EMPLOYEES")

    'Set the origin of the employee list
    Set rngStartEmployee = wsEmployees.Range("A7")
    Set rngEmails = wsTemplate.Range("A2")

    'Count number of employees
    rowCount = rngStartEmployee.CurrentRegion.Rows.Count - 2

    'Set the folder path for exporting the reports
    folderPath = "C:\Users\User\Documents\Report [Name].pdf"

    'Create loop to get the employee name, find the employee's email address and send the emails
    For i = 1 To rowCount
        employeeName = rngStartEmployee.Offset(i - 1, 0).Value
        emailAddress = FindEmail(wsTemplate, employeeName)

        If emailAddress <> "" Then
            ' Replace [Name] with the current employee name to get the PDF file path
            pdfFilePath = Replace(folderPath, "[Name]", employeeName)

            ' Send email
            Call SendEmailToEmployee(employeeName, pdfFilePath, emailAddress)
        End If
    Next i

End Sub

Function FindEmail(ws As Worksheet, employeeName As String) As String

    'Declare variables
    Dim lastRow As Long
    Dim i As Long

    'Find the last row in the "EMAIL" sheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    'Search for the employee's email address
    For i = 3 To lastRow ' Data starts at row 3
        If ws.Cells(i, 1).Value = employeeName Then
            If ws.Cells(i, 4).Value <> "" Then
                FindEmail = ws.Cells(i, 4).Value
                Exit Function
            End If
        End If
    Next i

    ' If the email address is not found, return an empty string
    FindEmail = ""
End Function

Sub SendEmailToEmployee(employeeName As String, pdfFilePath As String, emailAddress As String)

    'Declare variables
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim currentWeekNumber As Integer
    Dim signatureHTML As String

    'Get the current week number and subtract 1
    currentWeekNumber = Format(Date, "ww") - 1

    'Create an instance of Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0)

    'Configure the email
    With outlookMail
        .To = emailAddress
        .Subject = "Report " & employeeName & " - Week " & currentWeekNumber
        .BodyFormat = 2 ' olFormatHTML
        .Display ' Display the email to load the signature

        ' Get the Outlook signature as HTML
        signatureHTML = .HTMLBody

        ' Set the email body with the signature
        .HTMLBody = "Good morning,<br><br>" & _
                    "Attached are the individual data for week " & currentWeekNumber & ".<br><br>" & _
                    "Best regards,<br><br>" & _
                    signatureHTML
        
        ' Add the PDF attachment and send the email
        .Attachments.Add pdfFilePath
        .Send
    End With

    ' Release the objects
    Set outlookMail = Nothing
    Set outlookApp = Nothing
End Sub
