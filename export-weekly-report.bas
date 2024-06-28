Sub CreatePDFs()

    ' Declare variables
    Dim wsReport As Worksheet
    Dim wsEmployees As Worksheet
    Dim rngEmployee As Range
    Dim rngStartEmployee As Range
    Dim rowCount As Long
    Dim i As Long
    Dim folderPath As String
    Dim tempPdfFilePath As String

    ' Stop screen updating to improve performance
    Application.ScreenUpdating = False

    ' Establish references
    Set wsReport = ActiveWorkbook.Sheets("REPORT")
    Set wsEmployees = ActiveWorkbook.Sheets("EMPLOYEES")
    Set rngEmployee = wsReport.Range("D11")
    Set rngStartEmployee = wsEmployees.Range("A7")
    folderPath = "C:\Users\User\Documents\Report [Name].pdf"

    ' Count number of employees
    rowCount = rngStartEmployee.CurrentRegion.Rows.Count - 2

    ' Loop to create the report for each employee and export it to PDF
    For i = 1 To rowCount
        rngEmployee.Value = rngStartEmployee.Offset(i - 1, 0).Value
        tempPdfFilePath = Replace(folderPath, "[Name]", rngEmployee.Value)
        wsReport.ExportAsFixedFormat Type:=xlTypePDF, Filename:=tempPdfFilePath
    Next i

    ' Re-enable screen updating
    Application.ScreenUpdating = True

End Sub
