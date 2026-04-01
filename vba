Option Compare Database
Option Explicit

Public Sub ImportPeakFilesManual()
    Dim grossFile As String
    Dim highFile As String
    Dim ninetyTenFile As String

    grossFile = PickExcelFile("Select the Gross Peak Excel file")
    If grossFile = "" Then
        MsgBox "Import cancelled."
        Exit Sub
    End If

    highFile = PickExcelFile("Select the High Peak Excel file")
    If highFile = "" Then
        MsgBox "Import cancelled."
        Exit Sub
    End If

    ninetyTenFile = PickExcelFile("Select the 90/10 Excel file")
    If ninetyTenFile = "" Then
        MsgBox "Import cancelled."
        Exit Sub
    End If

    ImportOnePeakWorkbookManual grossFile, "Gross Peak"
    ImportOnePeakWorkbookManual highFile, "High Peak"
    ImportOnePeakWorkbookManual ninetyTenFile, "90/10"

    MsgBox "All three files were imported successfully."
End Sub


Private Function PickExcelFile(dialogTitle As String) As String
    Dim fd As FileDialog

    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker

    With fd
        .Title = dialogTitle
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"

        If .Show = -1 Then
            PickExcelFile = .SelectedItems(1)
        Else
            PickExcelFile = ""
        End If
    End With
End Function


Private Sub ImportOnePeakWorkbookManual(ByVal filePath As String, ByVal targetField As String)
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object

    Dim lastRow As Long
    Dim r As Long
    Dim monthNum As Long

    Dim yr As Variant
    Dim cellVal As Variant

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlWb = xlApp.Workbooks.Open(filePath)
    Set xlWs = xlWb.Worksheets(1)

    lastRow = xlWs.Cells(xlWs.Rows.Count, 1).End(-4162).Row   ' xlUp

    ' Expected layout:
    ' Row 1 = title/date
    ' Row 2 = headers (Year, Jan, Feb, ... Dec)
    ' Row 3+ = data

    For r = 3 To lastRow
        yr = xlWs.Cells(r, 1).Value

        If Not IsNull(yr) And yr <> "" And IsNumeric(yr) Then
            For monthNum = 1 To 12
                cellVal = xlWs.Cells(r, monthNum + 1).Value

                If Not IsNull(cellVal) And cellVal <> "" And IsNumeric(cellVal) Then
                    UpdateCorporateLoadManual CLng(yr), monthNum, CLng(cellVal), targetField
                End If
            Next monthNum
        End If
    Next r

    xlWb.Close False
    xlApp.Quit

    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
End Sub


Private Sub UpdateCorporateLoadManual(ByVal yr As Long, ByVal monthNum As Long, ByVal loadVal As Long, ByVal targetField As String)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String

    Set db = CurrentDb

    sqlText = "SELECT * FROM [Corporate Load Forecast] " & _
              "WHERE [Year] = " & yr & " AND [Month] = " & monthNum

    Set rs = db.OpenRecordset(sqlText, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs.Fields("Year").Value = yr
        rs.Fields("Month").Value = monthNum
        rs.Fields(targetField).Value = loadVal
        rs.Update
    Else
        rs.Edit
        rs.Fields(targetField).Value = loadVal
        rs.Update
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub