Attribute VB_Name = "modTerminPlanner"
Option Explicit

Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const DATA_SHEET As String = "Appointments"
Private Const GUIDE_SHEET As String = "Guide"
Private Const APPOINTMENT_TABLE As String = "tblAppointments"

' These column constants document the table contract shared by every macro.
Private Const COL_ID As Long = 1
Private Const COL_DATE As Long = 2
Private Const COL_START As Long = 3
Private Const COL_END As Long = 4
Private Const COL_SUBJECT As Long = 5
Private Const COL_CLIENT As Long = 6
Private Const COL_CATEGORY As Long = 7
Private Const COL_STATUS As Long = 8
Private Const COL_LOCATION As Long = 9
Private Const COL_CONTACT As Long = 10
Private Const COL_NOTES As Long = 11
Private Const COL_REMINDER As Long = 12
Private Const COL_CREATED As Long = 13
Private Const COL_UPDATED As Long = 14

Public Sub SetupSession()
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.StatusBar = "TerminPlanner ready · local workbook"
End Sub

Public Sub AddAppointment()
    Dim appointmentDate As Variant
    Dim startTime As Variant
    Dim endTime As Variant
    Dim subject As String
    Dim client As String
    Dim table As ListObject
    Dim row As ListRow

    appointmentDate = FormRange("FormDate").Value
    startTime = FormRange("FormStart").Value
    endTime = FormRange("FormEnd").Value
    subject = Trim$(CStr(FormRange("FormSubject").Value))
    client = Trim$(CStr(FormRange("FormClient").Value))

    If Not IsDate(appointmentDate) Then
        ShowValidation "Choose a valid appointment date.", "FormDate"
        Exit Sub
    End If
    If Not IsDate(startTime) Or Not IsDate(endTime) Then
        ShowValidation "Choose a valid start and end time.", "FormStart"
        Exit Sub
    End If
    If TimeValue(endTime) <= TimeValue(startTime) Then
        ShowValidation "The end time must be later than the start time.", "FormEnd"
        Exit Sub
    End If
    If Len(subject) = 0 Or Len(client) = 0 Then
        ShowValidation "Subject and client are required.", "FormSubject"
        Exit Sub
    End If

    Set table = Worksheets(DATA_SHEET).ListObjects(APPOINTMENT_TABLE)
    If table.ListRows.Count = 1 And Len(CStr(table.ListRows(1).Range.Cells(1, COL_ID).Value)) = 0 Then
        Set row = table.ListRows(1)
    Else
        Set row = table.ListRows.Add
    End If

    With row.Range
        .Cells(1, COL_ID).Value = CreateAppointmentId()
        .Cells(1, COL_DATE).Value = DateValue(appointmentDate)
        .Cells(1, COL_START).Value = TimeValue(startTime)
        .Cells(1, COL_END).Value = TimeValue(endTime)
        .Cells(1, COL_SUBJECT).Value = subject
        .Cells(1, COL_CLIENT).Value = client
        .Cells(1, COL_CATEGORY).Value = FormRange("FormCategory").Value
        .Cells(1, COL_STATUS).Value = FormRange("FormStatus").Value
        .Cells(1, COL_LOCATION).Value = FormRange("FormLocation").Value
        .Cells(1, COL_CONTACT).Value = FormRange("FormContact").Value
        .Cells(1, COL_NOTES).Value = FormRange("FormNotes").Value
        .Cells(1, COL_REMINDER).Value = FormRange("FormReminder").Value
        .Cells(1, COL_CREATED).Value = Now
        .Cells(1, COL_UPDATED).Value = Now
    End With

    SortAppointments table
    ClearForm
    RefreshPlanner
    MsgBox "Appointment saved locally.", vbInformation, "TerminPlanner"
End Sub

Public Sub ClearForm()
    FormRange("FormDate").Value = Date
    FormRange("FormStart").Value = TimeSerial(9, 0, 0)
    FormRange("FormEnd").Value = TimeSerial(9, 30, 0)
    FormRange("FormSubject").ClearContents
    FormRange("FormClient").ClearContents
    FormRange("FormCategory").Value = "Meeting"
    FormRange("FormStatus").Value = "Planned"
    FormRange("FormLocation").ClearContents
    FormRange("FormContact").ClearContents
    FormRange("FormNotes").ClearContents
    FormRange("FormReminder").Value = "15 minutes"
    FormRange("FormSubject").Select
End Sub

Public Sub RefreshPlanner()
    Dim table As ListObject
    Dim dataRow As ListRow
    Dim dashboard As Worksheet
    Dim outputRow As Long
    Dim appointmentDate As Variant
    Dim status As String

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Set table = Worksheets(DATA_SHEET).ListObjects(APPOINTMENT_TABLE)
    Set dashboard = Worksheets(DASHBOARD_SHEET)
    dashboard.Range("B27:H36").ClearContents
    outputRow = 27

    If Not table.DataBodyRange Is Nothing Then
        SortAppointments table
        For Each dataRow In table.ListRows
            appointmentDate = dataRow.Range.Cells(1, COL_DATE).Value
            status = CStr(dataRow.Range.Cells(1, COL_STATUS).Value)
            If IsDate(appointmentDate) And DateValue(appointmentDate) >= Date _
                    And LCase$(status) <> "cancelled" Then
                dashboard.Cells(outputRow, 2).Value = appointmentDate
                dashboard.Cells(outputRow, 3).Value = dataRow.Range.Cells(1, COL_START).Value
                dashboard.Cells(outputRow, 4).Value = dataRow.Range.Cells(1, COL_SUBJECT).Value
                dashboard.Cells(outputRow, 5).Value = dataRow.Range.Cells(1, COL_CLIENT).Value
                dashboard.Cells(outputRow, 6).Value = dataRow.Range.Cells(1, COL_CATEGORY).Value
                dashboard.Cells(outputRow, 7).Value = status
                dashboard.Cells(outputRow, 8).Value = dataRow.Range.Cells(1, COL_LOCATION).Value
                outputRow = outputRow + 1
                If outputRow > 36 Then Exit For
            End If
        Next dataRow
    End If

    dashboard.Calculate
    Application.StatusBar = "TerminPlanner refreshed · " & Format$(Now, "hh:nn")
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
CleanFail:
    Application.ScreenUpdating = True
    MsgBox "The planner could not be refreshed: " & Err.Description, vbExclamation, "TerminPlanner"
End Sub

Public Sub DeleteSelectedAppointment()
    Dim table As ListObject
    Dim selectedIndex As Long

    If ActiveSheet.Name <> DATA_SHEET Then
        MsgBox "Open Appointments and select a cell in the row to delete.", vbInformation, "TerminPlanner"
        Worksheets(DATA_SHEET).Activate
        Exit Sub
    End If

    Set table = ActiveSheet.ListObjects(APPOINTMENT_TABLE)
    If table.DataBodyRange Is Nothing Or Intersect(Selection, table.DataBodyRange) Is Nothing Then
        MsgBox "Select a cell inside an appointment row first.", vbInformation, "TerminPlanner"
        Exit Sub
    End If

    selectedIndex = Selection.Row - table.DataBodyRange.Row + 1
    If MsgBox("Delete the selected appointment?", vbQuestion + vbYesNo, "TerminPlanner") = vbYes Then
        table.ListRows(selectedIndex).Delete
        RefreshPlanner
    End If
End Sub

Public Sub SearchAppointments()
    Dim query As String
    Dim table As ListObject
    Dim row As ListRow
    Dim haystack As String
    Dim cell As Range

    query = Trim$(InputBox("Search subject, client, category, status or location:", "TerminPlanner search"))
    If Len(query) = 0 Then Exit Sub
    Set table = Worksheets(DATA_SHEET).ListObjects(APPOINTMENT_TABLE)
    Worksheets(DATA_SHEET).Activate

    If table.DataBodyRange Is Nothing Then Exit Sub
    table.DataBodyRange.EntireRow.Hidden = False
    For Each row In table.ListRows
        haystack = vbNullString
        For Each cell In row.Range.Cells
            haystack = haystack & " " & CStr(cell.Value)
        Next cell
        row.Range.EntireRow.Hidden = (InStr(1, haystack, query, vbTextCompare) = 0)
    Next row
    Application.StatusBar = "Search active · use Clear search to show every appointment"
End Sub

Public Sub ClearSearch()
    Dim table As ListObject
    Set table = Worksheets(DATA_SHEET).ListObjects(APPOINTMENT_TABLE)
    If Not table.DataBodyRange Is Nothing Then table.DataBodyRange.EntireRow.Hidden = False
    On Error Resume Next
    If table.ShowAutoFilter Then table.AutoFilter.ShowAllData
    On Error GoTo 0
    Application.StatusBar = "Search cleared"
End Sub

Public Sub GoToToday()
    Dim table As ListObject
    Dim row As ListRow
    Set table = Worksheets(DATA_SHEET).ListObjects(APPOINTMENT_TABLE)
    Worksheets(DATA_SHEET).Activate
    If table.DataBodyRange Is Nothing Then Exit Sub

    For Each row In table.ListRows
        If IsDate(row.Range.Cells(1, COL_DATE).Value) Then
            If DateValue(row.Range.Cells(1, COL_DATE).Value) >= Date Then
                row.Range.Cells(1, COL_DATE).Select
                Exit Sub
            End If
        End If
    Next row
    MsgBox "No upcoming appointment was found.", vbInformation, "TerminPlanner"
End Sub

Public Sub ExportUpcomingCsv()
    Dim table As ListObject
    Dim row As ListRow
    Dim filePath As String
    Dim fileNumber As Integer
    Dim values(1 To 7) As String

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save the workbook before exporting.", vbInformation, "TerminPlanner"
        Exit Sub
    End If

    filePath = ThisWorkbook.Path & Application.PathSeparator & _
               "TerminPlanner_Export_" & Format$(Date, "yyyymmdd") & ".csv"
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, "Date,Start,End,Subject,Client,Status,Location"

    Set table = Worksheets(DATA_SHEET).ListObjects(APPOINTMENT_TABLE)
    If Not table.DataBodyRange Is Nothing Then
        For Each row In table.ListRows
            If IsDate(row.Range.Cells(1, COL_DATE).Value) Then
                If DateValue(row.Range.Cells(1, COL_DATE).Value) >= Date Then
                    values(1) = Format$(row.Range.Cells(1, COL_DATE).Value, "yyyy-mm-dd")
                    values(2) = Format$(row.Range.Cells(1, COL_START).Value, "hh:nn")
                    values(3) = Format$(row.Range.Cells(1, COL_END).Value, "hh:nn")
                    values(4) = CStr(row.Range.Cells(1, COL_SUBJECT).Value)
                    values(5) = CStr(row.Range.Cells(1, COL_CLIENT).Value)
                    values(6) = CStr(row.Range.Cells(1, COL_STATUS).Value)
                    values(7) = CStr(row.Range.Cells(1, COL_LOCATION).Value)
                    Print #fileNumber, CsvLine(values)
                End If
            End If
        Next row
    End If
    Close #fileNumber
    MsgBox "Export created:" & vbCrLf & filePath, vbInformation, "TerminPlanner"
End Sub

Public Sub OpenGuide()
    Worksheets(GUIDE_SHEET).Activate
End Sub

Private Sub SortAppointments(ByVal table As ListObject)
    If table.DataBodyRange Is Nothing Then Exit Sub
    With table.Sort
        .SortFields.Clear
        .SortFields.Add Key:=table.ListColumns(COL_DATE).Range, Order:=xlAscending
        .SortFields.Add Key:=table.ListColumns(COL_START).Range, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
End Sub

Private Sub ShowValidation(ByVal message As String, ByVal rangeName As String)
    MsgBox message, vbExclamation, "TerminPlanner"
    Worksheets(DASHBOARD_SHEET).Activate
    FormRange(rangeName).Select
End Sub

Private Function FormRange(ByVal rangeName As String) As Range
    Set FormRange = ThisWorkbook.Names(rangeName).RefersToRange
End Function

Private Function CreateAppointmentId() As String
    Randomize
    CreateAppointmentId = Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(Int(Rnd() * 900 + 100), "000")
End Function

Private Function CsvLine(ByRef values() As String) As String
    Dim index As Long
    Dim output As String
    For index = LBound(values) To UBound(values)
        If index > LBound(values) Then output = output & ","
        output = output & """" & Replace(values(index), """", """""") & """"
    Next index
    CsvLine = output
End Function
