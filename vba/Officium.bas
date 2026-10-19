Attribute VB_Name = "Officium"
' Officium - calendar reports for classic Outlook.
Option Explicit

' Windows clipboard API, declared for 32- and 64-bit Office.
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWndNewOwner As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal length As LongPtr)

' Asks for period, customer and grouping and shows the totals of the matching
' appointments of the default calendar.
Public Sub OfficiumReport()
    Dim fromDate As Date, untilDate As Date, customer As String, grouping As Long
    If Not AskPeriod(fromDate, untilDate) Then Exit Sub
    If Not AskCustomer(customer) Then Exit Sub
    If Not AskChoice("Gruppieren nach:" & vbCrLf & vbCrLf & "1 = Monat" & vbCrLf & _
        "2 = Quartal" & vbCrLf & "3 = keine Gruppierung", 3, ogMonth, grouping) Then Exit Sub

    On Error GoTo Failed
    Dim title As String, report As Variant
    title = ReportTitle(customer, fromDate, untilDate)
    report = CollectReport(fromDate, untilDate, customer, grouping)
    If IsEmpty(report) Then
        MsgBox "Keine passenden Termine." & vbCrLf & vbCrLf & title, vbInformation, "Officium"
        Exit Sub
    End If
    Dim csvPath As String, results As String, xl As Object
    csvPath = DocumentsFolder() & "\" & CsvFileName(customer, fromDate, untilDate)
    results = SaveCsv(csvPath, report)
    If CopyToClipboard(ClipboardText(title, report)) Then
        results = results & vbCrLf & "Bericht als Text in der Zwischenablage."
    Else
        results = results & vbCrLf & "Zwischenablage war belegt, Text nicht kopiert."
    End If
    Set xl = StartExcel()
    If Not xl Is Nothing Then results = results & vbCrLf & "Nach OK öffnet sich der Bericht in Excel."

    ' The message comes before Excel opens: shown later it could end up behind
    ' the Excel window and Outlook would seem to hang.
    MsgBox title & vbCrLf & vbCrLf & TotalsText(report) & vbCrLf & vbCrLf & results, _
        vbInformation, "Officium"
    If xl Is Nothing Then
        OpenFile csvPath
    ElseIf Not ShowInExcel(xl, title, report) Then
        OpenFile csvPath
    End If
    Exit Sub

Failed:
    MsgBox "Officium wurde mit einem Fehler abgebrochen:" & vbCrLf & vbCrLf & _
        Err.Description & " (" & Err.Number & ")", vbCritical, "Officium"
End Sub

' Report (see COL_* in OfficiumCore) of the appointments that start in
' [fromDate, untilDate), belong to the customer and count as work.
' Empty if there are none.
Private Function CollectReport(ByVal fromDate As Date, ByVal untilDate As Date, _
                               ByVal customer As String, ByVal grouping As OfficiumGrouping) As Variant
    Dim found As Outlook.Items, item As Object, matches As New Collection
    Set found = AppointmentsBetween(fromDate, untilDate)
    Set item = found.GetFirst
    Do While Not item Is Nothing
        If item.Class = olAppointment Then
            ' Restrict also returns appointments that started before the period;
            ' they belong to the previous one.
            If item.Start >= fromDate Then
                If CountsAsWork(item.AllDayEvent, item.MeetingStatus, item.ResponseStatus) Then
                    If MatchesCustomer(customer, item.Subject, item.Categories) Then
                        matches.Add Array(item.Start, item.End, DateDiff("n", item.Start, item.End), _
                            item.Subject, item.Categories, item.Body, GroupKey(item.Start, grouping))
                    End If
                End If
            End If
        End If
        Set item = found.GetNext
    Loop
    If matches.Count = 0 Then Exit Function

    Dim report() As Variant, match As Variant, r As Long, c As Long
    ReDim report(1 To matches.Count, 1 To COL_COUNT)
    For Each match In matches
        r = r + 1
        For c = 1 To COL_COUNT
            report(r, c) = match(c - 1)
        Next
    Next
    CollectReport = report
End Function

' Puts Unicode text on the clipboard via the Windows API. MSForms.DataObject
' is avoided: it pastes "??" while Explorer windows are open.
' Returns False if another program holds the clipboard.
Private Function CopyToClipboard(ByVal content As String) As Boolean
    Const GMEM_MOVEABLE_ZEROINIT As Long = &H42
    Const CF_UNICODETEXT As Long = 13
    Dim hMem As LongPtr, pMem As LongPtr

    ' Two extra zero bytes terminate the UTF-16 text.
    hMem = GlobalAlloc(GMEM_MOVEABLE_ZEROINIT, LenB(content) + 2)
    If hMem = 0 Then Exit Function
    pMem = GlobalLock(hMem)
    If pMem = 0 Then
        GlobalFree hMem
        Exit Function
    End If
    CopyMemory pMem, StrPtr(content), LenB(content)
    GlobalUnlock hMem

    If OpenClipboard(0) = 0 Then
        GlobalFree hMem
        Exit Function
    End If
    EmptyClipboard
    ' On success the memory belongs to Windows and must not be freed.
    If SetClipboardData(CF_UNICODETEXT, hMem) = 0 Then
        GlobalFree hMem
    Else
        CopyToClipboard = True
    End If
    CloseClipboard
End Function

' A new, still hidden Excel instance, or Nothing if Excel is not installed.
Private Function StartExcel() As Object
    On Error Resume Next
    Set StartExcel = CreateObject("Excel.Application")
End Function

' Fills a new workbook in the given Excel and shows it: sheet "Termine" with an
' autofilter and a total of the visible rows, sheet "Summen" with the totals
' per group. On an error Excel is closed again and False returned.
Private Function ShowInExcel(ByVal xl As Object, ByVal title As String, ByVal report As Variant) As Boolean
    On Error GoTo Failed
    Dim book As Object, sheet As Object, grid() As Variant, n As Long, r As Long, c As Long
    Dim headers As Variant
    n = UBound(report, 1)
    Set book = xl.Workbooks.Add
    Set sheet = book.Worksheets(1)
    sheet.Name = "Termine"

    ' Formats before values: text columns stay text (no date from "2026-10",
    ' no formula from "=..."), durations keep counting past 24 hours.
    ' NumberFormat and Formula always take the English codes.
    sheet.Columns(1).NumberFormat = "dd.mm.yyyy"
    sheet.Range("B:C").NumberFormat = "hh:mm"
    sheet.Columns(4).NumberFormat = "[h]:mm"
    sheet.Range("E:H").NumberFormat = "@"

    headers = Array("Datum", "Beginn", "Ende", "Dauer", "Gruppe", "Betreff", "Kategorien", "Text")
    ReDim grid(1 To n + 1, 1 To 8)
    For c = 1 To 8
        grid(1, c) = headers(c - 1)
    Next
    For r = 1 To n
        grid(r + 1, 1) = DateSerial(Year(report(r, COL_START)), Month(report(r, COL_START)), Day(report(r, COL_START)))
        grid(r + 1, 2) = TimeSerial(Hour(report(r, COL_START)), Minute(report(r, COL_START)), 0)
        grid(r + 1, 3) = TimeSerial(Hour(report(r, COL_END)), Minute(report(r, COL_END)), 0)
        grid(r + 1, 4) = report(r, COL_MINUTES) / 1440
        grid(r + 1, 5) = report(r, COL_GROUP)
        grid(r + 1, 6) = report(r, COL_SUBJECT)
        grid(r + 1, 7) = report(r, COL_CATEGORIES)
        ' Excel breaks lines at LF only; a cell holds at most 32767 characters.
        grid(r + 1, 8) = Left$(Replace(TrimText(report(r, COL_BODY)), vbCr, ""), 32000)
    Next
    ' One assignment for all cells; cell by cell is slow with many appointments.
    sheet.Range("A1").Resize(n + 1, 8).Value = grid
    sheet.Range("A1").Resize(1, 8).Font.Bold = True
    sheet.Range("A1").Resize(n + 1, 8).AutoFilter
    sheet.Cells(n + 3, 3).Value = "Summe"
    ' 109 = SUM over the rows the autofilter shows.
    sheet.Cells(n + 3, 4).Formula = "=SUBTOTAL(109,D2:D" & (n + 1) & ")"
    sheet.Cells(n + 3, 3).Resize(1, 2).Font.Bold = True
    sheet.Range("A:G").Columns.AutoFit
    sheet.Columns(8).ColumnWidth = 80
    sheet.Columns(8).WrapText = False

    Set sheet = book.Worksheets.Add(After:=sheet)
    sheet.Name = "Summen"
    sheet.Columns(1).NumberFormat = "@"
    sheet.Columns(2).NumberFormat = "[h]:mm"
    sheet.Cells(1, 1).Value = title
    sheet.Range("A3:C3").Value = Array("Gruppe", "Dauer", "Minuten")
    Dim keys() As String, sums() As Long, groups As Long, g As Long, outRow As Long, total As Long
    groups = SumByGroup(report, keys, sums)
    outRow = 3
    For g = 1 To groups
        total = total + sums(g)
        If Len(keys(g)) > 0 Then
            outRow = outRow + 1
            sheet.Cells(outRow, 1).Value = keys(g)
            sheet.Cells(outRow, 2).Value = sums(g) / 1440
            sheet.Cells(outRow, 3).Value = sums(g)
        End If
    Next
    outRow = outRow + 1
    sheet.Cells(outRow, 1).Value = "Gesamt"
    sheet.Cells(outRow, 2).Value = total / 1440
    sheet.Cells(outRow, 3).Value = total
    sheet.Range("A1,A3:C3").Font.Bold = True
    sheet.Cells(outRow, 1).Resize(1, 3).Font.Bold = True
    sheet.Range("A3:C" & outRow).Columns.AutoFit

    book.Worksheets(1).Activate
    xl.Visible = True
    ShowInExcel = True
    Exit Function

Failed:
    MsgBox "Excel-Fehler: " & Err.Description & vbCrLf & vbCrLf & _
        "Officium öffnet stattdessen die CSV-Datei.", vbExclamation, "Officium"
    ' Resume leaves the error handler; inside it On Error Resume Next has no effect.
    Resume CloseExcel
CloseExcel:
    On Error Resume Next
    xl.DisplayAlerts = False
    xl.Quit
End Function

' Saves the report as CSV. Returns a status line for the final message.
Private Function SaveCsv(ByVal filePath As String, ByVal report As Variant) As String
    On Error GoTo Failed
    SaveUtf8 filePath, CsvText(report)
    SaveCsv = "CSV gespeichert: " & filePath
    Exit Function

Failed:
    ' Typically the file of an earlier run is still open in Excel.
    SaveCsv = "CSV nicht gespeichert (" & Err.Description & "): " & filePath
End Function

' Opens a file like a double click; does nothing if Windows Script Host is
' blocked or the file is missing.
Private Sub OpenFile(ByVal filePath As String)
    On Error Resume Next
    CreateObject("WScript.Shell").Run """" & filePath & """"
End Sub

' Writes text as UTF-8 with byte order mark, which German Excel needs to show
' umlauts. ADODB.Stream adds the mark itself.
Private Sub SaveUtf8(ByVal filePath As String, ByVal content As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile filePath, 2  ' adSaveCreateOverWrite
    stream.Close
End Sub

' The user's documents folder, also when it is redirected to OneDrive.
Private Function DocumentsFolder() As String
    On Error Resume Next
    DocumentsFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
    On Error GoTo 0
    If Len(DocumentsFolder) = 0 Then DocumentsFolder = Environ$("USERPROFILE") & "\Documents"
End Function

' Asks for a preset period or a free range of days; returns the half-open
' range [fromDate, untilDate). False if the user cancels.
Private Function AskPeriod(ByRef fromDate As Date, ByRef untilDate As Date) As Boolean
    Dim choice As Long
    If Not AskChoice("Zeitraum:" & vbCrLf & vbCrLf & _
        "1 = dieser Monat" & vbCrLf & "2 = letzter Monat" & vbCrLf & _
        "3 = dieses Quartal" & vbCrLf & "4 = letztes Quartal" & vbCrLf & _
        "5 = selbst wählen", 5, opLastMonth, choice) Then Exit Function

    If choice < 5 Then
        PeriodRange choice, Date, fromDate, untilDate
    Else
        Dim toDate As Date
        If Not AskDate("Von (erster Tag):", DateSerial(Year(Date), Month(Date), 1), fromDate) Then Exit Function
        If Not AskDate("Bis (letzter Tag, einschließlich):", DateAdd("m", 1, fromDate) - 1, toDate) Then Exit Function
        If toDate < fromDate Then
            MsgBox "Das Bis-Datum liegt vor dem Von-Datum.", vbExclamation, "Officium"
            Exit Function
        End If
        untilDate = toDate + 1
    End If
    AskPeriod = True
End Function

' Asks for the customer, suggesting the last one. An empty answer means all
' appointments. False if the user cancels.
Private Function AskCustomer(ByRef customer As String) As Boolean
    Dim answer As String
    answer = InputBox("Kunde: Outlook-Kategorie oder Stichwort im Betreff." & vbCrLf & vbCrLf & _
        "Leer lassen für alle Termine im Zeitraum.", "Officium", _
        GetSetting("Officium", "Report", "Customer", ""))
    ' Cancel returns a null string, OK with an empty field an empty one.
    If StrPtr(answer) = 0 Then Exit Function
    customer = Trim$(answer)
    SaveSetting "Officium", "Report", "Customer", customer
    AskCustomer = True
End Function

' Asks for a number from 1 to maxChoice. False if the user cancels or enters
' something else.
Private Function AskChoice(ByVal prompt As String, ByVal maxChoice As Long, _
                           ByVal suggestion As Long, ByRef result As Long) As Boolean
    Dim answer As String
    answer = Trim$(InputBox(prompt, "Officium", CStr(suggestion)))
    If Len(answer) = 0 Then Exit Function
    If Not answer Like "#" Or Val(answer) < 1 Or Val(answer) > maxChoice Then
        MsgBox "Bitte eine Zahl von 1 bis " & maxChoice & " eingeben.", vbExclamation, "Officium"
        Exit Function
    End If
    result = CLng(answer)
    AskChoice = True
End Function

' Lists every appointment that overlaps the given days in the Immediate window
' (Ctrl+G), including each occurrence of recurring series.
Public Sub OfficiumListAppointments()
    Dim fromDate As Date, toDate As Date
    If Not AskDate("Von (erster Tag):", Date - Weekday(Date, vbMonday) + 1, fromDate) Then Exit Sub
    If Not AskDate("Bis (letzter Tag, einschließlich):", fromDate + 6, toDate) Then Exit Sub
    If toDate < fromDate Then
        MsgBox "Das Bis-Datum liegt vor dem Von-Datum.", vbExclamation, "Officium"
        Exit Sub
    End If

    Dim found As Outlook.Items
    Set found = AppointmentsBetween(fromDate, toDate + 1)

    Debug.Print "--- Officium: " & Format(fromDate, "Short Date") & " bis " & Format(toDate, "Short Date") & " ---"
    Dim item As Object, total As Long
    Set item = found.GetFirst
    Do While Not item Is Nothing
        If item.Class = olAppointment Then
            total = total + 1
            Debug.Print Format(item.Start, "ddddd hh:nn") & " - " & Format(item.End, "ddddd hh:nn") _
                & vbTab & DateDiff("n", item.Start, item.End) & " min" _
                & vbTab & item.Subject _
                & vbTab & "[" & item.Categories & "]"
        End If
        Set item = found.GetNext
    Loop
    Debug.Print total & " Termine"

    MsgBox total & " Termine gefunden. Die Liste steht im Direktfenster (Strg+G im VBA-Editor).", _
        vbInformation, "Officium"
End Sub

' Appointments of the default calendar that overlap [fromTime, toTime),
' with recurring series expanded into single occurrences.
' Count is meaningless on the result while recurrences are included:
' iterate with GetFirst/GetNext.
Private Function AppointmentsBetween(ByVal fromTime As Date, ByVal toTime As Date) As Outlook.Items
    Dim all As Outlook.Items
    Set all = Application.Session.GetDefaultFolder(olFolderCalendar).Items
    ' Sort before IncludeRecurrences, otherwise series are returned only once.
    all.Sort "[Start]"
    all.IncludeRecurrences = True
    Set AppointmentsBetween = all.Restrict( _
        "[Start] < '" & Format(toTime, "ddddd hh:nn") & "' AND " & _
        "[End] > '" & Format(fromTime, "ddddd hh:nn") & "'")
End Function

' Asks for a date; returns False if the user cancels or enters no valid date.
Private Function AskDate(ByVal prompt As String, ByVal suggestion As Date, ByRef result As Date) As Boolean
    Dim answer As String
    answer = InputBox(prompt, "Officium", Format(suggestion, "Short Date"))
    If Len(answer) = 0 Then Exit Function
    If Not IsDate(answer) Then
        MsgBox """" & answer & """ ist kein gültiges Datum.", vbExclamation, "Officium"
        Exit Function
    End If
    result = DateValue(CDate(answer))
    AskDate = True
End Function
