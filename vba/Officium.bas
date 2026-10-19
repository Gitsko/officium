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

' Windows API for the progress window, a plain text window without a form.
Private Declare PtrSafe Function CreateWindowExW Lib "user32" (ByVal exStyle As Long, ByVal classPtr As LongPtr, _
    ByVal textPtr As LongPtr, ByVal windowStyle As Long, ByVal posX As Long, ByVal posY As Long, _
    ByVal sizeX As Long, ByVal sizeY As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, _
    ByVal hInstance As LongPtr, ByVal createParam As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowTextW Lib "user32" (ByVal hWnd As LongPtr, ByVal textPtr As LongPtr) As Long
Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal metric As Long) As Long
Private Declare PtrSafe Function SendMessageW Lib "user32" _
    (ByVal hWnd As LongPtr, ByVal message As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal moduleName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal objectIndex As Long) As LongPtr

' The step ShowInExcel is working on, named in its error message.
Private excelStep As String
' The progress window while a report is made, 0 when none is shown.
Private progressWindow As LongPtr
' True while OfficiumReport runs; the progress window lets Outlook take clicks.
Private reportRunning As Boolean

' Asks for period, customer and grouping and shows the totals of the matching
' appointments of all calendars, including shared ones.
Public Sub OfficiumReport()
    If reportRunning Then Exit Sub
    reportRunning = True
    MakeReport
    CloseProgress
    reportRunning = False
End Sub

' Asks the questions and makes the report; OfficiumReport guards it.
Private Sub MakeReport()
    Dim fromDate As Date, untilDate As Date, customer As String, grouping As Long
    If Not AskPeriod(fromDate, untilDate) Then Exit Sub
    If Not AskCustomer(customer) Then Exit Sub
    If Not AskChoice("Gruppieren nach:" & vbCrLf & vbCrLf & "1 = Monat" & vbCrLf & _
        "2 = Quartal" & vbCrLf & "3 = keine Gruppierung", 3, ogMonth, grouping) Then Exit Sub

    On Error GoTo Failed
    Dim title As String, report As Variant, calendars As Collection, skipped As String, note As String
    title = ReportTitle(customer, fromDate, untilDate)
    ShowProgress "Kalender werden gesucht"
    Set calendars = AllCalendars(skipped)
    report = CollectReport(calendars, fromDate, untilDate, customer, grouping, skipped)
    note = CalendarNote(calendars.Count, skipped)
    If IsEmpty(report) Then
        CloseProgress
        MsgBox "Keine passenden Termine." & vbCrLf & vbCrLf & title & vbCrLf & vbCrLf & note, _
            vbInformation, "Officium"
        Exit Sub
    End If
    Dim csvPath As String, results As String, xl As Object
    ShowProgress "CSV-Datei und Zwischenablage"
    csvPath = DocumentsFolder() & "\" & CsvFileName(customer, fromDate, untilDate)
    results = SaveCsv(csvPath, report)
    If CopyToClipboard(ClipboardText(title, report)) Then
        results = results & vbCrLf & "Bericht als Text in der Zwischenablage."
    Else
        results = results & vbCrLf & "Zwischenablage war belegt, Text nicht kopiert."
    End If
    ShowProgress "Excel wird gestartet"
    Set xl = StartExcel()
    If Not xl Is Nothing Then results = results & vbCrLf & "Nach OK öffnet sich der Bericht in Excel."

    ' The message comes before Excel opens: shown later it could end up behind
    ' the Excel window and Outlook would seem to hang.
    CloseProgress
    MsgBox title & vbCrLf & vbCrLf & TotalsText(report) & vbCrLf & vbCrLf & note & vbCrLf & results, _
        vbInformation, "Officium"
    If xl Is Nothing Then
        OpenFile csvPath
    ElseIf Not ShowInExcel(xl, customer, fromDate, untilDate, report) Then
        OpenFile csvPath
    End If
    Exit Sub

Failed:
    CloseProgress
    MsgBox "Officium wurde mit einem Fehler abgebrochen:" & vbCrLf & vbCrLf & _
        Err.Description & " (" & Err.Number & ")", vbCritical, "Officium"
End Sub

' "Kalender durchsucht: 3." plus the calendars that could not be read.
Private Function CalendarNote(ByVal searched As Long, ByVal skipped As String) As String
    CalendarNote = "Kalender durchsucht: " & searched & "."
    If Len(skipped) > 0 Then CalendarNote = CalendarNote & vbCrLf & "Nicht lesbar: " & skipped
End Function

' Every calendar the user sees in Outlook, each as Array(folder, name): the
' ones in the navigation pane of the calendar view, including shared
' calendars, then the default calendar of each mailbox in the profile.
' A calendar reached both ways is listed once. Calendars of the navigation
' pane that cannot be opened are appended to skipped.
Private Function AllCalendars(ByRef skipped As String) As Collection
    Dim calendars As New Collection, calModule As Object, grp As Object, nav As Object
    Dim st As Object, calendarFolder As Object, calendarPath As String
    ' Without an Outlook window or with a shared calendar that is gone, single
    ' steps fail; the calendars reached so far still count.
    On Error Resume Next
    Set calModule = Application.ActiveExplorer.NavigationPane.Modules.GetNavigationModule(olModuleCalendar)
    If Not calModule Is Nothing Then
        For Each grp In calModule.NavigationGroups
            For Each nav In grp.NavigationFolders
                Set calendarFolder = Nothing
                Set calendarFolder = nav.Folder
                If calendarFolder Is Nothing Then
                    AppendName skipped, nav.DisplayName
                Else
                    calendarPath = ""
                    calendarPath = calendarFolder.FolderPath
                    AddCalendar calendars, calendarFolder, CalendarLabel(nav.DisplayName, calendarPath)
                End If
            Next
        Next
    End If
    For Each st In Application.Session.Stores
        Set calendarFolder = Nothing
        Set calendarFolder = st.GetDefaultFolder(olFolderCalendar)
        If Not calendarFolder Is Nothing Then
            calendarPath = ""
            calendarPath = calendarFolder.FolderPath
            AddCalendar calendars, calendarFolder, CalendarLabel(calendarFolder.Name, calendarPath)
        End If
    Next
    Set AllCalendars = calendars
End Function

' Adds a calendar unless it is already listed. Two different calendars with
' the same name get a number, so they stay apart in the report.
Private Sub AddCalendar(ByVal calendars As Collection, ByVal calendarFolder As Object, _
                        ByVal calendarName As String)
    Dim entry As Variant, uniqueName As String, n As Long
    On Error Resume Next
    uniqueName = Trim$(calendarName)
    n = 1
    For Each entry In calendars
        If StrComp(entry(1), uniqueName, vbTextCompare) = 0 Then
            n = n + 1
            uniqueName = Trim$(calendarName) & " (" & n & ")"
        End If
    Next
    ' The EntryID as key: a duplicate raises error 457 and is skipped.
    calendars.Add Array(calendarFolder, uniqueName), calendarFolder.EntryID
End Sub

' Appends a name to a list separated by commas.
Private Sub AppendName(ByRef list As String, ByVal itemName As String)
    If Len(list) > 0 Then list = list & ", "
    list = list & itemName
End Sub

' Report (see COL_* in OfficiumCore) of the appointments that start in
' [fromDate, untilDate), belong to the customer and count as work, from all
' given calendars. Calendars that fail are listed in skipped. Empty if there
' are no appointments.
Private Function CollectReport(ByVal calendars As Collection, ByVal fromDate As Date, ByVal untilDate As Date, _
                               ByVal customer As String, ByVal grouping As OfficiumGrouping, _
                               ByRef skipped As String) As Variant
    Dim entry As Variant, matches As New Collection, calendarNumber As Long
    For Each entry In calendars
        calendarNumber = calendarNumber + 1
        ShowProgress "Kalender " & calendarNumber & " von " & calendars.Count & ": " & entry(1)
        If Not AddMatches(matches, entry(0), entry(1), fromDate, untilDate, customer, grouping) Then
            AppendName skipped, entry(1)
        End If
    Next
    If matches.Count = 0 Then Exit Function

    Dim report() As Variant, match As Variant, r As Long, c As Long
    ReDim report(1 To matches.Count, 1 To COL_COUNT)
    For Each match In matches
        r = r + 1
        For c = 1 To COL_COUNT
            report(r, c) = match(c - 1)
        Next
    Next
    CollectReport = SortedWithoutDuplicates(report)
End Function

' Appends the matching appointments of one calendar to matches. Returns False
' and appends nothing if the calendar cannot be read, for example a shared
' calendar without permission.
Private Function AddMatches(ByVal matches As Collection, ByVal calendarFolder As Object, _
                            ByVal calendarName As String, ByVal fromDate As Date, ByVal untilDate As Date, _
                            ByVal customer As String, ByVal grouping As OfficiumGrouping) As Boolean
    On Error GoTo Failed
    Dim found As Outlook.Items, item As Object, own As New Collection, match As Variant
    Set found = AppointmentsBetween(calendarFolder, fromDate, untilDate)
    Set item = found.GetFirst
    Do While Not item Is Nothing
        If item.Class = olAppointment Then
            ' Restrict also returns appointments that started before the period;
            ' they belong to the previous one.
            If item.Start >= fromDate Then
                If CountsAsWork(item.AllDayEvent, item.MeetingStatus, item.ResponseStatus) Then
                    If MatchesCustomer(customer, item.Subject) Then
                        own.Add Array(item.Start, item.End, DateDiff("n", item.Start, item.End), _
                            item.Subject, ItemText(item, "Categories"), ItemText(item, "Body"), _
                            GroupKey(item.Start, grouping), _
                            calendarName)
                    End If
                End If
            End If
        End If
        Set item = found.GetNext
    Loop
    For Each match In own
        matches.Add match
    Next
    AddMatches = True
    Exit Function

Failed:
End Function

' A text property of an appointment, or "" if Outlook refuses it, as for a
' calendar shared with "free/busy, subject, location" only.
Private Function ItemText(ByVal item As Object, ByVal propertyName As String) As String
    On Error Resume Next
    ItemText = CallByName(item, propertyName, VbGet)
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

' Fills a new workbook in the given Excel and shows it. Sheet "Bericht" is
' meant for the customer: appointments by date with their full text, a total
' per group and overall, ready to print on A4. One sheet per calendar follows
' in the same layout. On an error Excel is closed again and False returned.
Private Function ShowInExcel(ByVal xl As Object, ByVal customer As String, ByVal fromDate As Date, _
                             ByVal untilDate As Date, ByVal report As Variant) As Boolean
    On Error GoTo Failed
    Dim book As Object, sheet As Object, takenNames As New Collection
    Dim names() As String, sums() As Long, calendarCount As Long, i As Long
    SetExcelStep "Mappe anlegen"
    calendarCount = SumByColumn(report, COL_CALENDAR, names, sums)
    Set book = xl.Workbooks.Add(-4167)  ' xlWBATWorksheet: a single sheet
    Set sheet = book.Worksheets(1)
    sheet.Name = SheetName("Bericht", takenNames)
    FillCustomerSheet sheet, customer, fromDate, untilDate, report, calendarCount > 1, ""

    For i = 1 To calendarCount
        SetExcelStep "Blatt " & names(i)
        Set sheet = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
        sheet.Name = SheetName(names(i), takenNames)
        FillCustomerSheet sheet, customer, fromDate, untilDate, RowsWhere(report, COL_CALENDAR, names(i)), _
            False, names(i)
    Next

    SetExcelStep "Mappe anzeigen"
    book.Worksheets(1).Activate
    CloseProgress
    xl.Visible = True
    HideGridlines xl
    ShowInExcel = True
    Exit Function

Failed:
    CloseProgress
    MsgBox "Excel-Fehler bei """ & excelStep & """: " & Err.Description & " (" & Err.Number & ")" & _
        vbCrLf & vbCrLf & _
        "Officium öffnet stattdessen die CSV-Datei.", vbExclamation, "Officium"
    ' Resume leaves the error handler; inside it On Error Resume Next has no effect.
    Resume CloseExcel
CloseExcel:
    On Error Resume Next
    xl.DisplayAlerts = False
    xl.Quit
End Function

' Remembers the step for the error message and shows it in the progress window.
Private Sub SetExcelStep(ByVal stepName As String)
    excelStep = stepName
    ShowProgress "Excel-Bericht: " & stepName
End Sub

' Shows the text in a small window in the middle of the screen, opening it if
' needed. DoEvents lets Windows paint it and keeps Outlook from being marked
' as not responding. Without the window the report still runs.
Private Sub ShowProgress(ByVal statusText As String)
    Const WS_POPUP As Long = &H80000000, WS_VISIBLE As Long = &H10000000, WS_DLGFRAME As Long = &H400000
    Const SS_CENTER As Long = 1, WS_EX_TOPMOST As Long = 8, WS_EX_TOOLWINDOW As Long = &H80
    Const WM_SETFONT As Long = &H30, DEFAULT_GUI_FONT As Long = 17
    Const WINDOW_WIDTH As Long = 460, WINDOW_HEIGHT As Long = 90
    On Error Resume Next
    If progressWindow = 0 Then
        progressWindow = CreateWindowExW(WS_EX_TOPMOST Or WS_EX_TOOLWINDOW, StrPtr("STATIC"), 0, _
            WS_POPUP Or WS_VISIBLE Or WS_DLGFRAME Or SS_CENTER, _
            (GetSystemMetrics(0) - WINDOW_WIDTH) \ 2, (GetSystemMetrics(1) - WINDOW_HEIGHT) \ 2, _
            WINDOW_WIDTH, WINDOW_HEIGHT, 0, 0, GetModuleHandleW(0), 0)
        If progressWindow = 0 Then Exit Sub
        SendMessageW progressWindow, WM_SETFONT, GetStockObject(DEFAULT_GUI_FONT), 0
    End If
    SetWindowTextW progressWindow, StrPtr(vbCrLf & "Officium erstellt den Bericht ..." & vbCrLf & vbCrLf & statusText)
    UpdateWindow progressWindow
    DoEvents
End Sub

' Closes the progress window. It stays on top, so it has to go before every
' message box.
Private Sub CloseProgress()
    If progressWindow = 0 Then Exit Sub
    DestroyWindow progressWindow
    progressWindow = 0
End Sub

' A sheet for the customer: title block in rows 1 to 4, the table from
' row 6. Per group a heading row and a subtotal, at the end the total.
' The column "Kalender" says whose calendar an appointment comes from; without
' showCalendar it is removed. A calendarName goes into the title block, for the
' sheet of one calendar. Errors go to the caller.
Private Sub FillCustomerSheet(ByVal sheet As Object, ByVal customer As String, ByVal fromDate As Date, _
                              ByVal untilDate As Date, ByVal report As Variant, ByVal showCalendar As Boolean, _
                              ByVal calendarName As String)
    Const HEADER_ROW As Long = 6
    Const LAST_COL As Long = 7
    Dim keys() As String, sums() As Long, groups As Long
    groups = SumByGroup(report, keys, sums)

    ' Rows of the table below the header: appointments, at most a heading and
    ' a subtotal per group, the total. Unused rows stay empty.
    Dim grid() As Variant, n As Long, r As Long, firstRow As Long, currentGroup As String
    Dim headingRows As New Collection, subtotalRows As New Collection
    ReDim grid(1 To UBound(report, 1) + 2 * groups + 1, 1 To LAST_COL)
    For r = 1 To UBound(report, 1)
        If r = 1 Or report(r, COL_GROUP) <> currentGroup Then
            If r > 1 Then AddSubtotal grid, n, HEADER_ROW, firstRow, currentGroup, subtotalRows
            currentGroup = report(r, COL_GROUP)
            If Len(currentGroup) > 0 Then
                n = n + 1
                grid(n, 1) = GroupLabel(currentGroup)
                headingRows.Add HEADER_ROW + n
            End If
            firstRow = HEADER_ROW + n + 1
        End If
        n = n + 1
        grid(n, 1) = DateSerial(Year(report(r, COL_START)), Month(report(r, COL_START)), Day(report(r, COL_START)))
        grid(n, 2) = TimeSerial(Hour(report(r, COL_START)), Minute(report(r, COL_START)), 0)
        grid(n, 3) = TimeSerial(Hour(report(r, COL_END)), Minute(report(r, COL_END)), 0)
        grid(n, 4) = report(r, COL_MINUTES) / 1440
        grid(n, 5) = report(r, COL_CALENDAR)
        grid(n, 6) = report(r, COL_SUBJECT)
        ' Excel breaks lines at LF only; a cell holds at most 32767 characters.
        grid(n, 7) = Left$(Replace(TrimText(report(r, COL_BODY)), vbCr, ""), 32000)
    Next
    AddSubtotal grid, n, HEADER_ROW, firstRow, currentGroup, subtotalRows
    n = n + 1
    grid(n, 1) = "Gesamt"

    Dim lastRow As Long, rowNumber As Variant, subtotal As Variant
    lastRow = HEADER_ROW + n
    SetExcelStep "Blatt Bericht, Formate"
    ' Formats before values: labels and texts stay text (Excel would read
    ' "Oktober 2026" as a date), durations keep counting past 24 hours.
    sheet.Columns(1).NumberFormat = "dd.mm.yyyy"
    sheet.Range("B:C").NumberFormat = "hh:mm"
    sheet.Columns(4).NumberFormat = "[h]:mm"
    sheet.Range("E:G").NumberFormat = "@"
    sheet.Range("A1:A4").NumberFormat = "@"
    For Each rowNumber In headingRows
        sheet.Cells(rowNumber, 1).NumberFormat = "@"
    Next
    For Each subtotal In subtotalRows
        sheet.Cells(subtotal(0), 1).NumberFormat = "@"
    Next
    sheet.Cells(lastRow, 1).NumberFormat = "@"

    SetExcelStep "Blatt Bericht, Werte"
    sheet.Range("A1").Value = "Leistungsnachweis"
    sheet.Range("A2").Value = "Kunde: " & CustomerLabel(customer)
    If Len(calendarName) > 0 Then sheet.Range("A2").Value = sheet.Range("A2").Value & ", Kalender: " & calendarName
    sheet.Range("A3").Value = "Zeitraum: " & DateText(fromDate) & " bis " & DateText(untilDate - 1)
    sheet.Range("A4").Value = "Erstellt am: " & DateText(Date)
    sheet.Range("A" & HEADER_ROW).Resize(1, LAST_COL).Value = _
        Array("Datum", "Beginn", "Ende", "Dauer", "Kalender", "Betreff", "Beschreibung")
    ' One assignment for all cells; cell by cell is slow with many appointments.
    sheet.Range("A" & (HEADER_ROW + 1)).Resize(n, LAST_COL).Value = grid
    ' Formula takes the English names. SUBTOTAL skips the subtotals above, so
    ' each appointment counts once in the total.
    For Each subtotal In subtotalRows
        sheet.Cells(subtotal(0), 4).Formula = "=SUBTOTAL(9,D" & subtotal(1) & ":D" & (subtotal(0) - 1) & ")"
    Next
    sheet.Cells(lastRow, 4).Formula = "=SUBTOTAL(9,D" & (HEADER_ROW + 1) & ":D" & (lastRow - 1) & ")"

    SetExcelStep "Blatt Bericht, Gestaltung"
    sheet.Range("A1").Font.Size = 16
    sheet.Range("A1").Font.Bold = True
    sheet.Columns(1).ColumnWidth = 11
    sheet.Range("B:C").ColumnWidth = 7
    sheet.Columns(4).ColumnWidth = 8
    sheet.Columns(5).ColumnWidth = 20
    sheet.Columns(6).ColumnWidth = 35
    sheet.Columns(7).ColumnWidth = 70
    With sheet.Range("A" & HEADER_ROW).Resize(1, LAST_COL)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With
    With sheet.Range("A" & (HEADER_ROW + 1)).Resize(n, LAST_COL)
        .VerticalAlignment = -4160  ' xlTop
        .Columns(5).WrapText = True
        .Columns(6).WrapText = True
        .Columns(7).WrapText = True
        .Borders(12).LineStyle = 1  ' xlInsideHorizontal, xlContinuous
        .Borders(12).Color = RGB(217, 217, 217)
    End With
    For Each rowNumber In headingRows
        sheet.Cells(rowNumber, 1).Resize(1, LAST_COL).Interior.Color = RGB(221, 235, 247)
        sheet.Cells(rowNumber, 1).Font.Bold = True
    Next
    For Each subtotal In subtotalRows
        sheet.Cells(subtotal(0), 1).Resize(1, LAST_COL).Font.Bold = True
    Next
    With sheet.Cells(lastRow, 1).Resize(1, LAST_COL)
        .Font.Bold = True
        .Borders(8).LineStyle = 1        ' xlEdgeTop
        .Borders(9).LineStyle = -4119    ' xlEdgeBottom, xlDouble
    End With
    Dim lastColumn As String
    lastColumn = "G"
    If Not showCalendar Then
        sheet.Columns(5).Delete
        lastColumn = "F"
    End If
    sheet.Range("A" & (HEADER_ROW + 1)).Resize(n, LAST_COL).Rows.AutoFit
    SetExcelStep "Blatt Bericht, Seite einrichten"
    SetUpPage sheet, "$" & HEADER_ROW & ":$" & HEADER_ROW, "A1:" & lastColumn & lastRow
End Sub

' Closes a named group in the customer sheet with a row "Summe <group>".
' subtotalRows gets Array(row, first row of the group) for the formula.
Private Sub AddSubtotal(ByRef grid() As Variant, ByRef n As Long, ByVal headerRow As Long, _
                        ByVal firstRow As Long, ByVal groupKeyText As String, ByVal subtotalRows As Collection)
    If Len(groupKeyText) = 0 Then Exit Sub
    n = n + 1
    grid(n, 1) = "Summe " & GroupLabel(groupKeyText)
    subtotalRows.Add Array(headerRow + n, firstRow)
End Sub

' A4 landscape, as wide as one page, header row on every page, page numbers.
' Skipped without a printer, where Excel cannot read the page setup.
' PrintCommunication off sends all settings to the printer driver at once
' instead of one call per property, which takes seconds.
Private Sub SetUpPage(ByVal sheet As Object, ByVal titleRows As String, ByVal area As String)
    On Error Resume Next
    sheet.Application.PrintCommunication = False
    With sheet.PageSetup
        .PrintArea = area
        .PrintTitleRows = titleRows
        .Orientation = 2        ' xlLandscape
        .PaperSize = 9          ' xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterFooter = "Seite &P von &N"
    End With
    sheet.Application.PrintCommunication = True
End Sub

' Gridlines belong to the window, so each sheet is activated in turn; only
' after the workbook is shown. The first sheet stays active.
Private Sub HideGridlines(ByVal xl As Object)
    Dim ws As Object
    On Error Resume Next
    For Each ws In xl.ActiveWorkbook.Worksheets
        ws.Activate
        xl.ActiveWindow.DisplayGridlines = False
    Next
    xl.ActiveWorkbook.Worksheets(1).Activate
End Sub

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
    answer = InputBox("Kunde: Name oder Stichwort aus dem Betreff." & vbCrLf & vbCrLf & _
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

    Dim calendars As Collection, entry As Variant, found As Long, total As Long, skipped As String
    Set calendars = AllCalendars(skipped)
    If Len(skipped) > 0 Then Debug.Print "Nicht lesbar: " & skipped
    Debug.Print "--- Officium: " & Format(fromDate, "Short Date") & " bis " & Format(toDate, "Short Date") & " ---"
    For Each entry In calendars
        Debug.Print "== " & entry(1)
        found = PrintAppointments(entry(0), fromDate, toDate + 1)
        If found < 0 Then
            Debug.Print "nicht lesbar"
        Else
            Debug.Print found & " Termine"
            total = total + found
        End If
    Next

    MsgBox total & " Termine in " & calendars.Count & " Kalendern gefunden. Die Liste steht im " & _
        "Direktfenster (Strg+G im VBA-Editor).", vbInformation, "Officium"
End Sub

' Prints the appointments of one calendar that overlap [fromTime, toTime) and
' returns their number, or -1 if the calendar cannot be read.
Private Function PrintAppointments(ByVal calendarFolder As Object, ByVal fromTime As Date, _
                                   ByVal toTime As Date) As Long
    On Error GoTo Failed
    Dim found As Outlook.Items, item As Object, total As Long
    Set found = AppointmentsBetween(calendarFolder, fromTime, toTime)
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
    PrintAppointments = total
    Exit Function

Failed:
    PrintAppointments = -1
End Function

' Appointments of a calendar that overlap [fromTime, toTime),
' with recurring series expanded into single occurrences.
' Count is meaningless on the result while recurrences are included:
' iterate with GetFirst/GetNext.
Private Function AppointmentsBetween(ByVal calendarFolder As Object, ByVal fromTime As Date, _
                                     ByVal toTime As Date) As Outlook.Items
    Dim all As Outlook.Items
    Set all = calendarFolder.Items
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
