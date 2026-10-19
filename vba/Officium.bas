Attribute VB_Name = "Officium"
' Officium - calendar reports for classic Outlook.
Option Explicit

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
    MsgBox title & vbCrLf & vbCrLf & TotalsText(report), vbInformation, "Officium"
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
