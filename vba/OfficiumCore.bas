Attribute VB_Name = "OfficiumCore"
' Report logic without Outlook objects, covered by OfficiumSelfTest.
Option Explicit

Public Enum OfficiumPeriod
    opThisMonth = 1
    opLastMonth = 2
    opThisQuarter = 3
    opLastQuarter = 4
End Enum

' Outlook constants as numbers, so this module needs no Outlook objects.
Private Const MEETING_CANCELED As Long = 5           ' olMeetingCanceled
Private Const MEETING_RECEIVED_CANCELED As Long = 7  ' olMeetingReceivedAndCanceled
Private Const RESPONSE_DECLINED As Long = 4          ' olResponseDeclined

' Numbers as offered in the grouping prompt.
Public Enum OfficiumGrouping
    ogMonth = 1
    ogQuarter = 2
    ogNone = 3
End Enum

' Columns of a report: a 2D array (1 To appointments, 1 To COL_COUNT) shared
' by the summary and every export. Empty if no appointment matched.
Public Const COL_START As Long = 1       ' Date
Public Const COL_END As Long = 2         ' Date
Public Const COL_MINUTES As Long = 3     ' Long
Public Const COL_SUBJECT As Long = 4     ' String
Public Const COL_CATEGORIES As Long = 5  ' String
Public Const COL_BODY As Long = 6        ' String
Public Const COL_GROUP As Long = 7       ' String, see GroupKey
Public Const COL_CALENDAR As Long = 8    ' String, name of the calendar
Public Const COL_COUNT As Long = 8

' Half-open range [fromDate, untilDate) of a preset period relative to today.
Public Sub PeriodRange(ByVal period As OfficiumPeriod, ByVal today As Date, _
                       ByRef fromDate As Date, ByRef untilDate As Date)
    Dim firstOfMonth As Date, firstOfQuarter As Date
    firstOfMonth = DateSerial(Year(today), Month(today), 1)
    firstOfQuarter = DateSerial(Year(today), ((Month(today) - 1) \ 3) * 3 + 1, 1)

    Select Case period
    Case opThisMonth
        fromDate = firstOfMonth
        untilDate = DateAdd("m", 1, firstOfMonth)
    Case opLastMonth
        fromDate = DateAdd("m", -1, firstOfMonth)
        untilDate = firstOfMonth
    Case opThisQuarter
        fromDate = firstOfQuarter
        untilDate = DateAdd("m", 3, firstOfQuarter)
    Case opLastQuarter
        fromDate = DateAdd("m", -3, firstOfQuarter)
        untilDate = firstOfQuarter
    Case Else
        Err.Raise 5, "PeriodRange", "Unknown period " & period
    End Select
End Sub

' True if the subject contains the customer, ignoring case. Categories do not
' count. An empty customer matches every appointment.
Public Function MatchesCustomer(ByVal customer As String, ByVal subjectText As String) As Boolean
    customer = Trim$(customer)
    MatchesCustomer = (Len(customer) = 0) Or (InStr(1, subjectText, customer, vbTextCompare) > 0)
End Function

' False for appointments that are no time spent: all-day events, cancelled
' meetings and meetings the user declined. Free and tentative ones count.
Public Function CountsAsWork(ByVal isAllDay As Boolean, ByVal meetingState As Long, _
                             ByVal response As Long) As Boolean
    If isAllDay Then Exit Function
    If meetingState = MEETING_CANCELED Or meetingState = MEETING_RECEIVED_CANCELED Then Exit Function
    If response = RESPONSE_DECLINED Then Exit Function
    CountsAsWork = True
End Function

' Group of an appointment start, sortable as text: "2026-10" by month,
' "2026-Q4" by quarter, "" without grouping.
Public Function GroupKey(ByVal startTime As Date, ByVal grouping As OfficiumGrouping) As String
    Select Case grouping
    Case ogMonth
        GroupKey = Year(startTime) & "-" & Format(Month(startTime), "00")
    Case ogQuarter
        GroupKey = Year(startTime) & "-Q" & ((Month(startTime) - 1) \ 3 + 1)
    Case ogNone
        GroupKey = ""
    Case Else
        Err.Raise 5, "GroupKey", "Unknown grouping " & grouping
    End Select
End Function

' Readable name of a group key: "2026-10" -> "Oktober 2026", "2026-Q4" ->
' "4. Quartal 2026", "" stays empty. Month names are fixed, not from Windows.
Public Function GroupLabel(ByVal key As String) As String
    If Len(key) = 0 Then Exit Function
    If Mid$(key, 6, 1) = "Q" Then
        GroupLabel = Mid$(key, 7) & ". Quartal " & Left$(key, 4)
    Else
        GroupLabel = Split("Januar Februar März April Mai Juni Juli August September Oktober November Dezember")( _
            CLng(Mid$(key, 6, 2)) - 1) & " " & Left$(key, 4)
    End If
End Function

' Minutes as hours:minutes that keeps counting past 24 hours, e.g. 2250 -> "37:30".
Public Function FormatMinutes(ByVal minutes As Long) As String
    FormatMinutes = minutes \ 60 & ":" & Format(minutes Mod 60, "00")
End Function

' Sums the minutes of a report per value of one column, e.g. COL_CALENDAR.
' Fills keys and sums (1-based, values in order of first appearance) and
' returns the number of values.
Public Function SumByColumn(ByVal report As Variant, ByVal keyColumn As Long, _
                            ByRef keys() As String, ByRef sums() As Long) As Long
    If IsEmpty(report) Then Exit Function
    Dim n As Long, r As Long, g As Long
    ReDim keys(1 To UBound(report, 1))
    ReDim sums(1 To UBound(report, 1))
    For r = 1 To UBound(report, 1)
        g = 1
        Do While g <= n
            If keys(g) = report(r, keyColumn) Then Exit Do
            g = g + 1
        Loop
        If g > n Then
            n = g
            keys(n) = report(r, keyColumn)
        End If
        sums(g) = sums(g) + report(r, COL_MINUTES)
    Next
    ReDim Preserve keys(1 To n)
    ReDim Preserve sums(1 To n)
    SumByColumn = n
End Function

' Sums the minutes of a report per group, see SumByColumn.
Public Function SumByGroup(ByVal report As Variant, ByRef keys() As String, ByRef sums() As Long) As Long
    SumByGroup = SumByColumn(report, COL_GROUP, keys, sums)
End Function

' The rows of a report whose column holds the given value, as a report of
' their own. Empty if there are none.
Public Function RowsWhere(ByVal report As Variant, ByVal keyColumn As Long, ByVal wanted As String) As Variant
    If IsEmpty(report) Then Exit Function
    Dim n As Long, r As Long, c As Long, part() As Variant
    For r = 1 To UBound(report, 1)
        If report(r, keyColumn) = wanted Then n = n + 1
    Next
    If n = 0 Then Exit Function
    ReDim part(1 To n, 1 To COL_COUNT)
    n = 0
    For r = 1 To UBound(report, 1)
        If report(r, keyColumn) = wanted Then
            n = n + 1
            For c = 1 To COL_COUNT
                part(n, c) = report(r, c)
            Next
        End If
    Next
    RowsWhere = part
End Function

' The report sorted by start, without rows that repeat an earlier row's start,
' end and subject: a meeting in the own and a shared calendar counts once,
' for the calendar listed first. Rows with the same start keep their order.
Public Function SortedWithoutDuplicates(ByVal report As Variant) As Variant
    If IsEmpty(report) Then Exit Function
    Dim order() As Long, n As Long, r As Long, i As Long, j As Long, moving As Long
    ReDim order(1 To UBound(report, 1))
    ' Insertion sort over row numbers; reports hold at most a few hundred rows.
    For r = 1 To UBound(report, 1)
        i = n
        Do While i >= 1
            If report(order(i), COL_START) <= report(r, COL_START) Then Exit Do
            i = i - 1
        Loop
        For j = n To i + 1 Step -1
            order(j + 1) = order(j)
        Next
        order(i + 1) = r
        n = n + 1
    Next

    Dim kept() As Long, keptCount As Long, isDuplicate As Boolean
    ReDim kept(1 To n)
    For i = 1 To n
        moving = order(i)
        isDuplicate = False
        ' Duplicates share the start, so they sit right before this row.
        j = keptCount
        Do While j >= 1
            If report(kept(j), COL_START) <> report(moving, COL_START) Then Exit Do
            If report(kept(j), COL_END) = report(moving, COL_END) Then
                If report(kept(j), COL_SUBJECT) = report(moving, COL_SUBJECT) Then isDuplicate = True
            End If
            j = j - 1
        Loop
        If Not isDuplicate Then
            keptCount = keptCount + 1
            kept(keptCount) = moving
        End If
    Next

    Dim result() As Variant, c As Long
    ReDim result(1 To keptCount, 1 To COL_COUNT)
    For i = 1 To keptCount
        For c = 1 To COL_COUNT
            result(i, c) = report(kept(i), c)
        Next
    Next
    SortedWithoutDuplicates = result
End Function

' Excel sheet name for a calendar: the characters Excel forbids become "-",
' at most 31 characters, and " (2)", " (3)" ... if the name is in takenNames
' already (ignoring case). The result is added to takenNames.
Public Function SheetName(ByVal calendarName As String, ByVal takenNames As Collection) As String
    Dim forbidden As Variant, baseName As String, candidate As String, n As Long
    baseName = calendarName
    For Each forbidden In Array(":", "\", "/", "?", "*", "[", "]", "'")
        baseName = Replace(baseName, forbidden, "-")
    Next
    baseName = Trim$(baseName)
    If Len(baseName) = 0 Then baseName = "Kalender"
    candidate = Left$(baseName, 31)
    n = 1
    Do While IsTaken(takenNames, candidate)
        n = n + 1
        candidate = Left$(baseName, 31 - Len(CStr(n)) - 3) & " (" & n & ")"
    Loop
    takenNames.Add candidate, candidate
    SheetName = candidate
End Function

' Collection keys are compared ignoring case, like Excel sheet names.
Private Function IsTaken(ByVal takenNames As Collection, ByVal candidate As String) As Boolean
    Dim probe As Variant
    On Error Resume Next
    probe = takenNames.Item(candidate)
    IsTaken = (Err.Number = 0)
End Function

' One line per group and a grand total with the number of appointments,
' separated by tabs, e.g. "2026-10<Tab>37:30 h".
Public Function TotalsText(ByVal report As Variant) As String
    Dim keys() As String, sums() As Long, n As Long, g As Long, total As Long, lines As String
    n = SumByGroup(report, keys, sums)
    For g = 1 To n
        total = total + sums(g)
        If Len(keys(g)) > 0 Then lines = lines & keys(g) & vbTab & FormatMinutes(sums(g)) & " h" & vbCrLf
    Next
    Dim appointments As Long
    If Not IsEmpty(report) Then appointments = UBound(report, 1)
    TotalsText = lines & "Gesamt" & vbTab & FormatMinutes(total) & " h (" & appointments _
        & IIf(appointments = 1, " Termin)", " Termine)")
End Function

' Date as dd.mm.yyyy and time as hh:mm, independent of the Windows formats.
Public Function DateText(ByVal d As Date) As String
    DateText = Format(Day(d), "00") & "." & Format(Month(d), "00") & "." & Year(d)
End Function

Public Function TimeText(ByVal d As Date) As String
    TimeText = Format(Hour(d), "00") & ":" & Format(Minute(d), "00")
End Function

' Heading of a report for [fromDate, untilDate), e.g. "Acme, 01.10.2026 bis 31.10.2026".
Public Function ReportTitle(ByVal customer As String, ByVal fromDate As Date, ByVal untilDate As Date) As String
    ReportTitle = CustomerLabel(customer) & ", " & DateText(fromDate) & " bis " & DateText(untilDate - 1)
End Function

' File name of the CSV for a report, e.g. "Officium Acme 2026-10-01 bis 2026-10-31.csv".
Public Function CsvFileName(ByVal customer As String, ByVal fromDate As Date, ByVal untilDate As Date) As String
    CsvFileName = "Officium " & SafeFileName(CustomerLabel(customer)) & " " _
        & Format(fromDate, "yyyy-mm-dd") & " bis " & Format(untilDate - 1, "yyyy-mm-dd") & ".csv"
End Function

' The customer as shown in titles: "Alle Termine" if empty.
Public Function CustomerLabel(ByVal customer As String) As String
    CustomerLabel = Trim$(customer)
    If Len(CustomerLabel) = 0 Then CustomerLabel = "Alle Termine"
End Function

' Text usable in a Windows file name: \ / : * ? " < > | become "-".
Public Function SafeFileName(ByVal fileText As String) As String
    Dim forbidden As Variant
    For Each forbidden In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        fileText = Replace(fileText, forbidden, "-")
    Next
    SafeFileName = Trim$(fileText)
End Function

' Text without leading and trailing spaces, tabs and line breaks, which
' appointment bodies usually end with.
Public Function TrimText(ByVal fullText As String) As String
    Dim first As Long, last As Long
    first = 1
    last = Len(fullText)
    Do While first <= last
        If Not IsBlank(Mid$(fullText, first, 1)) Then Exit Do
        first = first + 1
    Loop
    Do While last >= first
        If Not IsBlank(Mid$(fullText, last, 1)) Then Exit Do
        last = last - 1
    Loop
    TrimText = Mid$(fullText, first, last - first + 1)
End Function

Private Function IsBlank(ByVal character As String) As Boolean
    IsBlank = InStr(" " & vbTab & vbCr & vbLf, character) > 0
End Function

' One field of a CSV for German Excel: quoted if it contains ";", a quote or a
' line break, with quotes doubled. Excel reads text starting with = + - @ as a
' formula even when quoted, so such text gets a leading space.
Public Function CsvField(ByVal fieldText As String) As String
    If fieldText Like "[=+@-]*" Then fieldText = " " & fieldText
    If InStr(fieldText, ";") > 0 Or InStr(fieldText, """") > 0 _
       Or InStr(fieldText, vbCr) > 0 Or InStr(fieldText, vbLf) > 0 Then
        CsvField = """" & Replace(fieldText, """", """""") & """"
    Else
        CsvField = fieldText
    End If
End Function

' Report as semicolon-separated CSV: a header, one line per appointment with
' the full body, then the totals per group and overall.
Public Function CsvText(ByVal report As Variant) As String
    Dim lines As String, r As Long
    lines = "Datum;Beginn;Ende;Dauer;Minuten;Gruppe;Kalender;Betreff;Kategorien;Text" & vbCrLf
    For r = 1 To UBound(report, 1)
        lines = lines & DateText(report(r, COL_START)) & ";" & TimeText(report(r, COL_START)) & ";" _
            & TimeText(report(r, COL_END)) & ";" & FormatMinutes(report(r, COL_MINUTES)) & ";" _
            & report(r, COL_MINUTES) & ";" & CsvField(report(r, COL_GROUP)) & ";" _
            & CsvField(report(r, COL_CALENDAR)) & ";" & CsvField(report(r, COL_SUBJECT)) & ";" & CsvField(report(r, COL_CATEGORIES)) & ";" _
            & CsvField(TrimText(report(r, COL_BODY))) & vbCrLf
    Next

    Dim keys() As String, sums() As Long, n As Long, g As Long, total As Long
    n = SumByGroup(report, keys, sums)
    lines = lines & vbCrLf & "Summe;Dauer;Minuten" & vbCrLf
    For g = 1 To n
        total = total + sums(g)
        If Len(keys(g)) > 0 Then lines = lines & CsvField(keys(g)) & ";" & FormatMinutes(sums(g)) & ";" & sums(g) & vbCrLf
    Next
    CsvText = lines & "Gesamt;" & FormatMinutes(total) & ";" & total & vbCrLf
End Function

' First line of a text that is not blank, e.g. the opening line of a body.
Public Function FirstLine(ByVal fullText As String) As String
    Dim lines() As String, i As Long
    lines = Split(Replace(fullText, vbCr, vbLf), vbLf)
    For i = 0 To UBound(lines)
        If Len(TrimText(lines(i))) > 0 Then
            FirstLine = TrimText(lines(i))
            Exit Function
        End If
    Next
End Function

' Plain-text report to paste into a mail: title, totals, then one line per
' appointment with the first line of its body.
Public Function ClipboardText(ByVal title As String, ByVal report As Variant) As String
    Dim lines As String, r As Long, note As String
    lines = title & vbCrLf & vbCrLf & TotalsText(report) & vbCrLf & vbCrLf
    For r = 1 To UBound(report, 1)
        lines = lines & DateText(report(r, COL_START)) & " " & TimeText(report(r, COL_START)) _
            & "-" & TimeText(report(r, COL_END)) & vbTab & FormatMinutes(report(r, COL_MINUTES)) & " h" _
            & vbTab & report(r, COL_SUBJECT)
        note = FirstLine(report(r, COL_BODY))
        If Len(note) > 0 Then lines = lines & " - " & note
        lines = lines & vbCrLf
    Next
    ClipboardText = lines
End Function
