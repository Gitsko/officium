Attribute VB_Name = "OfficiumTests"
' Self-test for OfficiumCore. Run OfficiumSelfTest; failures are listed in the
' Immediate window (Ctrl+G).
Option Explicit

Private passed As Long
Private failed As Long

Public Sub OfficiumSelfTest()
    passed = 0
    failed = 0

    TestPeriodRange
    TestMatchesCustomer
    TestCountsAsWork
    TestGrouping
    TestCalendars
    TestReportTitle
    TestCsv
    TestClipboardText

    Dim summary As String
    summary = passed & " bestanden, " & failed & " fehlgeschlagen"
    Debug.Print "--- OfficiumSelfTest: " & summary & " ---"
    If failed > 0 Then
        MsgBox summary & vbCrLf & "Details im Direktfenster (Strg+G im VBA-Editor).", vbExclamation, "Officium-Selbsttest"
    Else
        MsgBox summary, vbInformation, "Officium-Selbsttest"
    End If
End Sub

' Date literals in VBA code are always month/day/year, whatever the locale.
Private Sub TestPeriodRange()
    Dim f As Date, u As Date

    PeriodRange opThisMonth, #10/19/2026#, f, u
    AssertEqual "this month: from", #10/1/2026#, f
    AssertEqual "this month: until", #11/1/2026#, u

    PeriodRange opThisMonth, #12/31/2026#, f, u
    AssertEqual "this month in December: until", #1/1/2027#, u

    PeriodRange opLastMonth, #1/15/2027#, f, u
    AssertEqual "last month across new year: from", #12/1/2026#, f
    AssertEqual "last month across new year: until", #1/1/2027#, u

    PeriodRange opThisQuarter, #10/19/2026#, f, u
    AssertEqual "this quarter: from", #10/1/2026#, f
    AssertEqual "this quarter: until", #1/1/2027#, u

    PeriodRange opThisQuarter, #3/31/2027#, f, u
    AssertEqual "this quarter on its last day: from", #1/1/2027#, f
    AssertEqual "this quarter on its last day: until", #4/1/2027#, u

    PeriodRange opLastQuarter, #2/28/2027#, f, u
    AssertEqual "last quarter across new year: from", #10/1/2026#, f
    AssertEqual "last quarter across new year: until", #1/1/2027#, u
End Sub

Private Sub TestMatchesCustomer()
    AssertEqual "customer: subject contains it, ignoring case", True, MatchesCustomer("Acme", "Call mit ACME GmbH")
    AssertEqual "customer: input trimmed", True, MatchesCustomer(" Acme ", "Acme: Workshop")
    AssertEqual "customer: no match", False, MatchesCustomer("Acme", "Workshop")
    AssertEqual "customer: empty input matches all", True, MatchesCustomer(" ", "Workshop")
End Sub

' Meeting status 1 = meeting, 3 = received; response 2 = tentative, 3 = accepted.
Private Sub TestCountsAsWork()
    AssertEqual "counts: appointment", True, CountsAsWork(False, 0, 0)
    AssertEqual "counts: accepted meeting", True, CountsAsWork(False, 3, 3)
    AssertEqual "counts: tentative meeting", True, CountsAsWork(False, 3, 2)
    AssertEqual "counts: not all-day", False, CountsAsWork(True, 0, 0)
    AssertEqual "counts: not cancelled", False, CountsAsWork(False, 5, 0)
    AssertEqual "counts: not cancelled by organizer", False, CountsAsWork(False, 7, 3)
    AssertEqual "counts: not declined", False, CountsAsWork(False, 3, 4)
End Sub

Private Sub TestGrouping()
    AssertEqual "group: month", "2026-10", GroupKey(#10/19/2026 9:00:00 AM#, ogMonth)
    AssertEqual "group: month below ten", "2027-01", GroupKey(#1/4/2027#, ogMonth)
    AssertEqual "group: quarter on new year's eve", "2026-Q4", GroupKey(#12/31/2026 11:30:00 PM#, ogQuarter)
    AssertEqual "group: quarter after new year", "2027-Q1", GroupKey(#1/1/2027#, ogQuarter)
    AssertEqual "group: none", "", GroupKey(#10/19/2026#, ogNone)
    AssertEqual "group label: month", "März 2027", GroupLabel("2027-03")
    AssertEqual "group label: quarter", "4. Quartal 2026", GroupLabel("2026-Q4")
    AssertEqual "group label: none", "", GroupLabel("")

    AssertEqual "minutes: below one hour", "0:45", FormatMinutes(45)
    AssertEqual "minutes: above 24 hours", "37:30", FormatMinutes(2250)

    Dim keys() As String, sums() As Long
    AssertEqual "sum: groups", 2, SumByGroup(ReportOf("2026-Q4", 90, "2027-Q1", 60, "2026-Q4", 30), keys, sums)
    AssertEqual "sum: first group", "2026-Q4", keys(1)
    AssertEqual "sum: first total", 120, sums(1)
    AssertEqual "sum: second total", 60, sums(2)
    AssertEqual "sum: empty report", 0, SumByGroup(Empty, keys, sums)

    AssertEqual "totals: per group and overall", _
        "2026-10" & vbTab & "37:30 h" & vbCrLf & "2026-11" & vbTab & "1:00 h" & vbCrLf & _
        "Gesamt" & vbTab & "38:30 h (2 Termine)", _
        TotalsText(ReportOf("2026-10", 2250, "2026-11", 60))
    AssertEqual "totals: without grouping", "Gesamt" & vbTab & "0:45 h (1 Termin)", TotalsText(ReportOf("", 45))
End Sub

Private Sub TestCalendars()
    Dim names() As String, sums() As Long, part As Variant, taken As New Collection
    AssertEqual "calendars: count", 2, SumByColumn(SampleReport(), COL_CALENDAR, names, sums)
    AssertEqual "calendars: first", "Kalender", names(1)
    AssertEqual "calendars: second total", 45, sums(2)

    part = RowsWhere(SampleReport(), COL_CALENDAR, "Kalender - Max")
    AssertEqual "rows where: one row", 1, UBound(part, 1)
    AssertEqual "rows where: keeps the columns", "Call", part(1, COL_SUBJECT)
    AssertEqual "rows where: none", True, IsEmpty(RowsWhere(SampleReport(), COL_CALENDAR, "Fremd"))

    AssertEqual "sheet name: free", "Summen", SheetName("Summen", taken)
    AssertEqual "sheet name: forbidden characters", "Kalender- Team-Projekt", SheetName("Kalender: Team/Projekt", taken)
    AssertEqual "sheet name: taken ignoring case", "summen (2)", SheetName("summen", taken)
    AssertEqual "sheet name: at most 31 characters", "Kalender - Maximilian Musterman", _
        SheetName("Kalender - Maximilian Mustermann", taken)
    AssertEqual "sheet name: shortened when taken", "Kalender - Maximilian Muste (2)", _
        SheetName("Kalender - Maximilian Mustermann", taken)
    AssertEqual "sheet name: blank", "Kalender", SheetName(" ", taken)

    ' Rows: the November call of SampleReport, the same call from a shared
    ' calendar, then the October workshop.
    Dim mixed() As Variant, sample As Variant, c As Long, merged As Variant
    sample = SampleReport()
    ReDim mixed(1 To 3, 1 To COL_COUNT)
    For c = 1 To COL_COUNT
        mixed(1, c) = sample(2, c)
        mixed(2, c) = sample(2, c)
        mixed(3, c) = sample(1, c)
    Next
    mixed(2, COL_CALENDAR) = "Kalender - Team"
    merged = SortedWithoutDuplicates(mixed)
    AssertEqual "merged: duplicate dropped", 2, UBound(merged, 1)
    AssertEqual "merged: sorted by start", "Workshop; Planung", merged(1, COL_SUBJECT)
    AssertEqual "merged: first calendar kept", "Kalender - Max", merged(2, COL_CALENDAR)
End Sub

Private Sub TestReportTitle()
    AssertEqual "date text", "05.01.2027", DateText(#1/5/2027 8:15:00 AM#)
    AssertEqual "time text", "08:05", TimeText(#1/5/2027 8:05:00 AM#)
    AssertEqual "title: customer", "Acme, 01.10.2026 bis 31.10.2026", ReportTitle(" Acme ", #10/1/2026#, #11/1/2026#)
    AssertEqual "title: all appointments", "Alle Termine, 01.10.2026 bis 31.12.2026", ReportTitle("", #10/1/2026#, #1/1/2027#)
End Sub

Private Sub TestCsv()
    AssertEqual "csv field: plain", "Acme, Intern", CsvField("Acme, Intern")
    AssertEqual "csv field: separator", Quoted("'a;b'"), CsvField("a;b")
    AssertEqual "csv field: quotes doubled", Quoted("'Punkt ''A'''"), CsvField(Quoted("Punkt 'A'"))
    AssertEqual "csv field: line break", Quoted("'a" & vbCrLf & "b'"), CsvField("a" & vbCrLf & "b")
    AssertEqual "csv field: formula-like text", " - Punkt", CsvField("- Punkt")
    AssertEqual "trim text", "a" & vbCrLf & "b", TrimText(vbTab & " a" & vbCrLf & "b " & vbCrLf & vbCrLf)
    AssertEqual "trim text: only blanks", "", TrimText(" " & vbCrLf)
    AssertEqual "csv file name", "Officium A-B- C- 2026-10-01 bis 2026-10-31.csv", _
        CsvFileName("A/B: C?", #10/1/2026#, #11/1/2026#)

    Dim csv As String, totals As String
    csv = CsvText(SampleReport())
    AssertEqual "csv: header first", 1, InStr(csv, "Datum;Beginn;Ende;Dauer;Minuten;Gruppe;Kalender;Betreff;Kategorien;Text" & vbCrLf)
    AssertEqual "csv: row with quoted fields", True, InStr(csv, vbCrLf & _
        "06.10.2026;09:00;10:30;1:30;90;2026-10;Kalender;" & Quoted("'Workshop; Planung'") & ";Acme;" & _
        Quoted("'Agenda" & vbCrLf & "Punkt ''A'''") & vbCrLf) > 0
    AssertEqual "csv: plain row", True, InStr(csv, vbCrLf & _
        "02.11.2026;14:00;14:45;0:45;45;2026-11;Kalender - Max;Call;Acme, Intern;" & vbCrLf) > 0
    totals = vbCrLf & vbCrLf & "Summe;Dauer;Minuten" & vbCrLf & "2026-10;1:30;90" & vbCrLf & _
        "2026-11;0:45;45" & vbCrLf & "Gesamt;2:15;135" & vbCrLf
    AssertEqual "csv: totals at the end", totals, Right$(csv, Len(totals))
End Sub

Private Sub TestClipboardText()
    AssertEqual "first line: skips blank lines", "Agenda", FirstLine(vbCrLf & "  " & vbCrLf & " Agenda " & vbCrLf & "Punkt")
    AssertEqual "first line: empty body", "", FirstLine("")

    AssertEqual "clipboard text", _
        "Acme, 01.10.2026 bis 30.11.2026" & vbCrLf & vbCrLf & _
        "2026-10" & vbTab & "1:30 h" & vbCrLf & _
        "2026-11" & vbTab & "0:45 h" & vbCrLf & _
        "Gesamt" & vbTab & "2:15 h (2 Termine)" & vbCrLf & vbCrLf & _
        "06.10.2026 09:00-10:30" & vbTab & "1:30 h" & vbTab & "Workshop; Planung - Agenda" & vbCrLf & _
        "02.11.2026 14:00-14:45" & vbTab & "0:45 h" & vbTab & "Call" & vbCrLf, _
        ClipboardText("Acme, 01.10.2026 bis 30.11.2026", SampleReport())
End Sub

' Text with ' turned into ", to keep expected CSV values readable.
Private Function Quoted(ByVal fieldText As String) As String
    Quoted = Replace(fieldText, "'", """")
End Function

' Two appointments in October and November from two calendars, with every
' column filled.
Private Function SampleReport() As Variant
    Dim report() As Variant
    ReDim report(1 To 2, 1 To COL_COUNT)
    report(1, COL_START) = #10/6/2026 9:00:00 AM#
    report(1, COL_END) = #10/6/2026 10:30:00 AM#
    report(1, COL_MINUTES) = 90
    report(1, COL_SUBJECT) = "Workshop; Planung"
    report(1, COL_CATEGORIES) = "Acme"
    report(1, COL_BODY) = "Agenda" & vbCrLf & "Punkt ""A""" & vbCrLf & vbCrLf
    report(1, COL_GROUP) = "2026-10"
    report(1, COL_CALENDAR) = "Kalender"
    report(2, COL_START) = #11/2/2026 2:00:00 PM#
    report(2, COL_END) = #11/2/2026 2:45:00 PM#
    report(2, COL_MINUTES) = 45
    report(2, COL_SUBJECT) = "Call"
    report(2, COL_CATEGORIES) = "Acme, Intern"
    report(2, COL_BODY) = ""
    report(2, COL_GROUP) = "2026-11"
    report(2, COL_CALENDAR) = "Kalender - Max"
    SampleReport = report
End Function

' Report with only the group and minutes columns filled, from pairs of values.
Private Function ReportOf(ParamArray groupAndMinutes() As Variant) As Variant
    Dim report() As Variant, r As Long
    ReDim report(1 To (UBound(groupAndMinutes) + 1) \ 2, 1 To COL_COUNT)
    For r = 1 To UBound(report, 1)
        report(r, COL_GROUP) = groupAndMinutes(2 * r - 2)
        report(r, COL_MINUTES) = groupAndMinutes(2 * r - 1)
    Next
    ReportOf = report
End Function

Private Sub AssertEqual(ByVal testName As String, ByVal expected As Variant, ByVal actual As Variant)
    If expected = actual Then
        passed = passed + 1
    Else
        failed = failed + 1
        Debug.Print "FEHLER " & testName & ": erwartet " & expected & ", erhalten " & actual
    End If
End Sub
