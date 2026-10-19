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
    TestReportTitle

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
    AssertEqual "customer: category", True, MatchesCustomer("Acme", "Workshop", "Acme")
    AssertEqual "customer: category ignores case, ';' list", True, MatchesCustomer("acme", "Workshop", "Intern; ACME")
    AssertEqual "customer: category in ',' list", True, MatchesCustomer("Acme", "Workshop", "Intern, Acme")
    AssertEqual "customer: category must match whole", False, MatchesCustomer("Acme", "Workshop", "Acme Corp")
    AssertEqual "customer: subject contains it", True, MatchesCustomer("Acme", "Call mit ACME GmbH", "")
    AssertEqual "customer: no match", False, MatchesCustomer("Acme", "Workshop", "Intern")
    AssertEqual "customer: empty input matches all", True, MatchesCustomer(" ", "Workshop", "")
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

Private Sub TestReportTitle()
    AssertEqual "date text", "05.01.2027", DateText(#1/5/2027 8:15:00 AM#)
    AssertEqual "time text", "08:05", TimeText(#1/5/2027 8:05:00 AM#)
    AssertEqual "title: customer", "Acme, 01.10.2026 bis 31.10.2026", ReportTitle(" Acme ", #10/1/2026#, #11/1/2026#)
    AssertEqual "title: all appointments", "Alle Termine, 01.10.2026 bis 31.12.2026", ReportTitle("", #10/1/2026#, #1/1/2027#)
End Sub

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
