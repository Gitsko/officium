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

Private Sub AssertEqual(ByVal testName As String, ByVal expected As Variant, ByVal actual As Variant)
    If expected = actual Then
        passed = passed + 1
    Else
        failed = failed + 1
        Debug.Print "FEHLER " & testName & ": erwartet " & expected & ", erhalten " & actual
    End If
End Sub
