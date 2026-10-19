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

Private Sub AssertEqual(ByVal name As String, ByVal expected As Variant, ByVal actual As Variant)
    If expected = actual Then
        passed = passed + 1
    Else
        failed = failed + 1
        Debug.Print "FEHLER " & name & ": erwartet " & expected & ", erhalten " & actual
    End If
End Sub
