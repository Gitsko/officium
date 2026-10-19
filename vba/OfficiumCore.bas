Attribute VB_Name = "OfficiumCore"
' Report logic without Outlook objects, covered by OfficiumSelfTest.
Option Explicit

Public Enum OfficiumPeriod
    opThisMonth = 1
    opLastMonth = 2
    opThisQuarter = 3
    opLastQuarter = 4
End Enum

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
