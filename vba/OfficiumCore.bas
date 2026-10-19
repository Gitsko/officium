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

' True if one of the categories equals the customer or the subject contains it,
' both ignoring case. Categories are separated by "," or ";", depending on the
' Windows list separator. An empty customer matches every appointment.
Public Function MatchesCustomer(ByVal customer As String, ByVal subjectText As String, _
                                ByVal categoryList As String) As Boolean
    customer = Trim$(customer)
    If Len(customer) = 0 Then
        MatchesCustomer = True
        Exit Function
    End If
    If InStr(1, subjectText, customer, vbTextCompare) > 0 Then
        MatchesCustomer = True
        Exit Function
    End If
    Dim category As Variant
    For Each category In Split(Replace(categoryList, ";", ","), ",")
        If StrComp(Trim$(category), customer, vbTextCompare) = 0 Then
            MatchesCustomer = True
            Exit Function
        End If
    Next
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
