Attribute VB_Name = "Officium"
' Officium - calendar reports for classic Outlook.
Option Explicit

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
