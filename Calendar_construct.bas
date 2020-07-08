Attribute VB_Name = "Calendar_construct"
Option Explicit
'В этом модуле работа с календарем

Public Sub CalendarCreate()
Dim clsCalendar As cCalendar
Dim clsmonth As CMonthes
Dim i As Long
Dim ft As ListObject
Dim shCal As Worksheet
Dim num As Integer
Dim iDayNumber As Integer
Dim dtLastDayInYear As Date
Dim dt As Date: dt = Date
dtLastDayInYear = DateSerial(Year(dt), 12, 31)
iDayNumber = DateDiff("d", CDate("1/1/" & Year(dt)), dtLastDayInYear) + 1
Set shCal = ActiveWorkbook.Worksheets("Celendar2")
Set ft = shCal.ListObjects(1)

Set clsCalendar = New cCalendar
For i = 1 To 12
    Set clsmonth = New CMonthes
    clsmonth.i = i
    clsmonth.month = monthname(i)
    clsCalendar.Add clsmonth
Next
For i = 1 To clsCalendar.Count
        With clsCalendar.Item(i)
            Debug.Print .month, .i
        End With
Next i
MsgBox iDayNumber
End Sub

Sub create_listmonth()
Dim month(12) As String
For i = 1 To 12
month(i) = monthname(i)
Next
MsgBox month(5)
End Sub


