VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim MyMonthes As New Collection
Dim i, X, Y, rowtarget As Integer, output As String
Dim multimonth() As Integer
Dim bdrange, rangetot As range
Dim pvtitem As PivotItem, pf As PivotField
Dim MyMonth, MonthList
Dim TheName As String
'на стороне сводной таблицЫ
 Set Target = PivotTables("Calendars2")
 Set pf = Target.PivotFields("Month")
 Set bdrange = Target.DataBodyRange
 Set MyAddonMonthes = New Collection
 With Target
    With pf
        For Each pvtitem In .PivotItems
        TheName = pvtitem.name
        If Target.RowRange.Cells(2 + i).Value = TheName Then
            MyMonthes.Add TheName
            i = i + 1
        End If
        Next
    End With
    i = Empty
 End With
 X = bdrange.Rows.Count
 ReDim multimonth(1 To X)
 For rowtarget = 1 To X
        Set rangetot = bdrange.Cells(1).Offset(rowtarget - 1, 0)
        multimonth(rowtarget) = rangetot.Value
        output = output & MyMonthes(rowtarget) & " " & multimonth(rowtarget) & " " & rowtarget & vbNewLine
 Next rowtarget
 'MsgBox output, , "event"
    For i = 1 To rowtarget - 1
        CreateStatus multimonth(i), i, MyMonthes(i)
    Next
Set MyMonthes = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub




