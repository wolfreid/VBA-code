Attribute VB_Name = "Merging"
Option Explicit
'Здесь работаем с Дашбордом

'working only with single month
'singlemonth нужeн чтобы  передать значение для последующего генерирования формулы
'В  коде rngOrigin нет координт значений листа базы, данные обновляются на основе текущей формульной разметки, потому она имее постоянный диапазон в 31 столбец
'переменная pivotcolscoun требуетс для обновления разметки, будет новая разметка с учетом  диапазона дашборда, на одну больше
'colcont - число столбцов по текущему значению singlemonth
'jstart - определяет начальную позицию новой интерактивной таблицы, раньше была константной
'dimstart - текущая пременная, статическая
Public Dashboard As Worksheet, shCal As Worksheet
Public pivotCalendar As PivotTable
Private ft As ListObject
'Public dyear As Integer

Public Sub CreateStatus(singlemonth, iter, mName)
Dim lastdate As Date, firstdate As Date
Dim u, i, j, colcnt As Integer
Static Data As Collection
Static dimstart  As Integer
 Set Dashboard = ActiveWorkbook.Sheets("Dashboard")
 Set shCal = ActiveWorkbook.Worksheets("Celendar2")
 Set ft = shCal.ListObjects(1)
 If iter = 1 Then
    Set Data = New Collection
    dimstart = 6
    Rows(3).UnMerge
 End If
 Data.Add singlemonth
 colcnt = singlemonth - 1
 If iter > 1 Then
    singlemonth = Data(Int(iter - 1))
    dimstart = dimstart + singlemonth
 End If
    For i = 1 To ft.ListRows.Count
        If ft.DataBodyRange.Cells(i, 2).Value = mName Then
            Dashboard.Cells(4, dimstart + j) = ft.DataBodyRange.Cells(i, 1).Value
            Dashboard.Cells(4, dimstart + j).NumberFormat = "dd"
            j = j + 1
        End If
    Next i
i = Empty
j = Empty
Cells(5, dimstart).formula = "=" & Cells(4, dimstart).Address(RowAbsolute:=False, ColumnAbsolute:=False)
range(Cells(5, dimstart), Cells(5, dimstart + colcnt)).Select
Selection.FillRight

For i = 4 To 5
    range(Cells(i, dimstart + colcnt + 1), Cells(i, dimstart + colcnt + 1).End(xlToRight)).ClearContents
Next i

Cells(3, dimstart) = UCase(mName)
Cells(3, dimstart).Select
With Selection
    .Font.name = "Arial"
    .Font.size = 14
    .Font.Color = vbWhite
    .Interior.Color = RGB(58, 56, 56)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With
range(Cells(3, dimstart), Cells(3, dimstart + colcnt)).Select
With Selection
    .Merge
    .Borders.Color = vbWhite
End With
range(Cells(3, dimstart + colcnt + 1), Cells(3, curdaysinyear())).Select
With Selection
    .ClearContents
    .Interior.ColorIndex = xlNone
    .Borders.LineStyle = xlNone
End With
range("A1").Activate
End Sub











