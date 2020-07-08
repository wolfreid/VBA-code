Attribute VB_Name = "ListObjects"
Option Explicit
'Фабрика умных таблиц, здесь настраиваются и производятся умные таблицы

'assign every listobjects default style
Sub StyleLstObj()
Dim tbl As ListObject
Dim ws As Worksheet
Dim i, c, t As Integer
Dim fnt As Font, fnt2 As Font
'On Error Resume Next

For Each ws In ThisWorkbook.Worksheets
With ws
    For Each tbl In .ListObjects
        Set fnt = tbl.DataBodyRange.Font
        Set fnt2 = tbl.HeaderRowRange.Font
        c = tbl.DataBodyRange.Rows.Count
        t = tbl.DataBodyRange.Columns.Count
        tbl.TableStyle = "Black&White2"
        With fnt
            .Color = vbBlack
            .size = 10
            .name = "Arial Narrow"
            .Bold = False
        End With
        With fnt2
            .Color = vbWhite
            .name = "Calibri"
            .Bold = False
            .size = 12
        End With
        tbl.HeaderRowRange.WrapText = False
        tbl.DataBodyRange.WrapText = False
        tbl.range.Columns.AutoFit
    Next
End With
Next
End Sub


Sub CreateLstObj()
Dim FinalRow, LastColumn As Integer
Dim itemObject As сListObject
Dim strName, strSheetname As String
Set itemObject = New сListObject
strName = InputBox("Type Data name")
strSheetname = InputBox("Type target sheet name") '"Status_data"
FinalRow = ThisWorkbook.Worksheets(strSheetname).Cells(Rows.Count, "A").End(xlUp).row
LastColumn = ThisWorkbook.Worksheets(strSheetname).Cells(1, Columns.Count).End(xlToLeft).column
itemObject.column = LastColumn
itemObject.row = FinalRow
itemObject.name = strName
itemObject.WorksheetName = strSheetname
itemObject.CreateObject
Set itemObject = Nothing
End Sub







