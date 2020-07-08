Attribute VB_Name = "ModStyles"
Option Explicit

'Фабрика стилей,Здесь хранятся стили для различных обьектов, модель стилей прикручивается к обьекту здесь

'вывод список всех стилей на лист
Sub Show_ListStyles()
    Dim oSt As Style
    Dim oCell As range
    Dim lCount As Long
    Dim oStylesh As Worksheet
    Set oStylesh = ThisWorkbook.Worksheets("ConfigStyles")
    With oStylesh
        lCount = oStylesh.UsedRange.Rows.Count + 1
        For Each oSt In ThisWorkbook.Styles
            On Error Resume Next
            Set oCell = Nothing
            Set oCell = Intersect(oStylesh.UsedRange, oStylesh.range("A:A")).Find(oSt.name, _
                oStylesh.range("A1"), xlValues, xlWhole, , , False)
            If oCell Is Nothing Then
            lCount = lCount + 1
            .Cells(lCount, 1).Style = oSt.name
            .Cells(lCount, 1).Value = oSt.NameLocal
            .Cells(lCount, 2).Style = oSt.name
            End If
        Next
    End With
End Sub

'модель стиля для ячеек
Sub Make_Style()
    ActiveWorkbook.Styles.Add name:="PrZemo1"
    With ActiveWorkbook.Styles("PrZemo1")
        .IncludeNumber = True
        .IncludeFont = True
        .IncludeAlignment = True
        .IncludeBorder = True
        .IncludePatterns = True
        .IncludeProtection = True
    End With
    With ActiveWorkbook.Styles("PrZemo1").Font
        .name = "Arial Narrow"
        .size = 11
        .Bold = False
        .Italic = False
        .Underline = xlUnderlineStyleNone
        .Strikethrough = False
        .Color = -16776961
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With ActiveWorkbook.Styles("PrZemo1")
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
    End With
End Sub
'модель стиля для форматируемой таблицы
Sub Make_tableStyle()
Dim oStyle As TableStyle
Set oStyle = ThisWorkbook.TableStyles.Add("Black&White2")
With oStyle
    .TableStyleElements(xlHeaderRow).Font.Color = vbWhite
    .ShowAsAvailableTableStyle = True
    .TableStyleElements(xlHeaderRow).Interior.Color = vbBlack
    .TableStyleElements(xlHeaderRow).Borders.Color = vbWhite
    .TableStyleElements(xlHeaderRow).Font.Bold = False
    .TableStyleElements(xlRowStripe1).Interior.Color = vbWhite
    .TableStyleElements(xlRowStripe2).Interior.Color = vbWhite
    .TableStyleElements(xlRowStripe1).Borders.Color = vbBlack
    .TableStyleElements(xlRowStripe2).Borders.Color = vbBlack
    .TableStyleElements(xlRowStripe1).Font.Color = vbBlack
    .TableStyleElements(xlRowStripe2).Font.Color = vbBlack
End With
ActiveWorkbook.DefaultTableStyle = "Black&White2"
End Sub
'создание формы обьекта типа фигуры и редактирование фигуры






