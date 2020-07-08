Attribute VB_Name = "SheetsInteractive"

Private m As Integer
Public filledspace As range
Public Init_Interface As Boolean

Sub ClearData_Щелчок()
Dim nameForm As String: nameForm = "Searching"
Dim objBlock As cProjectBlocks
Dim NumofBlock As Variant
Set objBlock = New cProjectBlocks
objBlock.ShowListData
forms.Generate_form nameForm, objBlock.words
ListObjects.ClearLstObj
End Sub


Sub NameRanges_Show_Щелчок()
Call RangeManeger.Ranges_Show
End Sub

Sub Data_Show()
Dim project As New cProjectBlocks
project.ShowListData
ActiveSheet.range("A15") = project.words
Set project = Nothing
End Sub



Public Sub ProjectRow_Activate()
Application.EnableEvents = False
Application.ScreenUpdating = False
Dim triangleUp As shape
Static i As Integer
Set triangleUp = InputSheet.Shapes("new_row")
Set workspace = InputSheet.range("workspace")
Set filledspace = workspace
If i > workspace.Rows.Count Then
    workspace.Interior.Color = xlNone
    workspace.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    workspace.Clear
End
End If
If i > m Then: i = m + 1: Else: i = m
If i = Empty Then i = 1
workspace.Rows(i).Select
With Selection
    .Interior.Color = vbWhite
    .Borders.Color = vbBlack
    .Style = "Текст"
    .Cells(1) = i
    With .Cells(.Columns.Count)
        .Font.name = "Wingdings"
        .Font.size = 18
        .Value = Chr(111)
    End With
End With
i = i + 1: m = i
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub

Public Sub ProjectRow_Deactivate()
Application.EnableEvents = False
Application.ScreenUpdating = False
Dim triangleDown As shape
Dim workspace As range
Static k As Integer
Set triangleDown = InputSheet.Shapes("del_row")
Set workspace = InputSheet.range("workspace")
If m < 1 Then End
If m > k Then: k = m - 1: Else: k = m
workspace.Rows(k).Select
With Selection
    .Interior.Color = xlNone
    .Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    .Clear
End With
k = k - 1: m = k
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub

Public Sub ProjectRow_Reset()
    m = Empty
End Sub
