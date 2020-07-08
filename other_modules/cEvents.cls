VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents thisProjectnative As Excel.Application
Attribute thisProjectnative.VB_VarHelpID = -1
Private WithEvents mouseEvent As cMouse
Attribute mouseEvent.VB_VarHelpID = -1
Private colInterface As Collection

Public Property Get Count() As Long
    Count = colInterface.Count
End Property

Public Sub Add(objControl As range)
    colInterface.Add objControl
End Sub

Private Sub Class_Initialize()
Set thisProjectnative = Excel.Application
Set mouseEvent = New cMouse
Set colInterface = New Collection
End Sub

Private Sub Class_Terminate()
mouseEvent.Terminate
Set thisProject_native = Nothing
Set mouseEvent = Nothing
Set colInterface = Nothing
End Sub

Private Sub mouseEvent_MouseClick(ByVal Target As range) '(ByVal shape As Object)
   If Target.Cells.Count > 1 Then Exit Sub
    If Not Intersect(Target, colInterface(2)) Is Nothing Then
            If .Value = Chr(111) Then
               .Value = Chr(254)
            Else
               .Value = Chr(111)
            End If
        End With
    End If
    Application.Cursor = xlNorthwestArrow
End Sub
Private Sub thisProjectNative_SheetChange(ByVal Sh As Object, ByVal Target As Excel.range)
'    LogEvent "SheetChange: " & Target.Address(False, False) & " in " & Sh.name
End Sub

Private Sub thisProjectNative_SheetActivate(ByVal Sh As Object)
'    LogEvent "SheetActivate: " & Sh.name
    apiEvent.Initialize
End Sub

Private Sub thisProjectNative_SheetDeactivate(ByVal Sh As Object)
'    LogEvent "SheetDeactivate: " & Sh.name
    mouseEvent.Terminate
End Sub


Private Sub thisProjectNative_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.range)
If Target.Address = colInterface(1).Address Then
    If mouseEvent Is Nothing Then
        Set mouseEvent = New cMouse
    Else
        Application.EnableEvents = False
        Exit Sub
    End If
    mouseEvent.Initialize
    Application.EnableEvents = False
End If
End Sub

Private Sub thisProjectNative_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Excel.range, Cancel As Boolean)
Dim anyRng As range
Set anyRng = Selection
    If Selection.Address = colInterface(1).Address Then
        Cancel = True
    End If
End Sub
'присваивание результата события происходит в Caption

Private Sub thisProjectNative_NewWorkbook(ByVal Wb As Excel.Workbook)
'    LogEvent "NewWorkbook: " & Wb.name
End Sub

Private Sub thisProjectNative_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Excel.range, Cancel As Boolean)
'    LogEvent "SheetBeforeRightClick: " & Target.Address(False, False) & " in " & Sh.name
End Sub

Private Sub thisProjectNative_SheetCalculate(ByVal Sh As Object)
'    LogEvent "SheetCalculate: " & Sh.name
End Sub

Private Sub thisProjectNative_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
'    LogEvent "SheetFollowHyperlink: " & Target.name & " in " & Sh.name
End Sub


Private Sub thisProjectNative_WindowActivate(ByVal Wb As Excel.Workbook, ByVal Wn As Excel.Window)
    LogEvent "WindowActivate: " & Wn.Caption & " in " & Wb.name
End Sub

Private Sub thisProjectNative_WindowDeactivate(ByVal Wb As Excel.Workbook, ByVal Wn As Excel.Window)
    LogEvent "WindowDeactivate: " & Wn.Caption & " in " & Wb.name
End Sub

Private Sub thisProjectNative_WindowResize(ByVal Wb As Excel.Workbook, ByVal Wn As Excel.Window)
    LogEvent "WindowResize: " & Wn.Caption & " in " & Wb.name
End Sub

Private Sub thisProjectNative_WorkbookActivate(ByVal Wb As Excel.Workbook)
    LogEvent "WorkbookActivate: " & Wb.name
End Sub

Private Sub thisProjectNative_WorkbookAddinInstall(ByVal Wb As Excel.Workbook)
    LogEvent "WorkbookAddinInstall: " & Wb.name
End Sub

Private Sub thisProjectNative_WorkbookAddinUninstall(ByVal Wb As Excel.Workbook)
'    LogEvent "WorkbookAddinUninstall: " & Wb.name
End Sub

Private Sub thisProjectNative_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, Cancel As Boolean)
'    LogEvent "WorkbookBeforeClose: " & Wb.name
End Sub

Private Sub thisProjectNative_WorkbookBeforePrint(ByVal Wb As Excel.Workbook, Cancel As Boolean)
'    LogEvent "WorkbookBeforePrint: " & Wb.name
End Sub

Private Sub thisProjectNative_WorkbookBeforeSave(ByVal Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
'    LogEvent "WorkbookBeforeSave: " & Wb.name
End Sub

Private Sub thisProjectNative_WorkbookDeactivate(ByVal Wb As Excel.Workbook)
'    LogEvent "WorkbookDeactivate: " & Wb.name
End Sub

Private Sub thisProjectNative_WorkbookNewSheet(ByVal Wb As Excel.Workbook, ByVal Sh As Object)
'    LogEvent "WorkbookNewSheet: " & Sh.name & " in " & Wb.name
End Sub

Private Sub thisProjectNative_WorkbookOpen(ByVal Wb As Excel.Workbook)
'    LogEvent "WorkbookOpen: " & Wb.name
End Sub

'Public Property Set FormTracking(ByRef oForm As UserForm)
'    Set CalcForms = oForm
'End Property




