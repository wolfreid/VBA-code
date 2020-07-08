VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Type Ranges
    workspace As range
    fixCell As range
End Type
Private dyn_environment As cEvents
Private rng As Ranges
Private Const DefReg = 11

Private Sub btnExit_Click()
    Call Выход
End Sub

Private Sub sbtnNewProject_Click()
     'Call NewProjectRecord
     Call StopTrackingEvents
End Sub

Private Sub Worksheet_Activate()
    Init_Interface = True
    With rng
        Set .fixCell = [i8]
        Set .workspace = [workspace]
    End With
End Sub

Private Sub Worksheet_Change(ByVal Target As range)
Dim iRow As Integer: Dim app As New Excel.Application
If rng.workspace Is Nothing Then Exit Sub
If InRange(Target, rng.workspace) Then
    iRow = Target.row
    If app.CountA(rng.workspace.Rows(iRow)) < 3 Then Exit Sub
    For i = 2 To 5 Step 1
        If Target.column = i Then
            If Not IsEmpty(Target) And Target.Style.name = "Текст" Then
                 Application.EnableEvents = False
                Target.Value = UCase(Target.Value)
                Exit For
            ElseIf Not IsEmpty(Target.Value) And Target.Style.name <> "Текст" Then
                MsgBox "cell in workspace did not create"
                Application.EnableEvents = False
                Target.Clear
                Exit For
            End If
        End If
    Next
    If app.CountA(rng.workspace.Rows(iRow)) >= 6 Then
        MsgBox app.CountA(rng.workspace.Rows(iRow))
        If Not Environment_Initialize() Then Exit Sub
    End If
Else
    MsgBox "Active Cell NOT In Range!" & vbNewLine & _
    "This cell cannot been changed"
    Application.EnableEvents = False
    If InRange(Target, rng.fixCell) Then: Target.ClearContents: Else: Target.Clear
End If
Application.EnableEvents = True
End Sub
Private Sub Worksheet_Deactivate()
If rng.workspace Is Nothing Then Exit Sub
Init_Interface = False: ProjectRow_Reset
Application.EnableEvents = False
Do While Application.CountA(rng.workspace.Columns(1)) <> 0
    rng.workspace.Rows(1).Delete Shift:=xlUp
    inc = inc + 1
Loop
With ThisWorkbook.Names.Item("workspace")
    .RefersTo = .RefersToRange.Resize(rng.workspace.Rows.Count + inc, rng.workspace.Columns.Count)
End With
Set rng.workspace = Nothing
Set rng.fixCell = Nothing
Application.EnableEvents = True
End Sub

Private Function Environment_Initialize() As Boolean
Environment_Initialize = True
If dyn_environment Is Nothing Then
    Set dyn_environment = New cEvents
     dyn_environment.Add rng.fixCell, dyn_environment.Count '1
     dyn_environment.Add rng.workspace, dyn_environment.Count '2
End If
Environment_Initialize = False
End Function
'
Sub StopTrackingEvents()
If Init_Interface = True Then
    ProjectRow_Reset
    GlobalCancel = False
    If Not dyn_environment Is Nothing Then Set dyn_environment = Nothing
End If
End Sub
