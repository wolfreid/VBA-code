VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public frm1 As Object 'Userform
Public WithEvents objControl As MSForms.CommandButton
Attribute objControl.VB_VarHelpID = -1
Public WithEvents objInputBox As MSForms.textBox
Attribute objInputBox.VB_VarHelpID = -1
Public WithEvents objLabel As MSForms.Label 'Button
Attribute objLabel.VB_VarHelpID = -1
Public Dict As Object
Private Nom As String
Private colText As New Collection
Private PrBody As New cProjectBlocks
Private msCount As Integer

Private Sub Class_Initialize()
'create class
Set Dict = CreateObject("Scripting.dictionary")
End Sub

Public Function UseProjectProperties()
PrBody.ShowListData
Set colText = UseCollection(PrBody.Datanames)
UseProjectProperties = colText.Count
End Function
Property Let GetDataCount(sVal As Integer)
    msCount = sVal
End Property
Property Get GetDataCount() As Integer
    GetDataCount = msCount
End Property

Public Function Show_Form()
NewUsf "ClearDataList"
NewLabel "lblData", PrBody.words, 200, 100, 5, 5
NewInputBox "txtData", 120, 30, 5, 100
NewButton "btnOk", "Œ ", 120, 30, 5, 120
frm1.show
On Error GoTo fin
Show_Form = frm1.tag
Unload frm1
Exit Function
fin:
    If Show_Form = "" Then
        Err.Clear: Unload frm1: Exit Function
    Else
    MsgBox "Error detected" & Err.Number & Err.Description
    End If
 End Function


Private Sub NewUsf(monCaption As String)
'Creation of userform
Set frm1 = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
Nom = frm1.name
VBA.UserForms.Add (Nom)
Set frm1 = UserForms(UserForms.Count - 1)
With frm1
    .Caption = monCaption
    .Width = 150
    .Height = 200
End With
End Sub

Public Sub NewButton(name As String, Caption As String, Width As Double, Height As Double, Left As Double, Top As Double)
'Creation of a Control Button
Dim obj
Set obj = frm1.Controls.Add("forms.CommandButton.1", name)
If obj = True Then Exit Sub
Dim cls As New cForms
Set cls.frm1 = frm1
Set cls.objControl = obj
With cls.objControl
.Caption = Caption
.Move Left, Top, Width, Height
End With
Dict.Add name, cls
Set cls = Nothing
End Sub
Public Sub NewLabel(name As String, Caption As String, Width As Double, Hight As Double, Left As Double, Top As Double)
'Creation of a Control Button
Dim obj
Set obj = frm1.Controls.Add("Forms.Label.1", name)
If obj = True Then Exit Sub
Dim cls As New cForms
Set cls.frm1 = frm1
Set cls.objLabel = obj
With cls.objLabel
.Caption = Caption
.Move Left, Top, Width, 100
End With
Dict.Add name, cls
Set cls = Nothing
End Sub
Public Sub NewInputBox(name As String, Width As Double, Height As Double, Left As Double, Top As Double)
'Creation of a TextBox
Dim obj
Set obj = frm1.Controls.Add("Forms.TextBox.1", name)
If obj = True Then Exit Sub
Dim cls As New cForms
Set cls.frm1 = frm1
Set cls.objInputBox = obj
With cls.objInputBox
.Move Left, Top, Width, Height
End With
Dict.Add name, cls
Set cls = Nothing
End Sub

Private Sub objcontrol_Click()
' If Dict(2).Text = vbNullString Then
'        MsgBox "You need to enter an answer!", vbExclamation, "No Answer Found!"
'        objInputBox.SetFocus
'        Exit Sub
'    End If
frm1.Hide
End Sub

Private Sub objInputBox_Change()
    frm1.tag = objInputBox.Value
End Sub
Private Sub objInputBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub Class_Terminate()
Dim VBComp As VBComponent
Set Dict = Nothing
If Nom <> "" Then
Set VBComp = ThisWorkbook.VBProject.VBComponents(Nom) 'search
ThisWorkbook.VBProject.VBComponents.Remove VBComp 'delete
End If
End Sub







