VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TitleSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'Private WithEvents mouseEvent As cMouse
Private Sub Worksheet_SelectionChange(ByVal Target As range)
On Error Resume Next
If Not Intersect(Target, range("M8")) Is Nothing Then
Call LegalInform_Form
End If
End Sub

'Private Sub Worksheet_Activate()
'Set mouseEvent = New cMouse
'mouseEvent.Initialize
'End Sub

'Private Sub mouseEvent_MouseClick()
'MsgBox "Click Event Risen from Shape Selecting", vbInformation
'End Sub

'Private Sub Worksheet_Deactivate()
'mouseEvent.Terminate
'Set mouseEvent = Nothing
'End Sub

