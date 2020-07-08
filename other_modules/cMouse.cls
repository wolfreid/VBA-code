VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event MouseClick(ByVal Target As range)
Private mInitialized As Boolean
Public Function Initialize() As Boolean
    Initialize = False
    If mInitialized Then Exit Function
    GlobalCancel = True
       If Not MouseMoveTest(Me) Then Exit Function
       mInitialized = True
    Initialize = True
End Function
Sub Click(ByVal Target As range)
    RaiseEvent MouseClick(Target)
End Sub

Private Sub Class_Initialize()
    mInitialized = False
End Sub

Private Sub Class_Terminate()
    If Not mInitialized Then Exit Sub
    Terminate
End Sub

Public Function Terminate()
    GlobalCancel = False
    mInitialized = False
End Function



