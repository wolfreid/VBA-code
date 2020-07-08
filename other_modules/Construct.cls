VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Construct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cLife
Private mModule As ListObject
Private Type myType
    objLife As cLife
    objModel As cModel
End Type
Private this As myType


Private Sub Class_Initialize()

    With this
        Set .objLife = New cLife
        Set .objModel = .objLife
    End With
End Sub

Private Property Get cLife_Model() As cModel
    Set cLife_Model = this.objModel
End Property






