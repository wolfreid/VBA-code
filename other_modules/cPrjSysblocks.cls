VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPrjSysblocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public conseq As Integer
Public ConseqStr As String
Private Activedata As New cProjectBlocks
Public Property Get ActionSet() As range
Set ActionSet = range("Action")
End Property

Sub ActionType(ByVal rCons As Integer)
Dim cell As range
For Each cell In ActionSet.Columns(1).Cells
    If cell = rCons Then
     ConseqStr = ActionSet.Cells(cell.row - 1, 2)
     Exit Sub
    End If
Next

End Sub


Public Property Get MySession() As range
Set MySession = Activedata.PermWorksheet.range("Q2")
End Property


Public Property Get WrngSession() As range
Set WrngSession = Activedata.PermWorksheet.range("R2")
End Property
'Public Property Get MyRibbon() As Range
'Set MyRibbon = Activedata..Range("Q2")
'End Property
Public Property Get IRibbon() As range
Set IRibbon = Activedata.TitleWorksheet.range("A1")
End Property
Public Property Get LegalInform() As range
Set LegalInform = Activedata.TitleWorksheet.range("M8")
End Property


