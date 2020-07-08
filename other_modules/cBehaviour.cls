VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBehaviour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public NewRecord As Long
Public Model As ListObject
Public Property Get id() As range
Set id = Model.DataBodyRange(NewRecord, 1)
End Property

Public Property Get Profile() As range
    Set Profile = Model.DataBodyRange(NewRecord, 4)
End Property


Public Property Get Login() As range
    Set Login = Model.DataBodyRange(NewRecord, 2)
End Property


Public Property Get time() As range
    Set time = Model.DataBodyRange(NewRecord, 3)
End Property

Public Property Get action() As range
    Set action = Model.DataBodyRange(NewRecord, 5)
End Property

Public Property Get Source() As range
    Set Source = Model.DataBodyRange(NewRecord, 6)
End Property

Public Property Get Record() As range
    Set Record = Model.DataBodyRange(NewRecord, 7)
End Property

Public Property Get ExeStatus() As range
    Set ExeStatus = Model.DataBodyRange(NewRecord, 8)
End Property

'empty value checking behaviour collect
Public Function Checkvalue(rObj As range, ncols As Integer) As String
Dim cell As range
Dim bIsEmpty As Boolean
    bIsEmpty = False
For Each cell In rObj
    If IsEmpty(cell) = True Then
        If rObj.Cells(1, ncols).Address = cell.Address Then Exit For
        'An empty cell was found. Exit loop
        bIsEmpty = True
        Checkvalue = range("execution_status").Cells(2)
        Exit Function
    End If
Next cell
    Checkvalue = range("execution_status").Cells(1)
End Function






