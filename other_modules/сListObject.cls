VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ÒListObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRow As Integer
Private mColumn  As Integer
Private mName As String
Private mWksname As String

Public Property Let row(ByVal lrow As Integer)
    mRow = lrow
End Property
'Public Property Get row() As Integer
'    row = mRow
'End Property
Public Property Let column(ByVal lcolumn As Integer)
    mColumn = lcolumn
End Property
'Public Property Get column() As Integer
'    column = mColumn
'End Property

Public Property Let name(lname As String)
    mName = lname
End Property

Public Property Let WorksheetName(lWs As String)
    mWksname = lWs
End Property

Public Sub CreateObject()
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets(mWksname)
ws.ListObjects.Add(xlSrcRange, range(Cells(1, 1), Cells(mRow, mColumn)), , xlYes).name = mName
ws.ListObjects(mName).TableStyle = "Black&White2"
End Sub


