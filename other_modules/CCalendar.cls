VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCalendar  As New Collection


Public Sub Add(clsmonth As CMonthes)
    mCalendar.Add clsmonth, CStr(clsmonth.i)
End Sub

Public Property Get Count() As Long
    Count = mCalendar.Count
End Property

Public Property Get Item(lId As Long) As CMonthes
    Set Item = mCalendar(lId)
End Property
