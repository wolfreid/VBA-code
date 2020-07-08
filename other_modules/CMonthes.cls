VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMonthes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mi As Long
Private mMonth As String
Public Property Let i(ByVal li As Long)
    mi = li
End Property
Public Property Get i() As Long
    i = mi
End Property
Public Property Let month(ByVal smonth As String)
    mMonth = smonth
End Property
Public Property Get month() As String
    month = mMonth
End Property
