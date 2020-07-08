VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Property Set DataForm(oList As ListObject): End Property
Property Get DataForm() As ListObject: End Property

 Property Let NewRecord(sVal As Long): End Property
Property Get NewRecord() As Long: End Property

Property Get id(sName As String) As range: End Property

Property Get Login(sName As String) As range: End Property

Property Get Statename(sName As String) As range: End Property


Property Get Online(sName As String) As range: End Property

Property Get Offline(sName As String) As range: End Property

Property Get Lifestate(sName As String) As range: End Property
