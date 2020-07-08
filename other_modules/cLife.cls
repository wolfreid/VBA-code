VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cModel
Private mData As ListObject
Private mNewRecord As Long

'Let's agree that every daughter class implements the method
'abstract 'super' that returns the IBase interface


Public Property Get Model() As cModel: End Property
Private Property Set cModel_DataForm(oList As ListObject)
Set mData = oList
End Property
Private Property Get cModel_DataForm() As ListObject
Set cModel_DataForm = mData
End Property
Private Property Let cModel_NewRecord(sVal As Long)
mNewRecord = sVal
End Property

Private Property Get cModel_NewRecord() As Long
cModel_NewRecord = mNewRecord
End Property


Private Property Get cModel_id(sName As String) As range
Set cModel_id = LifeSheet.ListObjects(sName).DataBodyRange(mNewRecord, 1)
End Property
Private Property Get cModel_StateName(sName As String) As range
Set cModel_StateName = LifeSheet.ListObjects(sName).DataBodyRange(mNewRecord, 3)
End Property

Private Property Get cModel_Login(sName As String) As range
Set cModel_Login = LifeSheet.ListObjects(sName).DataBodyRange(mNewRecord, 2)
End Property

Private Property Get cModel_Online(sName As String) As range
Set cModel_Online = LifeSheet.ListObjects(sName).DataBodyRange(mNewRecord, 4)
End Property

Private Property Get cModel_Offline(sName As String) As range
Set cModel_Offline = mData.DataBodyRange(mNewRecord, 5)
End Property

Private Property Get cModel_LifeState(sName As String) As range
Set cModel_LifeState = mData.DataBodyRange(mNewRecord, 6)
End Property






