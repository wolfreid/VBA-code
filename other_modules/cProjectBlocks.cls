VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProjectBlocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public words As String
Public Datanames As New Collection
Public Sheetnames  As New Collection

Public Property Get ProjectDashboard() As Workbook
    Set ProjectDashboard = ThisWorkbook
End Property
Public Property Get BehaviourWorksheet() As Worksheet
    Set BehaviourWorksheet = BehaviourSheet
End Property

Public Property Get CustomerWorksheet() As Worksheet
    Set CustomerWorksheet = CustomerSheet
End Property

Public Property Get StatusWorksheet() As Worksheet
    Set StatusWorksheet = StatusSheet
End Property

Public Property Get StageWorksheet() As Worksheet
    Set StageWorksheet = StageSheet
End Property

Public Property Get PermWorksheet() As Worksheet
    Set PermWorksheet = PermSheet
End Property

Public Property Get TitleWorksheet() As Worksheet
    Set TitleWorksheet = TitleSheet
End Property

Public Property Get DataList(iValue As Datasource) As ListObject
Dim tempStr As String: tempStr = objSetupData.Database(iValue) & iValue
Set DataList = range(tempStr).ListObject
End Property

Public Property Get DataHead(iValue As Datasource) As ListObject
Dim tempStr As String: tempStr = objSetupData.Database(iValue) & iValue
Set DataHead = range(tempStr).ListObject.HeaderRowRange
End Property

Public Sub ShowListData()
Dim SetList As Object
Dim key As Variant
Dim a As String, wisetab As String
Dim i, k As Integer, size As Integer
Const space As String = " "
Const totalsize As Integer = 20
Dim Sheet As Worksheet, Data As ListObject
Set SetList = CreateObject("Scripting.Dictionary")
i = 1
For Each Sheet In ProjectDashboard.Worksheets
    With Sheet
    For Each Data In .ListObjects
    size = totalsize - Len(Data.name)
        Do
            wisetab = wisetab + space
        Loop Until Len(wisetab) = size
        a = Data.name & wisetab & vbTab & Sheet.name
        wisetab = Empty
        SetList.Add a, str(i)
        Datanames.Add Data.name, str(i)
        Sheetnames.Add Sheet.name, str(i)
        i = i + 1
    Next
    End With
Next
For Each key In SetList.Keys
    words = words + SetList(key) + " " + key + vbCr
Next key
End Sub

'Public Sub ShowListSheets()
'Dim SetList As
'End Sub

Public Property Get Data(rId As Integer) As ListObject
Set Data = ThisWorkbook.Worksheets(Sheetnames(rId)).ListObjects(Datanames(rId))
End Property
