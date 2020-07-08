Attribute VB_Name = "DataCollection"
'����� �������� ������ ������
Private CollObj As Collection
Private objBehvr As cProjectBlocks
Private objData As cProjectBlocks
Private MainBase As Datasource
Private form As Integer
'nform �� ������������
'���� ����� ������� �� ������ �����
Sub BehaviourCollect(ByVal tBase As Datasource)
Dim Act As cBehaviour: Set Act = New cBehaviour
Dim CA_session As Config_Behaviour: CA_session = Behaviour_array '��������� ������������
Dim CollRef As Collection, Item As Variant, index As Integer
If Not CollObj Is Nothing Then
    Set CollRef = UseCollection(CollObj)
Else
    MsgBox "Alert, Collection does not exist", vbCritical: Stop: Exit Sub
End If
MainBase = DS_Session
'�������� ����, �� � ������� ��� �������� ������
Set objBehvr = New cProjectBlocks
Set Act.Model = objBehvr.DataList(MainBase)
'If nform = 0 Or nform = 1 Then Let DS = DS_life
DS = tBase
With Act
    .Model.ListRows.Add
    .NewRecord = .Model.ListRows.Count
    .id = .NewRecord '������ ��������
    .Login = CollRef("Login") '������ ��������
    .time = Now '�������� 3
    .Profile = CA_session.Profiles(Get_ALname(AL)) '�������� 4
    .action = CA_session.Actions(form) '�������� 5
    .Source = VerifyTable(CA_session.Source_databases(DS), DS) '�������� 6
    .Record = CollRef("ID") '�������� 7
    .ExeStatus = .Checkvalue(.Model.ListRows(.id).range, .Model.ListColumns.Count) '�������� 8
End With
Set Act = Nothing
Set objBehvr = Nothing
Set CollRef = Nothing
End Sub
'���� ����� �����
Sub LifeCollect(ByVal sName As String, nform As Integer)
Dim CA_life As Config_Lifedata: CA_life = Life_array
Dim objConstruct As Construct: Dim objILife As cLife: Dim DBname As String
Dim cell As range, colID As String, arrHeader As Variant
Dim fillers(1) As Variant: fillers(0) = "<Waiting>": fillers(1) = CDate(Now)
'Dim objConfig As cPrjSysblocks: Set objConfig = New cPrjSysblocks
    Set CollObj = New Collection
    Set objData = New cProjectBlocks
    Set objConstruct = New Construct
    Set objILife = objConstruct
    Set objILife.Model.DataForm = objData.DataList(DS_life)
'    For Each key In CollRef
    arrHeader = Get_Headers(objILife.Model.DataForm.HeaderRowRange)
    With objILife.Model
     Let DBname = .DataForm.name
     If nform = 0 Then '������ ���������
            .DataForm.ListRows.Add
            .NewRecord = .DataForm.ListRows.Count
            .id(DBname) = .NewRecord
            .Login(DBname) = sName  ' '������ ���������
            .Statename(DBname) = CA_life.Statename(nform)
            .Online(DBname) = Now
            .Offline(DBname) = fillers(nform)
            .Lifestate(DBname) = CA_life.Lifestate(nform)
     ElseIf nform = 1 Then
            .NewRecord = .DataForm.ListRows.Count
            .Offline(DBname) = fillers(nform)
            .Lifestate(DBname) = CA_life.Lifestate(nform)
     End If
     For Each cell In .DataForm.ListRows(.id(DBname)).range
        colID = Replace(CStr(arrHeader(CollObj.Count)), " ", "")
        CollObj.Add cell, colID
     Next
    End With
   ' Debug.Print CollObj("Login")
    form = nform
    BehaviourCollect (DS_life)
End Sub
Sub CustomerCollect(ByVal col As Collection)
End Sub







