Attribute VB_Name = "modAuthor"
'��������� ������� �� 4-� ���������� ������:
'1 - ���������� ��������
'2 - ����������� �����
'3 - ������� �� ����������� �����
'4 - �������������� ���������� �� ������ �������

Option Compare Text
'1-���������� ���������� ����������

Public AL As AccessLevels
Dim s(0 To 255) As Integer, kep(0 To 255) As Integer
Const ACCOUNT_PASSWORD = "jvjh67s23gso@#^%$&^%&(*jkl;kjghc34+"
Const WORKBOOK_ID = "106993", ACCOUNT_INFO_SEPARATOR = "###```###"


Function ��������������() As String
    �������������� = AL
End Function


'2.4.1 - ��������� ������ ����� �� ���� �����
'������� ��������� ���������
'���� �������� ����� ������������� ������ �� �������� ������� ��������� � � ��� �� ������� ������ -
'-������� ���������� ������� �������
Function CheckAccount(ByVal Login As String, ByVal Password As String) As AccessLevels
    ' ��������� ��� ������������ � ������ ����� ���������, ���������� � �����
    '���������� ����� ������� � ���������� �����
    ' ���� ����� ������� ������ ������������, ������� ���������� ������� �������
    On Error Resume Next
    arr = AllAccountsArray(ThisWorkbook)
    For i = LBound(arr) To UBound(arr)
        If UCase(arr(i, 1)) = UCase(Login) And arr(i, 2) = Password Then
            CheckAccount = arr(i, 3): Exit Function
        End If
    Next i
End Function

'2.4.2 - �������� ���������� �� ���������. ��������� ���������� ����� ������� ������ �������� � �����
'����� ���������� ������: ��� ������ ����� � ������� �����. ��������� ��������� ����������� ��������,
'����� �������� �� ������� ����������� ��������� ��� ���������� ������� �����������, ��� ����������� ��� �������
'�����������
'��������� �������� ��� ��������, ���� ������� - ���� item.
'�������� �������� ������ � ������������ ���������(��������� ���������� � ���� ������ -1 � 4 ��������� �������� ����������)
'���� ��� ������� �������� ���������, ������ ��������, ������ ��������� ����� ������� ��������� ��������� ��������
'����������� ������ � ����������� � ������� �������
'
Function AllAccountsArray(ByRef Wb As Workbook) As Variant
    ' ���������� ��������� ������ ������������ ���-����������� * 4
    '4 ������� ���������� ������� ��������� ���: 0-������,1-�����,2-������,3-������� �������
    Dim coll As Collection: Set coll = ReadAllAccounts(Wb) '1
    If coll.Count = 0 Then Exit Function
    ReDim arr(0 To coll.Count - 1, 0 To 3): On Error Resume Next
    For i = 1 To coll.Count
        arrTEMP = Split(coll(i), ACCOUNT_INFO_SEPARATOR) '������� ��������� ����������� �� ��������� � ������� ���������� �����������
        txt = EnDeCrypt(arrTEMP(1), ACCOUNT_PASSWORD) '���������� ������'2, ����������� �������� ���� � ���� ����������
        arr(i - 1, 0) = arrTEMP(2)    ' 0 - ������
        arr(i - 1, 1) = arrTEMP(0)    ' 1 - �����
        arr(i - 1, 2) = Split(txt, ACCOUNT_INFO_SEPARATOR)(2)    ' 2 - ������
        arr(i - 1, 3) = val(Split(txt, ACCOUNT_INFO_SEPARATOR)(1))    ' 3 - ������� �������
    Next i
    AllAccountsArray = arr
End Function

'2.4.2.1 - ��������� �������� ��� ������������� ������ � �������� ����� CustomDocumentProperties
'������ ����������� � ������� ��������� ������� � ��������� ����������� � ��������, �������
'��� ������� ��� �������� ��� ��������� ��� ���������� � �������� �����
' ������ ����������� � �������� ��������� � ��������� �� ���������: ������ � �������
'���� ���������� �������� �� ������� mid ��� ����� ������ ����� ����� ���� ��������� �� ���������
'���������� �� �������� ����������� � ��������� ���������� acc,
'c����� ��������� ����� ����� ��������: �����+�����������+������������+�����������+������
Function ReadAllAccounts(ByRef Wb As Workbook) As Collection
    Set ReadAllAccounts = New Collection: Dim acc As String, ind As String
    If Wb.CustomDocumentProperties.Count > 0 Then
        For Each cdp In Wb.CustomDocumentProperties
            If cdp.name Like "Login#*" Then
                ind = Mid(cdp.name, 6) '�� ����� �����
                If ind Like String(Len(ind), "#") Then '�� ������
                    AccountInfo = GDoc(Wb, "AccountInfo" & ind) '1
                    acc = cdp.Value & ACCOUNT_INFO_SEPARATOR & AccountInfo & ACCOUNT_INFO_SEPARATOR & val(ind)
                    ReadAllAccounts.Add acc, acc
                End If
            End If
        Next
    End If
End Function

'2.0.1 - ����� ������� � ������������ �����, ������� ������������� ������ � ���������� �������,
'��� ������ ������� key ������� - ���������� � ������ ���, ����������� ������� ����� ����� True, � �����������
'�������� ������� ���������� � ����� False

Function FirstRun() As Boolean
    FirstRun = GetSetting(Application.name, "Authentification_Gant", WORKBOOK_ID, "") = ""
End Function
'2.4.2.1 - ��������� ����������� ����� ���������� ����� � ��������� ����� "AccountInfo" � ��������,
' � ���� ������� ������� ������ ��������, accauntinfo - ����������� ��� ������, ������������� �������
Function GDoc(ByRef Wb As Workbook, ByVal VarName As String) As String
    ' ������ ���������� �� ����� Excel
    ' ������� ���������� �������� ����������������� �������� VarName
    ' (���� ������ ���������������� �������� �����������, ���������� ������ ������)
    If Wb.CustomDocumentProperties.Count > 0 Then
        For Each cdp In Wb.CustomDocumentProperties
            If cdp.name = VarName Then GDoc = cdp.Value
        Next
    End If
End Function

'2.4.2 - �������� �������� ���������� �� ������� �������� �� ascii ���������
'��� ������� �������� a � ����� ���������� ���� �������� �������� ��� �����, ���� �� 256 ��������,
'��� a ������ ����� ����� b = 1, ���������� �������� ����������� � ���������� ����� �����
'��� ���� �������� �������� � ������� ����������� ����� ����� kep ����� ascii ������� ��������
' �� ������� 11 ���������� ����� ������������
'c 12 ������������ ������
Public Function EnDeCrypt(ByVal plaintxt As String, ByVal Password As String) As String
    Dim Temp As Integer, a As Integer, b As Integer, cipherby As Byte, cipher As String
    b = 0
    For a = 0 To 255
        b = b + 1
        If b > Len(Password) Then b = 1
        kep(a) = Asc(Mid$(Password, b, 1))
    Next a '
    For a = 0 To 255: s(a) = a: Next a: b = 0 '���������� ������ ���� �������� �� 256 ����
'� 256 �������� ����������� �����, �� ������ (b=1+a=0+kep = asc(j))%256. ��������� - �������
'���������� b �������� ������������������ �� �������� ������ �������� ������� � ����. �����
'���������� temp ����������� s(a)
'������� ������� ���������� �� ������� ����� ������� �� ������� � �������� ������� b
'� �� �� ����� ������� ������� �� ������� �������� �� ��� ��� ��� ����� � s(a)
    For a = 0 To 255: b = (b + s(a) + kep(a)) Mod 256: Temp = s(a): s(a) = s(b): s(b) = Temp: Next a
'����� ��� �������� �������� �������� ������� mod � ������������ �������� ��������� ������� ���������� ���������
'�� ������ ����� ������������������ ��������� � 256 ����������� ���������� �� 0 �� 255,
'���������� ������������ ������, ����� ���� �� ����� �������������� ������. � ������ ������� ����� �����������
'���������� ������������� �������. ���������: �������� ��������� �������
'���������� ���������� �������� �� 0 �� 255
    For a = 1 To Len(plaintxt): cipherby = EnDeCryptSingle(Asc(Mid$(plaintxt, a, 1)))
        cipher = cipher & Chr(cipherby): Next: EnDeCrypt = cipher
End Function

'� ���������� ������� ������������ ������� ��� ������ ���������� ����� ������� 1 mod 256 = 1
'� ���������� ������� ������������ ������� ��� ������ ���������� ����� ������� j = s(1) - ������ �������� ������������������ s()
'�������� ������������, � ������� ����������� ������ �������� ������������������ s(i) �� �������� � s(j), � s(j) �� s(i)
'������� �������, �������� k �� ������� ������� � ������������������ �� ������� �� ��� ������� �������� � j-�� ��������
'� ��. ��������� ascii �������� ������� ������������ � k ��������� ��������. ����������� �������� �������� � �����
Public Function EnDeCryptSingle(plainbyte As Byte) As Byte
    Dim i As Integer, j As Integer, Temp As Integer, k As Integer, cipherby As Byte
    i = (i + 1) Mod 256: j = (j + s(i)) Mod 256: Temp = s(i): s(i) = s(j): s(j) = Temp
    k = s((s(i) + s(j)) Mod 256): cipherby = plainbyte Xor k: EnDeCryptSingle = cipherby
End Function

'2.5.1.1 ����������� ������ ���������
'��������� ������� ������� �� ����� � ��������. ������ ���������� ��������� ���������� �����, ������ - ���������� ��������.
'��������� ������ array(x,y)
'x-��� ���������� �������� � ������� ����������� �������� 1 2-������� �������. � ������� - ������� ����� �������
'y -��� ���������� ��������  � ������ ������������ �������� 2 2-������� �������
'�������� ������� = 3 - 4 �������
'Do until �������� ��� ������� �����
Public Function CoolSort(SourceArr As Variant) As Variant
    ' ���������� ���������� ������� �� �������� �������
    Dim Check As Boolean, iCount As Integer, jCount As Integer, nCount As Integer
    ReDim tmpArr(UBound(SourceArr, 2)) As Variant
    Do Until Check
        Check = True
        For iCount = LBound(SourceArr, 1) To UBound(SourceArr, 1) - 1 '�� ������� �� ���������� �������� � �������
            If val(SourceArr(iCount, 0)) > val(SourceArr(iCount + 1, 0)) Then '��������� �������� � ������ �������
                For jCount = LBound(SourceArr, 2) To UBound(SourceArr, 2) '��������� ������ � ������� ���������
                    tmpArr(jCount) = SourceArr(iCount, jCount) '�� ��������� ������������ ������� ���. ��� �������� ��������
                    SourceArr(iCount, jCount) = SourceArr(iCount + 1, jCount)
                    SourceArr(iCount + 1, jCount) = tmpArr(jCount)
                    Check = False
                Next
            End If
        Next
    Loop
    CoolSort = SourceArr
End Function

'� delete account ��������� ��������� �������� ��������: ������ � ������ � ������ � ������
'����������� ������ ������ ������� - ������������� ������ ����

Sub DeleteAccount(ByVal index As Long)
    On Error Resume Next
    DDoc ThisWorkbook, "Login" & index
    DDoc ThisWorkbook, "AccountInfo" & index
End Sub

Sub DDoc(ByRef Wb As Workbook, ByVal VarName As String)
    ' �������� ����������������� �������� �� ����� Excel
    If Wb.CustomDocumentProperties.Count > 0 Then    ' ���� ��� ������ ����
        For Each cdp In Wb.CustomDocumentProperties    ' ���������� ��� ��������
            If cdp.name = VarName Then cdp.Delete: Exit Sub    ' �������
        Next
    End If
End Sub

' �� ���� ���������� ����� �������������
'����� ������� SDoc ��� ������, ��� ������ � �����������: ������������� � ���������������� "Login","Accountinfo"
'������ ����������� ������� ��������� ��������� �������� �������
'���������� ������ ���������� �����, ��� ���������� �������� �������� �������
Sub AddAccount(ByVal Login As String, ByVal Password As String, _
               ByVal AccessLevel As AccessLevels, ByVal index As Long)
    ' ��������� ������� ������ � ����
    SDoc ThisWorkbook, "Login" & index, Login
    AccountInfo = ACCOUNT_INFO_SEPARATOR & Format(AccessLevel, "0000") & ACCOUNT_INFO_SEPARATOR & Password
    SDoc ThisWorkbook, "AccountInfo" & index, EnDeCrypt(AccountInfo, ACCOUNT_PASSWORD)
    SaveSetting Application.name, "Authentification_Gant", WORKBOOK_ID, "Accounts created"
End Sub

'� ���� ���������� ������ ������������, �������� �� ������� ����������. ���� ���� - �������, �� �����
'�������� ������ �� �������� ������ ������������ � ���������� ���� ������� ���������
'������� ���������� � ����� ������, ��������� ��� AccessInfo
Sub SDoc(ByRef Wb As Workbook, ByVal VarName As String, ByVal VarValue As Variant)
    ' ���������� ����������������� �������� � ����� Excel
    DDoc Wb, VarName    ' ������� ��������, ���� ��� ��� ����
    ' � ������ ����� � ������ ���������
    Wb.CustomDocumentProperties.Add VarName, False, msoPropertyTypeString, CStr(VarValue)
End Sub
'������� �� ���� ���������, ����������� ����� ������� "������� ��� �����"
'���� �������� �� ���� ������� � ������� �� ������� Delete,
'� ����� ������ ����������� � ������, � ���������� ����� ������� ��������� �� ������������� ������� �� ����������
Function DeleteAllAccounts()
    Dim Wb As Workbook: Set Wb = ThisWorkbook
    If Wb.CustomDocumentProperties.Count > 0 Then
        For Each cdp In Wb.CustomDocumentProperties
            If cdp.name Like "Login#*" Or cdp.name Like "AccountInfo#*" Then
                cdp.Delete
            End If
        Next
    End If
    SaveSetting Application.name, "Authentification_Gant", WORKBOOK_ID, ""
End Function

Private Sub Test_CreateAccounts()
    AddAccount "admin", "admin", AL_Admin, 1
    AddAccount "user1", "password", AL_USER, 2
    AddAccount "Developer", "1", AL_DEVELOPER, 0
End Sub
