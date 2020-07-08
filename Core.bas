Attribute VB_Name = "Core"
Private Sub delrange()
Dim objName As name
Set objName = range("UserNow").name
objName.Delete
End Sub
Function Get_array_tablesnames() As Variant
Dim i As Integer: i = 0
Dim arrays() As Variant
Dim Data As ListObject, Sheet As Worksheet
ReDim arrays(i)
For Each Sheet In ThisWorkbook.Worksheets
    With Sheet
        For Each Data In .ListObjects
        arrays(i) = Data.name
        i = i + 1
        ReDim Preserve arrays(i)
        Next
    End With
Next
Get_array_tablesnames = arrays
End Function
Public Function VerifyTable(sValue As Variant, customValue As Datasource) As String
Dim classdataname As cConfigurations
Set classdataname = New cConfigurations
Dim timeVar As String
tempvar = StrReverse(sValue)
If val(tempvar) = CInt(customValue) Then
    MsgBox "The order is not broken, because corresponds values" & sValue & "=" & customValue
    VerifyTable = classdataname.Database(customValue)
End If
Set classdataname = Nothing
End Function

Public Function RangeInArray(objRange As range) As Variant
Dim tempArr(), tempArr2() As Variant
X = objRange.Count
ReDim tempArr(X): ReDim tempArr2(X - 1)
tempArr = Application.Transpose(objRange)
For i = 0 To X - 1: tempArr2(i) = tempArr(i + 1): Next
RangeInArray = tempArr2
End Function

Public Function Get_Headers(objHeader As range) As Variant
Dim cell As range
Dim tempArray() As Variant
X = objHeader.EntireColumn.Count - 1
ReDim tempArray(X): i = 0
For Each cell In objHeader.Cells
    tempArray(i) = cell.Value
    i = i + 1
Next
Get_Headers = tempArray
End Function


Sub Read_Head2()
Dim arrHeader() As Variant, val As Variant
arrHeader = Application.Transpose(range("Session1").ListObject.HeaderRowRange)
For Each val In arrHeader
Debug.Print val
Next
MsgBox arrHeader(2, 1)
End Sub

Function Get_Sheetnames(ByRef ws As Workbook) As Variant
Dim i As Integer: i = 0
Dim arrays() As Variant
Dim Sheet As Worksheet
ReDim arrays(i)
For Each Sheet In ws.Worksheets
        arrays(i) = Sheet.name
        i = i + 1
        ReDim Preserve arrays(i)
Next
Get_Sheetnames = arrays
End Function

Public Function VerifyiedUser(ByVal atype As Activity, Optional sName As Variant = Null) As String
Static Get_VerifyCollection As Collection
If atype = A_on Then
    If Get_VerifyCollection Is Nothing And Not IsNull(sName) Then
            Set Get_VerifyCollection = New Collection
            Get_VerifyCollection.Add Item:=sName, key:=CStr(atype)
    End If
    VerifyiedUser = Get_VerifyCollection.Item(CStr(atype))
ElseIf atype = A_off Then
            Set Get_VerifyCollection = Nothing
            VerifyiedUser = Empty
End If
End Function




