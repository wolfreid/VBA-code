Attribute VB_Name = "DataMake"
'здесь производ€тс€ уникальные данные - обьекты
Public rights As String

Declare Function CoCreateGuid Lib "ole32" (pguid As GUID) As Long
 
Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
 
Public Function CreateGuid() As String
Const S_OK As Long = 0

Dim GuidPointer As GUID
Dim strData4 As String
Dim i As Byte
 
strData4 = ""
If CoCreateGuid(GuidPointer) = S_OK Then
    With GuidPointer
        CreateGuid = Hex(.Data1) & Hex(.Data2) & Hex(.Data3)
        For i = 0 To 7
            strData4 = strData4 & Hex(.Data4(i))
        Next i
    End With
End If
CreateGuid = CreateGuid & strData4
 
End Function




