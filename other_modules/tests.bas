Attribute VB_Name = "tests"
Type Person
  name As String
  surname As String
  age As Long
End Type
Sub callsht()
Dim sht As Worksheet
Set sht = BehaviourSheet
MsgBox BehaviourSheet.name
End Sub



Sub GetUserName_Environ()
    Dim idx As Integer
    'To Directly the value of a Environment Variable with its Name
    MsgBox VBA.Interaction.Environ$("UserName")
    
    'To get all the List of Environment Variables
    For idx = 1 To 255
        strEnvironVal = VBA.Interaction.Environ$(idx)
        ThisWorkbook.Sheets(2).Cells(idx, 1) = strEnvironVal
    Next idx
    
End Sub

Sub test_XOR()
Const a As Integer = 10 ' 1010 in binary
Const b As Integer = 8  ' 1000 in binary
Const c As Integer = 6  ' 0110 in binary
Dim dasd As Byte
Dim word As String
Dim firstPattern, secondPattern, thirdPattern As Integer
firstPattern = (a Xor b)  '  2, 0010 in binary
secondPattern = (a Xor c) ' 12, 1100 in binary
thirdPattern = (b Xor c)  ' 14, 1110 in binary
word = "h"
Debug.Print firstPattern, secondPattern, thirdPattern
dasd = Asc(word)
'dasd = dasd
Debug.Print dasd
End Sub

Sub sadas()
End
End Sub

'Example usage of created Type

Private Sub test_enumarrays()
Dim life As Config_Lifedata: life = Life_array
Debug.Print life.Statename(0)
End Sub

Sub sadasd()
Debug.Print TitleSheet.Shapes(5).Type
End Sub


Sub sadasd2()
Dim col As New Collection
Dim col2 As Collection
For i = 1 To 10: col.Add i, ("hello" & i): Next
Set col2 = UseCollection(col)
Debug.Print col2("hello1")
End Sub



Private Sub testShape()
Dim testShape As shape
Set testShape = TitleSheet.Shapes.AddShape(msoShapeActionButtonCustom, 375, 50, 200, 50)
End Sub

Sub asdasd()
MsgBox Application.CountA(range("workspace").Rows(1))
'range("workspace").Resize(range("workspace").Rows.Count - 11, range("workspace").Columns.Count).Select
End Sub



Function LastSaveDate()
 Application.Volatile True
 LastSaveDate = FileDateTime(ThisWorkbook.FullName)
End Function



