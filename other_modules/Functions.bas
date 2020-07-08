Attribute VB_Name = "Functions"
Public Function ÷¬≈“«¿À»¬ »(ﬂ◊≈… ¿ As range) As Double
    ÷¬≈“«¿À»¬ » = ﬂ◊≈… ¿.Interior.Color
End Function


Public Function SumByColor(ColorSample As range) As Double
     Dim Sum As Double
     Application.Volatile True
      For Each cell In ColorSample
         If cell.Interior.Color = ColorSample.Interior.Color Then
             Sum = cell.Interior.ColorIndex
         End If
     Next cell
     SumByColor = Sum
 End Function

Public Function adressAct(pointCell As range) As String
Dim adtext As String
adtext = pointCell.Address(RowAbsolute:=False)
adressAct = adtext
End Function


Public Function SumByFormat(ColorFormat As range, IndexValue As Integer) As Double
      SumByFormat = ColorFormat.FormatConditions(IndexValue).Interior.ColorIndex
End Function

Function IsFormula(ByVal cell As range, Optional ShowFormula As Boolean = False)
    If ShowFormula Then
        If cell.HasFormula Then
            IsFormula = "‘ÓÏÛÎ‡: " & IIf(cell.HasArray, "{" & cell.FormulaLocal & "}", cell.FormulaLocal)
        Else
            IsFormula = "«Ì‡˜ÂÌËÂ: " & cell.Value
        End If
    Else
        IsFormula = cell.HasFormula
    End If
End Function

Public Function recreate_val(arg1 As Date) As Integer
Dim d As String
    d = CStr(arg1) 'string transform max value
    recreate_val = CInt(Left(d, Len(d) - 8))
End Function



Public Function UseCollection(coll As Collection) _
                        As Collection
    Set UseCollection = coll
End Function

Public Function curdaysinyear() As Integer
Dim dtLastDayInYear As Date
Dim dt As Date: dt = Date
dtLastDayInYear = DateSerial(Year(dt), 12, 31)
curdaysinyear = DateDiff("d", CDate("1/1/" & Year(dt)), dtLastDayInYear) + 1
End Function

Public Function GetGUID() As String
    GetGUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
End Function

Function InRange(Range1 As range, Range2 As range) As Boolean
' returns True if Range1 is within Range2
Dim InterSectRange As range
Set InterSectRange = Application.Intersect(Range1, Range2)
InRange = Not InterSectRange Is Nothing
Set InterSectRange = Nothing
End Function


Sub TestInRange()
If InRange(ActiveCell, range("A1:D100")) Then
' code to handle that the active cell is within the right range
MsgBox "Active Cell In Range!"""
Else
' code to handle that the active cell is not within the right range
MsgBox "Active Cell NOT In Range!"""
End If
End Sub

Function GetUserName() As String
    GetUserName = Environ$("username")
    'or
    'GetUserName = Application.UserName
End Function

Function Get_Sheetnames_array() As Variant
Dim i As Integer: i = 0
Dim arrays() As Variant
Dim Sheet As Worksheet
ReDim arrays(i)
For Each Sheet In ThisWorkbook.Worksheets
        arrays(i) = Sheet.name
        i = i + 1
        ReDim Preserve arrays(i)
Next
Get_Sheetnames_array = arrays
End Function


Public Function Get_Headers_array(sName As String) As Variant
Dim cell As range
Dim objHeader As range
Dim tempArray() As Variant
Set objHeader = range(sName).ListObject.HeaderRowRange
X = objHeader.EntireColumn.Count - 1
ReDim tempArray(X): i = 0
For Each cell In objHeader.Cells
    'Debug.Print cell.Value
    tempArray(i) = cell.Value
    i = i + 1
Next
Get_Headers_array = tempArray
End Function

