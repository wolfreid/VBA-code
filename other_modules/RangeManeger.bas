Attribute VB_Name = "RangeManeger"


Sub Ranges_Show()
Dim nm As name
  For Each nm In ActiveWorkbook.Names
    Debug.Print nm.name, nm.RefersTo
  Next nm
End Sub
'процедура для работы с элементом новостной ленты
Sub NewsRow_Filler(ByVal sName As String)
Dim uRange As New cPrjSysblocks
    uRange.IRibbon = sName
End Sub


