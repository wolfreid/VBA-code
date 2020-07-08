Attribute VB_Name = "MarkerForms"

Const Leaving As String = "Leaving", Creating As String = "Creating"

'2-вызов окна ввода по кнопке из начального листа
Sub Authorization() 'кнопка
frmAuthorization.show
If AL = AL_UNKNOWN Then: Point_start: Else: View_Change True
Refresh_Formulas
TitleSheet.Activate
End Sub
Function GetNewAL() As AccessLevels
    frmAuthorization.show
    GetNewAL = AL
End Function

Sub Выход() 'кнопка
If AT = A_on Then
    ExistUser = VerifyiedUser(AT)
    Msg = MsgBox(Get_EnumSentence(S_Logout, ExistUser), vbYesNoCancel, "Logout") ' кнопка выхода на листе
    If Msg = vbYes Then
        InitExitConfiguration (ExistUser) ' запись в базу действия Выход
        AT = A_off
        AL = AL_UNKNOWN
        ExtactedUser = VerifyiedUser(AT)
        If IsEmpty(ExtractedUser) Then ExtactedUser = Destroy
        If frmAuthorization.Visible = False Then Call Point_start: Call Refresh_Formulas
    End If
Else
    Call Point_start: Call Refresh_Formulas
    Err.Raise Number:=vbObjectError + 513, _
              Description:="Missing code value for autorization procedure. Windows will be closed, any changes will destroyed "
    
    End
End If
End Sub
Sub InitExitConfiguration(sName As Variant)
'1 - получаем id действия
'2 - получаем имя действия
'3 - записываем в базу
Dim movecell As range

Set objSetupForm = New cConfigurations
Set Session = New cPrjSysblocks
objSetupForm.GetNameForm = Leaving
objSetupForm.activeForm = objSetupForm.KeepIdForm
DataCollection.LifeCollect sName, objSetupForm.activeForm
Set movecell = Session.MySession
movecell.ClearContents
'frmAutorization.rSession.ClearContents
End Sub

Sub NewProject()
Dim cell As range
Dim collProject As Collection
If filledspace Is Nothing Then MsgBox "range didn't set"
   For Each cell In filledspace.row
        If cell.column <> 3 Or cell.column <> 7 Then
        If cell.Value = "" Then GoTo DenyProс
        coll
   Next
End If
DenyProc:
MsgBox "Пустая ячейка, заполните все ключевые ячейки строки по текущему проекту"

End Sub

Sub NewProject_Record()

End Sub

