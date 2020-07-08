VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private conversation As String

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
If Success Then MsgBox "Прогресс в книге успешно сохранен  "
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Msg = MsgBox(Get_EnumSentence(S_Attemption), vbYesNo, "Make Decision before closing Program")
Select Case Msg
    Case vbYes
        conversation = ""
        NewsRow_Filler conversation: Cancel = False
        MsgBox (Get_EnumSentence(S_Goodbye))
        DefaultSystem
        ThisWorkbook.Saved = True: Application.Quit
    Case vbNo
        msg2 = MsgBox("Желаете сохранить сейчас нажав Ok и выйти, или вернутся в систему нажав Cancel?", _
            vbOKCancel, "Подтвеждение сохранения результатов")
        If msg2 = vbOK Then
            Timer_MsgBox
            ThisWorkbook.Save: DefaultSystem: Cancel = False: Application.Quit
        ElseIf msg2 = vbCancel Then
            Cancel = True
        End If
End Select
'Application.Calculation = xlAutomatic
End Sub


Private Sub Workbook_Open()
Application.Calculation = xlManual
'Application.DisplayAlerts = True
Point_start
If val(Application.Version) > 15 Then
        If ActiveWorkbook.AutoSaveOn Then ActiveWorkbook.AutoSaveOn = False
End If
    NewsRow_Filler Get_EnumSentence(S_Presenting)
    If GetNewAL >= AL_USER Then
        View_Change True
        Refresh_Formulas
        TitleSheet.Activate
    Else
        Call Point_start: Call Refresh_Formulas
    End If
    If AL >= AL_Admin Then MsgBox "Приветствую вас, Администратор", vbInformation
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not SaveAsUI Then
        If AL >= AL_USER Then
            MySave
        Else
        sometext = InputBox(Get_EnumSentence(S_Warning) & vbNewLine & _
            "Требуется кодовое слово для активации", "Проверка")
            If sometext = keyword Then
                MySave
                    Else
                    Cancel = True
            End If
        End If
    End If
End Sub
'    ' запрещаем сохранение файла под другим именем, SaveAsUi - сообщение сохранить как
'
'
Function MySave()
 On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ThisWorkbook.Save
    Application.EnableEvents = True
    View_Change AL >= AL_USER
    Application.ScreenUpdating = True
End Function
'   'End Function

Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
MsgBox Sh.name & "  was calculated"
End Sub
