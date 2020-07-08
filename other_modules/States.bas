Attribute VB_Name = "States"
Option Explicit

'Здесь возвращаются устойчивые состояния системы
Public Sub DefaultSystem()
    Let DS = Empty
    Let AT = Empty
    Set TitleInterCells = Nothing 'К этой переменной не относятся обьекты Userform, прим. для полей
    Set Session = Nothing 'К этой переменной относятся только обьекты Userform
    Set objSetupForm = Nothing  'К этой переменной относятся только инстантные обьекты Userform
    Set objSetupData = Nothing ' К этой переменной относятся только инстантные обьекты Listobject, существует только один экзмепляр в рамках процедуры
End Sub

Public Sub Point_start()
    View_Change False
    NewsRow_Filler Get_EnumSentence(S_Presenting)
    Authorization_shapes Destroy
    DefaultSystem
End Sub

Public Sub Refresh_Formulas()
    TitleSheet.Calculate
    Application.Worksheets("Title").range("J8:M8").Dirty
    Worksheets("Title").range("J8:M8").Calculate
End Sub


'3-задаем видимость листов
'Visible if=True - показать
'Visible if=False - Скрыть листы
Sub View_Change(ByVal Показать As Boolean)
    Dim Sh As Worksheet: Application.ScreenUpdating = False
    TitleSheet.Visible = xlSheetVisible
    Select Case AL
        Case AL_Admin, AL_DEVELOPER, AL_UNKNOWN
            For Each Sh In ThisWorkbook.Worksheets
                 If Not Sh Is TitleSheet Then Sh.Visible = IIf(Показать, xlSheetVisible, xlSheetVeryHidden)
            Next Sh
        Case AL_USER, AL_topuser
            For Each Sh In ThisWorkbook.Worksheets
            If Sh Is TitleSheet Then
                Sh.Visible = IIf(Показать, xlSheetVisible, xlSheetVeryHidden)
            ElseIf Sh Is InputSheet Then
                Sh.Visible = IIf(Показать, xlSheetVisible, xlSheetVeryHidden)
            ElseIf Sh Is Dashboard Then
                Sh.Visible = IIf(Показать, xlSheetVisible, xlSheetVeryHidden)
            Else
                Sh.Visible = xlSheetVeryHidden
            End If
            Next Sh
        End Select
    If AL = AL_Admin Or AL = AL_DEVELOPER Then
        TitleSheet.Columns(14).EntireColumn.Hidden = Not Показать
    ElseIf AL = AL_USER Or AL = AL_topuser Then
        TitleSheet.Rows(9).EntireRow.Hidden = Not Показать
    Else
        TitleSheet.Columns(14).EntireColumn.Hidden = True
        TitleSheet.Rows(9).EntireRow.Hidden = True
    End If
    Application.ScreenUpdating = True
End Sub



