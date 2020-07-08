Attribute VB_Name = "States"
Option Explicit

'����� ������������ ���������� ��������� �������
Public Sub DefaultSystem()
    Let DS = Empty
    Let AT = Empty
    Set TitleInterCells = Nothing '� ���� ���������� �� ��������� ������� Userform, ����. ��� �����
    Set Session = Nothing '� ���� ���������� ��������� ������ ������� Userform
    Set objSetupForm = Nothing  '� ���� ���������� ��������� ������ ���������� ������� Userform
    Set objSetupData = Nothing ' � ���� ���������� ��������� ������ ���������� ������� Listobject, ���������� ������ ���� ��������� � ������ ���������
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


'3-������ ��������� ������
'Visible if=True - ��������
'Visible if=False - ������ �����
Sub View_Change(ByVal �������� As Boolean)
    Dim Sh As Worksheet: Application.ScreenUpdating = False
    TitleSheet.Visible = xlSheetVisible
    Select Case AL
        Case AL_Admin, AL_DEVELOPER, AL_UNKNOWN
            For Each Sh In ThisWorkbook.Worksheets
                 If Not Sh Is TitleSheet Then Sh.Visible = IIf(��������, xlSheetVisible, xlSheetVeryHidden)
            Next Sh
        Case AL_USER, AL_topuser
            For Each Sh In ThisWorkbook.Worksheets
            If Sh Is TitleSheet Then
                Sh.Visible = IIf(��������, xlSheetVisible, xlSheetVeryHidden)
            ElseIf Sh Is InputSheet Then
                Sh.Visible = IIf(��������, xlSheetVisible, xlSheetVeryHidden)
            ElseIf Sh Is Dashboard Then
                Sh.Visible = IIf(��������, xlSheetVisible, xlSheetVeryHidden)
            Else
                Sh.Visible = xlSheetVeryHidden
            End If
            Next Sh
        End Select
    If AL = AL_Admin Or AL = AL_DEVELOPER Then
        TitleSheet.Columns(14).EntireColumn.Hidden = Not ��������
    ElseIf AL = AL_USER Or AL = AL_topuser Then
        TitleSheet.Rows(9).EntireRow.Hidden = Not ��������
    Else
        TitleSheet.Columns(14).EntireColumn.Hidden = True
        TitleSheet.Rows(9).EntireRow.Hidden = True
    End If
    Application.ScreenUpdating = True
End Sub



