Attribute VB_Name = "Forms"
Option Explicit
'Фабрика форм, здесб хранятся различнst формы
'Библиотека user32.dll подгружается для построения мимолетногосообщения
Private inShape As cShapes

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function MessageBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hWnd As Long, ByVal lpText As String, _
    ByVal lpCaption As String, ByVal uType As VbMsgBoxStyle, ByVal wLanguageId As Long, ByVal dwMilliseconds As Long) As Long
#Else
    Private Declare Function MessageBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hWnd As Long, ByVal lpText As String, _
    ByVal lpCaption As String, ByVal uType As VbMsgBoxStyle, ByVal wLanguageId As Long, ByVal dwMilliseconds As Long) As Long
#End If

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

Public Sub Timer_MsgBox()
    MessageBoxTimeOut Application.hWnd, "Подождите, данные сохраняются", "Уведомления", vbInformation + vbOKOnly, 0&, 1000
End Sub
Sub Generate_form()
Dim ClForm As cForms
Dim objBlock As cProjectBlocks
Dim val As Integer
Dim Cdata As ListObject
Dim inpvalue As Variant
On Error GoTo errHandler

hello:
Do
    Set ClForm = New cForms
    ClForm.GetDataCount = ClForm.UseProjectProperties()
    inpvalue = ClForm.Show_Form
    If Not IsEmpty(inpvalue) Then
        val = CInt(inpvalue)
        If val > ClForm.GetDataCount Then MsgBox "out of range, ubound value = " & ClForm.GetDataCount
    Else
        Exit Sub
    End If
Loop Until val <= ClForm.GetDataCount
Set objBlock = New cProjectBlocks
objBlock.ShowListData
MsgBox objBlock.Sheetnames(val)
objBlock.Data(val).DataBodyRange.Delete
Set objBlock = Nothing
Set ClForm = Nothing
Exit Sub
errHandler:
If Err.Number = 91 Then
    MsgBox "Nothing can do with some non existen table elements"
    Err.Clear
    Exit Sub
Else
MsgBox Err.Description & Err.Number
    Resume hello
End If
End Sub

Sub Authorization_shapes(ByVal login_n As Variant)
'В первом условии проверяем что форма
Static CloudShape As shape
Const sName As String = "CloudShape"
Set inShape = New cShapes
If login_n <> "" Then
    If CloudShape Is Nothing And AL > AL_UNKNOWN Then: Set CloudShape = inShape.Shape_Create() 'Create shape
    CloudShape.name = sName
    CloudShape.TextFrame.Characters.Text = Get_EnumSentence(S_Acceslevel, login_n)
Else
    If CloudShape Is Nothing And AL = AL_UNKNOWN Then: inShape.Delete_Shape sName: Exit Sub 'Delete saved nonstatic shape
    On Error GoTo 0
    CloudShape.Delete 'Delete static having shape
    Set CloudShape = Nothing
End If
End Sub

Sub LegalInform_Form()
Set TitleInterCells = New cPrjSysblocks
MsgBox Get_EnumLglInforms(TitleInterCells.LegalInform), vbInformation, "Access Description"
End Sub

Sub ShowModalDemo()
frmAdminInterface.show vbModeless
End Sub

Sub btnShtAddOrModifyAccounts_Click()
CallingAdminButton = TitleSheet.Shapes(Application.Caller).name
If GetNewAL <> AL_Admin Then Call Point_start: Call Refresh_Formulas
End Sub


Private Sub HookShape(ByVal shp As shape, ByVal Hook As Boolean)
    If Hook Then
        shp.AlternativeText = shp.AlternativeText & "**" & "Hooked"
    Else
        shp.AlternativeText = Replace(shp.AlternativeText, "**" & "Hooked", "")
    End If
End Sub

Public Sub SetShapesHook()
    'Add right-click macro to shapes 'Rectangle 1','Oval 1','Button 1'
    If GetProp(Application.hWnd, "HookSet") = 0 Then
        Call HookShape(TitleSheet.Shapes("Login_button"), True)
        Call HookShape(TitleSheet.Shapes("Logout_button"), True)
        Call SetProp(Application.hWnd, "HookSet", -1)
    Else
        MsgBox "Right-Click Macro already added to shapes."
    End If
End Sub





