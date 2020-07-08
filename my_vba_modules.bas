Attribute VB_Name = "API"
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type Msg
    #If VBA7 Then
        hWnd As LongPtr
        message As Long
        wParam As LongPtr
        lParam As LongPtr
        time As Long
        pt As POINTAPI
    #Else
        hWnd As Long
        message As Long
        wParam As Long
        lParam As Long
        time As Long
        pt As POINTAPI
    #End If
End Type
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function WaitMessage Lib "user32" () As Long
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202


Public Function MouseMoveTest(ByRef inoMouse As cMouse) As Boolean
    Dim lngCurPos As POINTAPI
    Dim DocZero As POINTAPI
    Dim PointsPerPixelY As Double
    Dim PointsPerPixelX As Double
    Dim hdc As Long
    Dim tMsg As Msg
    Dim oRange As Object
    Dim RangeDetected As range
    MouseMoveTest = True
    hdc = GetDC(0)
    PointsPerPixelY = 72 / GetDeviceCaps(hdc, 90)
    PointsPerPixelX = 72 / GetDeviceCaps(hdc, 88)
    ReleaseDC 0, hdc
 
    DocZero.Y = ActiveWindow.PointsToScreenPixelsY(0)
    DocZero.X = ActiveWindow.PointsToScreenPixelsX(0)
 
    Do
        GetCursorPos lngCurPos
        Rowposition = (lngCurPos.Y - DocZero.Y) * PointsPerPixelY
        Colposition = (lngCurPos.X - DocZero.X) * PointsPerPixelX

       Set oRange = ActiveWindow.RangeFromPoint(lngCurPos.X, lngCurPos.Y)
            Call WaitMessage
            If PeekMessage(tMsg, Application.hWnd, WM_LBUTTONDOWN, WM_LBUTTONUP, 1) Then
                    If GetAsyncKeyState(VBA.vbKeyLButton) Then
                    Cells(1, 4) = Cells(1, 4) + 1
                        If Not oRange Is Nothing Then
                            If TypeName(oRange) = "Range" Then
                                Set RangeDetected = oRange
                                    inoMouse.Click RangeDetected
                            End If
                        End If
                    End If
            End If
    DoEvents
    Loop Until Not GlobalCancel
    MouseMoveTest = False
End Function



