Attribute VB_Name = "basIMPLODE"
Option Explicit
Private Const IDANI_CAPTION = &H3
Public Const SPI_GETWORKAREA = 48
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Sub ImplodeForm(hWnd As Long, Optional Reverse As Boolean = False, Optional IsFormCentered As Boolean = False)
    On Error Resume Next
    Dim f As RECT, i As RECT
    GetWindowRect hWnd, f
    If IsFormCentered = True Then CenterRect f
    i.Left = f.Left + (f.Right - f.Left) / 2
    i.Right = i.Left
    i.Top = f.Top + (f.Bottom - f.Top) / 2
    i.Bottom = i.Top
    If Not Reverse Then
        DrawAnimatedRects hWnd, IDANI_CAPTION, f, i
    Else
        DrawAnimatedRects hWnd, IDANI_CAPTION, i, f
    End If
End Sub
Public Sub ImplodeFormToMouse(hWnd As Long, Optional Reverse As Boolean = False, Optional IsFormCentered As Boolean = False)
    On Error Resume Next
    Dim f As RECT, i As RECT, p As POINTAPI
    GetWindowRect hWnd, f
    If IsFormCentered = True Then CenterRect f
    GetCursorPos p
    i.Left = p.x
    i.Right = p.x
    i.Top = p.y
    i.Bottom = p.y
    If Not Reverse Then
        DrawAnimatedRects hWnd, IDANI_CAPTION, f, i
    Else
        DrawAnimatedRects hWnd, IDANI_CAPTION, i, f
    End If
End Sub
Public Sub ImplodeFormToTray(hWnd As Long, Optional Reverse As Boolean = False, Optional IsFormCentered As Boolean = False)
    On Error Resume Next
    Dim f As RECT, i As RECT, p As POINTAPI
    GetWindowRect hWnd, f
    If IsFormCentered = True Then CenterRect f
    GetWindowRect GetTrayhWnd, i
    i.Left = i.Left + ((i.Right - i.Left) / 2)
    i.Right = i.Left
    If Not Reverse Then
        DrawAnimatedRects hWnd, IDANI_CAPTION, f, i
    Else
        DrawAnimatedRects hWnd, IDANI_CAPTION, i, f
    End If
End Sub
Private Function GetTrayhWnd() As Long
    On Error Resume Next
    Dim OurParent As Long
    Dim OurHandle As Long
    OurParent = FindWindow("Shell_TrayWnd", "")
    OurHandle = FindWindowEx(OurParent&, 0, "TrayNotifyWnd", vbNullString)
    GetTrayhWnd = OurHandle
End Function
Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function
Public Sub CenterRect(ByRef R As RECT)
    On Error Resume Next
    Dim h As Long, w As Long, tbh As Long, sw As Long, sh As Long
    h = R.Bottom - R.Top
    w = R.Right - R.Left
    tbh = GetTaskbarHeight / Screen.TwipsPerPixelY
    sw = Screen.Width / Screen.TwipsPerPixelX
    sh = (Screen.Height / Screen.TwipsPerPixelY) - tbh
    R.Left = (sw / 2) - (w / 2)
    R.Right = (sw / 2) + (w / 2)
    R.Top = (sh / 2) - (h / 2)
    R.Bottom = (sh / 2) + (h / 2)
End Sub

