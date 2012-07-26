VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GPSLocation 
   Caption         =   "GPSLocation"
   ClientHeight    =   180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1965
   OleObjectBlob   =   "GPSLocation.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "GPSLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



' API function
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal crKey As Long, _
        ByVal bAlpha As Byte, _
        ByVal dwFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hmenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" ( _
        ByVal hmenu As Long, _
        ByVal nPosition As Long, _
        ByVal wFlags As Long) As Long
Private Declare Function DeleteMenu Lib "user32" ( _
        ByVal hmenu As Long, _
        ByVal nPosition As Long, _
        ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal bRevert As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function AnimateWindow Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal dwTime As Long, _
        ByVal dwFlags As Long) As Long
Private Declare Function MoveWindow Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Const GWL_STYLE = (-16)
Const GWL_EXSTYLE = (-20)
Const WS_MAXIMIZEBOX = &H10000
Const WS_MINIMIZEBOX = &H20000
Const WS_THICKFRAME = &H40000
Const WS_EX_LAYERED = &H80000
Const WS_SYSMENU = &H80000
Const WS_CAPTION = &HC00000
Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
Const LWA_ALPHA = &H2
Const MF_BYCOMMAND = &H0
Const MF_BYPOSITION = &H400&
Const MF_DISABLED = &H2&
Const MF_REMOVE = &H1000&
Const SC_CLOSE = &HF060
Const SC_MOVE = &HF010
Const WM_SYSCOMMAND = &H112
Const AW_ACTIVATE = &H20000
Const AW_BLEND = &H80000
 
Private Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
End Enum

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
 
Dim hWnd As Long
Dim lStyleOld As Long

Private Sub UserForm_Activate()
    If ShowGPSLocation.ismaprectified = False Then
        Unload Me
        MsgBox "The map hasn't been rectified."
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim x As Long, y As Long, rtn As Long
    hWnd = FindWindow(VBA.vbNullString, Me.Caption)
    ShowGPSLocation.GetScreenLocation x, y
    Me.left = x * 150 / 200 + 5
    Me.top = y * 150 / 200 + 15
    lStyleOld = GetWindowLong(hWnd, GWL_EXSTYLE)
    rtn = lStyleOld Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hWnd, 0, 180, LWA_ALPHA
    TitleBar False
    ShowGPSLocation.ismaprectified = False
    Dim utm As New UTMConverter
    
    Dim calibratepointrange As New ShapeRange
    Set calibratepointrange = Common.FindShapeByType(ActivePage.Shapes.All, "CalibratePoint")
    If calibratepointrange.Count <> 2 Then
       ShowGPSLocation.ismaprectified = False
       Exit Sub
    End If
    If calibratepointrange.Item(1).Type <> cdrEllipseShape Or calibratepointrange.Item(2).Type <> cdrEllipseShape Then
       ShowGPSLocation.ismaprectified = False
       Exit Sub
    End If
    
    ShowGPSLocation.x1 = calibratepointrange.Item(1).Ellipse.CenterX
    ShowGPSLocation.x2 = calibratepointrange.Item(2).Ellipse.CenterX
    ShowGPSLocation.y1 = calibratepointrange.Item(1).Ellipse.CenterY
    ShowGPSLocation.y2 = calibratepointrange.Item(2).Ellipse.CenterY
    
    ShowGPSLocation.lat1 = utm.DegToRad(Val(VBA.Split(calibratepointrange.Item(1).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(0)) / 3600)
    ShowGPSLocation.lat2 = utm.DegToRad(Val(VBA.Split(calibratepointrange.Item(2).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(0)) / 3600)
    ShowGPSLocation.lon1 = utm.DegToRad(Val(VBA.Split(calibratepointrange.Item(1).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(1)) / 3600)
    ShowGPSLocation.lon2 = utm.DegToRad(Val(VBA.Split(calibratepointrange.Item(2).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(1)) / 3600)
    ShowGPSLocation.ismaprectified = True
    
    TimerOn hWnd, 0.01
End Sub

Private Function TitleBar(ByVal bState As Boolean)
    Dim lStyle As Long
    Dim tR As RECT
 
    GetWindowRect hWnd, tR
    lStyle = GetWindowLong(hWnd, GWL_STYLE)

    If (bState) Then
        lStyle = lStyle Or WS_SYSMENU
        lStyle = lStyle Or WS_CAPTION
    Else
        Me.Caption = ""
        lStyle = lStyle And Not WS_SYSMENU
        lStyle = lStyle And Not WS_CAPTION
    End If
    SetWindowLong hWnd, GWL_STYLE, lStyle
    
    SetWindowPos hWnd, 0, tR.left, tR.top, tR.right - tR.left, tR.bottom - tR.top, _
        SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Function

Private Sub TimerOn(hWnd As Long, iSeconds As Long)
Dim iTimer As Long
iSeconds = iSeconds * 1000
iTimer = SetTimer(hWnd, 1, iSeconds, AddressOf ShowGPSLocation.TimeOutProcessor)
End Sub

Private Sub TimerOff(hWnd As Long)
Dim iKillTimer As Long
iKillTimer = KillTimer(hWnd, 1)
End Sub

Private Sub UserForm_Terminate()
    TimerOff hWnd
End Sub
