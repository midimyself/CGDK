Attribute VB_Name = "API"
'find a window
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'hide a window
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long '控制窗口的可见性
Public Const SW_HIDE = 0
'top a window
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST& = -1
Public Const HWND_TOP& = -2
Public Const SWP_NOSIZE& = &H1
Public Const SWP_NOMOVE& = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_HIDEWINDOW = &H80
'activate window
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

'set timmer

Private ITimerID As Long
Private ITimerID1 As Long

Public Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Sub StartTimer(IDuration As Long)
    If ITimerID = 0 Then
        ITimerID = SetTimer(0&, 0&, IDuration, AddressOf Geochemistry.Ontime)
    Else
        Call StopTimer
        ITimerID = SetTimer(0&, 0&, IDuration, AddressOf Geochemistry.Ontime)
    End If
End Sub

Sub StopTimer()
    KillTimer 0&, ITimerID
End Sub

Sub StartTimer1(IDuration As Long)
    If ITimerID1 = 1 Then
        ITimerID1 = SetTimer(0&, 1&, IDuration, AddressOf Geochemistry.Ontime1)
    Else
        Call StopTimer1
        ITimerID1 = SetTimer(0&, 1&, IDuration, AddressOf Geochemistry.Ontime1)
    End If
End Sub

Sub StopTimer1()
    KillTimer 0&, ITimerID1
End Sub



