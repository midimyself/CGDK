VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calibrator 
   Caption         =   "Guide for map calibration"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   OleObjectBlob   =   "Calibrator.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Calibrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const WS_SYSMENU = &H80000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const GWL_STYLE = (-16)


Private Sub CommandButton1_Click()
    If Me.ComboBox1.value = "N" And VBA.Abs(Val(Calibrator.TextBox1.value) * 3600# + Val(Calibrator.TextBox3.value) * 60# + Val(Calibrator.TextBox4.value)) > 84# * 3600# Then MsgBox "The latitude is out of range.": Exit Sub
    If Me.ComboBox1.value = "S" And VBA.Abs(Val(Calibrator.TextBox1.value) * 3600# + Val(Calibrator.TextBox3.value) * 60# + Val(Calibrator.TextBox4.value)) > 80# * 3600# Then MsgBox "The latitude is out of range.": Exit Sub
    If Me.ComboBox2.value = "E" And VBA.Abs(Val(Calibrator.TextBox1.value) * 3600# + Val(Calibrator.TextBox3.value) * 60# + Val(Calibrator.TextBox4.value)) > 180# * 3600# Then MsgBox "The latitude is out of range.": Exit Sub
    If Me.ComboBox2.value = "W" And VBA.Abs(Val(Calibrator.TextBox1.value) * 3600# + Val(Calibrator.TextBox3.value) * 60# + Val(Calibrator.TextBox4.value)) > 180# * 3600# Then MsgBox "The latitude is out of range.": Exit Sub
    CalibrateMap.IsWattingForCoordinates = False
End Sub

Private Sub CommandButton2_Click()
    CalibrateMap.IsWattingForCoordinates = False
    CalibrateMap.IsCancel = True
End Sub

Private Sub UserForm_Initialize()
Dim TempLng As Long, h As Long
h = API.FindWindow(VBA.vbNullString, Me.Caption)
TempLng = GetWindowLong(h, GWL_STYLE)
TempLng = TempLng And Not WS_SYSMENU
SetWindowLong h, GWL_STYLE, TempLng
Me.ComboBox1.AddItem "N"
Me.ComboBox1.AddItem "S"
Me.ComboBox2.AddItem "E"
Me.ComboBox2.AddItem "W"
End Sub
