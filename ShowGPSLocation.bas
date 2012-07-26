Attribute VB_Name = "ShowGPSLocation"
Option Explicit

Public x1 As Double, y1 As Double, lat1 As Double, lon1 As Double
Public x2 As Double, y2 As Double, lat2 As Double, lon2 As Double
Public ismaprectified As Boolean

Private Type POINTAPI
     x As Long
     y As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long 'get cursor location
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI)

Sub GetScreenLocation(x As Long, y As Long)
    Dim p As POINTAPI
    GetCursorPos p
    x = p.x
    y = p.y
End Sub
Sub GetDocumentLocation(x As Double, y As Double)
    Dim p As POINTAPI
    GetCursorPos p
    ActiveWindow.ScreenToDocument p.x, p.y, x, y
End Sub

Public Sub TimeOutProcessor(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Dim x As Long, y As Long, longitude As String, latitude As String
    If Application.AppWindow.WindowState = cdrWindowMinimized Then Exit Sub
    GetScreenLocation x, y
    GPSLocation.Move x * 150 / 200 + 5, y * 150 / 200 + 15
    GetLongitudeAndLatitude longitude, latitude
    GPSLocation.Label1.Caption = latitude
    GPSLocation.Label2.Caption = longitude
End Sub

Public Sub GetLongitudeAndLatitude(longitude As String, latitude As String)
    Dim px As Double, py As Double, plat As Double, plon As Double
    
    Dim utm As New UTMConverter
    
    Dim calibratepointrange As New ShapeRange
    Set calibratepointrange = Common.FindShapeByType(ActivePage.Shapes.All, "CalibratePoint")
    If Not ismaprectified Then
       longitude = "NULL": latitude = "NULL"
       Exit Sub
    End If

    On Error GoTo veryend
    GetDocumentLocation px, py

    If utm.GetPointLatLon(lat1, lon1, x1, y1, lat2, lon2, x2, y2, px, py, plat, plon) Then
        If plat < 0 Then latitude = "S:": plat = -plat Else latitude = "N:"
        If plon < 0 Then longitude = "W": plon = -plon Else longitude = "E:"
        Dim dd As Double, mm As Double, ss As Double
        dd = VBA.Fix(utm.RadToDeg(plon))
        mm = VBA.Fix((utm.RadToDeg(plon) - VBA.Fix(utm.RadToDeg(plon))) * 60)
        ss = ((utm.RadToDeg(plon) - VBA.Fix(utm.RadToDeg(plon))) * 60 - VBA.Fix((utm.RadToDeg(plon) - VBA.Fix(utm.RadToDeg(plon))) * 60)) * 60
        longitude = longitude & " " & dd & "¡ã" & mm & "¡ä" & VBA.Format(ss, "0.0000") & "¡å"
        dd = VBA.Fix(utm.RadToDeg(plat))
        mm = VBA.Fix((utm.RadToDeg(plat) - VBA.Fix(utm.RadToDeg(plat))) * 60)
        ss = ((utm.RadToDeg(plat) - VBA.Fix(utm.RadToDeg(plat))) * 60 - VBA.Fix((utm.RadToDeg(plat) - VBA.Fix(utm.RadToDeg(plat))) * 60)) * 60
        latitude = latitude & " " & dd & "¡ã" & mm & "¡ä" & VBA.Format(ss, "0.0000") & "¡å"
    Else
        longitude = "NULL": latitude = "NULL"
    End If
    
    Exit Sub
veryend:
    longitude = "NULL": latitude = "NULL"
End Sub



