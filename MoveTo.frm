VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MoveTo 
   Caption         =   "Move to"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3015
   OleObjectBlob   =   "MoveTo.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "MoveTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
    Dim calibratepointrange As New ShapeRange
    Dim x1 As Double, y1 As Double, lat1 As Double, lon1 As Double
    Dim x2 As Double, y2 As Double, lat2 As Double, lon2 As Double
    Dim utm As New UTMConverter
    Dim px As Double, py As Double, plat As Double, plon As Double
    plat = Val(Me.TextBox1.value) * 3600 + Val(Me.TextBox3.value) * 60 + Val(Me.TextBox4.value)
    plon = Val(Me.TextBox5.value) * 3600 + Val(Me.TextBox6.value) * 60 + Val(Me.TextBox7.value)
    If Me.ComboBox1.value = "S" Then plat = -plat
    If Me.ComboBox2.value = "W" Then plon = -plon
    If plat > 302400 Or plat < -288000 Then MsgBox "Can not move. The latitude is out of range": Exit Sub
    If plon > 648000 Or plon < -648000 Then MsgBox "Can not move. The longitude is out of range": Exit Sub
    'MsgBox "begin1"
    Set calibratepointrange = Common.FindShapeByType(ActivePage.Shapes.All, "CalibratePoint")
    If calibratepointrange.Count <> 2 Then
       MsgBox "The map hasn't been calibrated legally. Please recalibrate it."
       Exit Sub
    End If
    If calibratepointrange.Item(1).Type <> cdrEllipseShape Or calibratepointrange.Item(2).Type <> cdrEllipseShape Then
       MsgBox "The map hasn't been calibrated legally. Please recalibrate it."
       Exit Sub
    End If
    
    x1 = calibratepointrange.Item(1).Ellipse.CenterX
    x2 = calibratepointrange.Item(2).Ellipse.CenterX
    y1 = calibratepointrange.Item(1).Ellipse.CenterY
    y2 = calibratepointrange.Item(2).Ellipse.CenterY
    
    lat1 = Val(VBA.Split(calibratepointrange.Item(1).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(0)) / 3600
    lat2 = Val(VBA.Split(calibratepointrange.Item(2).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(0)) / 3600
    lon1 = Val(VBA.Split(calibratepointrange.Item(1).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(1)) / 3600
    lon2 = Val(VBA.Split(calibratepointrange.Item(2).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(1)) / 3600
    lat1 = utm.DegToRad(lat1)
    lat2 = utm.DegToRad(lat2)
    lon1 = utm.DegToRad(lon1)
    lon2 = utm.DegToRad(lon2)
    plat = utm.DegToRad(plat / 3600)
    plon = utm.DegToRad(plon / 3600)
    
    
    If Not utm.GetPointXY(lat1, lon1, x1, y1, lat2, lon2, x2, y2, plat, plon, px, py) Then MsgBox "Can not move. The position is not in the current UTM zone": Exit Sub
    
    'Call utm.GetPointXY(lat1, lon1, x1, y1, lat2, lon2, x2, y2, plat, plon, px, py)
    
    Dim s As Shape
    If ActiveSelectionRange.Count <> 1 Then MsgBox "You have to select a shape!": Exit Sub
    Set s = ActiveSelectionRange.Item(1)
    If s.Type = cdrCurveShape Then
        If s.Curve.Selection.Count = 1 Then
            s.Curve.Selection(1).PositionX = px
            s.Curve.Selection(1).PositionY = py
            Exit Sub
        End If
    End If
    Dim align As Long
    Common.MoveTo s, px, py
End Sub

Private Sub UserForm_Initialize()
    Me.ComboBox1.AddItem "N"
    Me.ComboBox1.AddItem "S"
    Me.ComboBox2.AddItem "E"
    Me.ComboBox2.AddItem "W"
End Sub
