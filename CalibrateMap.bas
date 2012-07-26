Attribute VB_Name = "CalibrateMap"
Option Explicit
Public IsWattingForCoordinates As Boolean
Public IsCancel As Boolean
Sub Calibrate()
    If Common.IsWaitingForAction Then Exit Sub
    
    If Not ActiveDocument.DataFields.IsPresent("Type") Then ActiveDocument.DataFields.AddEx "Type", cdrDataTypeString, "", , , "0", False, False, False
    If Not ActiveDocument.DataFields.IsPresent("CalibrateInfo") Then ActiveDocument.DataFields.AddEx "CalibrateInfo", cdrDataTypeString, "", , , "0", False, False, False
    Dim calibratepointrange As New ShapeRange
    Set calibratepointrange = Common.FindShapeByType(ActivePage.Shapes.All, "CalibratePoint")
    If calibratepointrange.Count <> 0 Then ' the map has been calibrated
        If MsgBox("The map has been calibrated. Do you want to recalibrate it?", vbYesNo, "Guide for map calibration") = vbYes Then
            calibratepointrange.Unlock
            calibratepointrange.Delete
        Else
            Exit Sub
        End If
    End If
    If MsgBox("The software need two points to calibrate a map. You can left click on the page to place a point and then enter the point coordinates in the diaglog box." & _
        "If you are ready, click OK button to place the first point.", vbYesNo, "Guide for map calibration") = vbYes Then
    Else
        Exit Sub
    End If
    Dim x As Double, y As Double, Shift As Long, b As Boolean, point1 As Shape, point2 As Shape, lat As Double, lon As Double, lat1 As Double, lon1 As Double
try1:
    Common.IsWaitingForAction = True
    b = ActiveDocument.GetUserClick(x, y, Shift, 30, False, cdrCursorWinCross)
    Common.IsWaitingForAction = False
    If b Then If MsgBox("You haven't place the point. Do you want to try again?", vbYesNo, "Guide for map calibration") = vbYes Then GoTo try1 Else Exit Sub
    Set point1 = ActiveLayer.CreateEllipse2(x, y, 0.1)
    IsWattingForCoordinates = True
    IsCancel = False
    Calibrator.Show 0
    
    Do While IsWattingForCoordinates
        waitsec 0.01
    Loop
    
    If IsCancel Then
        On Error Resume Next
        Unload Calibrator
        point1.Delete
        MsgBox "The calibration has been terminated.", vbOKOnly, "Guide for map calibration"
        Exit Sub
    End If
    
    On Error GoTo veryend1
        point1.ObjectData.Item("Type").value = "CalibratePoint"
        If Calibrator.ComboBox1.value = "N" Then
            lat = Val(Calibrator.TextBox1.value) * 3600 + Val(Calibrator.TextBox3.value) * 60 + Val(Calibrator.TextBox4.value)
        Else
            lat = -(Val(Calibrator.TextBox1.value) * 3600 + Val(Calibrator.TextBox3.value) * 60 + Val(Calibrator.TextBox4.value))
        End If
        If Calibrator.ComboBox2.value = "E" Then
            lon = Val(Calibrator.TextBox6.value) * 3600 + Val(Calibrator.TextBox7.value) * 60 + Val(Calibrator.TextBox5.value)
        Else
            lon = -(Val(Calibrator.TextBox6.value) * 3600 + Val(Calibrator.TextBox7.value) * 60 + Val(Calibrator.TextBox5.value))
        End If
        point1.ObjectData.Item("CalibrateInfo").value = lat & "," & lon
        Unload Calibrator
    On Error GoTo 0
    
try2:
    Common.IsWaitingForAction = True
    b = ActiveDocument.GetUserClick(x, y, Shift, 30, False, cdrCursorWinCross)
    Common.IsWaitingForAction = False
    If b Then If MsgBox("You haven't place the point. Do you want to try again?", vbYesNo, "Guide for map calibration") = vbYes Then GoTo try2 Else point1.Delete: Exit Sub
    Set point2 = ActiveLayer.CreateEllipse2(x, y, 0.1)
    
    IsWattingForCoordinates = True
    IsCancel = False
    
    Calibrator.Show 0
    Calibrator.CommandButton1.Caption = "Finish"
    Calibrator.Label4.Caption = "Enter the coordinates for the first point. You can move the point to a more suitable place, but do not delete it. Click Finsh button to place the second point."
    Do While IsWattingForCoordinates
        waitsec 0.01
    Loop
    
    If IsCancel Then
        On Error Resume Next
        Unload Calibrator
        point1.Delete
        point2.Delete
        MsgBox "The calibration has been terminated.", vbOKOnly, "Guide for map calibration"
        Exit Sub
    End If
    
    On Error GoTo veryend1
        point2.ObjectData.Item("Type").value = "CalibratePoint"
        If Calibrator.ComboBox1.value = "N" Then
            lat1 = Val(Calibrator.TextBox1.value) * 3600 + Val(Calibrator.TextBox3.value) * 60 + Val(Calibrator.TextBox4.value)
        Else
            lat1 = -(Val(Calibrator.TextBox1.value) * 3600 + Val(Calibrator.TextBox3.value) * 60 + Val(Calibrator.TextBox4.value))
        End If
        If Calibrator.ComboBox2.value = "E" Then
            lon1 = Val(Calibrator.TextBox6.value) * 3600 + Val(Calibrator.TextBox7.value) * 60 + Val(Calibrator.TextBox5.value)
        Else
            lon1 = -(Val(Calibrator.TextBox6.value) * 3600 + Val(Calibrator.TextBox7.value) * 60 + Val(Calibrator.TextBox5.value))
        End If
        
        Dim zone1 As Double, zone2 As Double
        zone1 = VBA.Int((lon / 3600 + 180#) / 6) + 1
        zone2 = VBA.Int((lon1 / 3600 + 180#) / 6) + 1
        If zone1 <> zone2 Then
            Unload Calibrator
            point2.Delete
            MsgBox "The two calibration points should be in the same UTM zone. Please relocate the second point.", vbOKOnly, "Guide for map calibration"
            GoTo try2
        End If
        
        point2.ObjectData.Item("CalibrateInfo").value = lat1 & "," & lon1
        Unload Calibrator
'        point1.Outline.width = 0
'        point2.Outline.width = 0
'        point1.Fill.ApplyNoFill
'        point2.Fill.ApplyNoFill
'        point1.Locked = True
'        point2.Locked = True
'        point2.RemoveFromSelection
        Common.IsWaitingForAction = False
        Exit Sub
veryend1:
    MsgBox "You delete the first point.Please relocate it.", vbOKOnly, "Guide for map calibration"
    Resume try1
    Exit Sub
veryend2:
    MsgBox "You delete the second point.Please relocate it.", vbOKOnly, "Guide for map calibration"
    Resume try2
    Exit Sub
End Sub

Private Sub waitsec(ByVal sec As Double)
Dim sTimer As Date
sTimer = Timer
Do
DoEvents
Loop While VBA.Format((Timer - sTimer), "0.000") < sec
End Sub
