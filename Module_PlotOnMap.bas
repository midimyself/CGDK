Attribute VB_Name = "Module_PlotOnMap"
Option Explicit

Function GetGPSValue(ByVal str As String) As Double
    Dim num() As String, ch() As String, i As Long, m As Long, n As Long, WrittingNum As Boolean, WrittingCh As Boolean
    If Len(str) = 0 Then GetGPSValue = 0: Exit Function
    If (VBA.Mid(str, 1, 1) >= "0" And VBA.Mid(str, 1, 1) <= 9) Or VBA.Mid(str, 1, 1) = "." Then
        ReDim Preserve num(1 To 1)
        m = 1
        WrittingNum = True
    Else
        ReDim Preserve ch(1 To 1)
        n = 1
        WrittingCh = True
    End If
    For i = 1 To Len(str)
        If (VBA.Mid(str, i, 1) >= "0" And VBA.Mid(str, i, 1) <= 9) Or VBA.Mid(str, i, 1) = "." Then
            If WrittingNum Then
                num(m) = num(m) & VBA.Mid(str, i, 1)
            Else
                m = m + 1
                ReDim Preserve num(1 To m)
                num(m) = num(m) & VBA.Mid(str, i, 1)
                WrittingNum = True
                WrittingCh = False
            End If
        Else
            If WrittingNum Then
                n = n + 1
                ReDim Preserve ch(1 To n)
                ch(n) = ch(n) & VBA.Mid(str, i, 1)
                WrittingNum = False
                WrittingCh = True
            Else
                ch(n) = ch(n) & VBA.Mid(str, i, 1)
            End If
        End If
    Next i
    If Common.HowManyElementsInArray(num) = 0 Then
        GetGPSValue = 0
    ElseIf Common.HowManyElementsInArray(num) = 1 Then
        GetGPSValue = Val(num(1)) * 3600
    ElseIf Common.HowManyElementsInArray(num) = 2 Then
        GetGPSValue = Val(num(1)) * 3600 + Val(num(2)) * 60
    Else
        GetGPSValue = Val(num(1)) * 3600 + Val(num(2)) * 60 + Val(num(3))
    End If
    Dim nagetive As Boolean
    nagetive = False
    For i = 1 To Common.HowManyElementsInArray(ch)
        If InStr(1, VBA.LCase(ch(i)), "s") > 0 Or InStr(1, VBA.LCase(ch(i)), "w") > 0 Or InStr(1, VBA.LCase(ch(i)), "-") > 0 Then nagetive = True
    Next i
    If nagetive Then GetGPSValue = -GetGPSValue
End Function

Sub Draw()
    Dim i As Long, j As Long, calibratepointrange As New ShapeRange, s As Shape, crv As Curve, sp As SubPath, v As Variant
    Dim symbolshaperange As New ShapeRange, nameshaperange As New ShapeRange, utm As New UTMConverter
    
    Dim x1 As Double, y1 As Double, lat1 As Double, lon1 As Double
    
    Dim x2 As Double, y2 As Double, lat2 As Double, lon2 As Double
    
    Dim px As Double, py As Double, plat As Double, plon As Double
    
    'On Error GoTo finally
    
    For Each v In PlotOnMap.Latitudes
        If GetGPSValue(v) > 324000 Or GetGPSValue(v) < -288000 Then MsgBox "Some latitude data are out of range!": Exit Sub
    Next v
    
    For Each v In PlotOnMap.Longitudes
        If GetGPSValue(v) > 648000 Or GetGPSValue(v) < -648000 Then MsgBox "Some longitude data are out of range!": Exit Sub
    Next v
    
    Set calibratepointrange = Common.FindShapeByType(ActivePage.Shapes.All, "CalibratePoint")
    If calibratepointrange.Count <> 2 Then
       MsgBox "The map hasn't been calibrated legally. Please recalibrate it."
       Exit Sub
    End If
    If calibratepointrange.Item(1).Type <> cdrEllipseShape Or calibratepointrange.Item(2).Type <> cdrEllipseShape Then
       MsgBox "The map hasn't been calibrated legally. Please recalibrate it."
       Exit Sub
    End If
    
    'Set crv = CreateCurve(ActiveDocument)
    
    x1 = calibratepointrange.Item(1).Ellipse.CenterX
    x2 = calibratepointrange.Item(2).Ellipse.CenterX
    y1 = calibratepointrange.Item(1).Ellipse.CenterY
    y2 = calibratepointrange.Item(2).Ellipse.CenterY
    'MsgBox y2
    lat1 = Val(VBA.Split(calibratepointrange.Item(1).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(0)) / 3600
    lat2 = Val(VBA.Split(calibratepointrange.Item(2).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(0)) / 3600
    lon1 = Val(VBA.Split(calibratepointrange.Item(1).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(1)) / 3600
    lon2 = Val(VBA.Split(calibratepointrange.Item(2).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(1)) / 3600
    lat1 = utm.DegToRad(lat1)
    lat2 = utm.DegToRad(lat2)
    lon1 = utm.DegToRad(lon1)
    lon2 = utm.DegToRad(lon2)
    
    ActiveDocument.BeginCommandGroup
    Common.AutoRefresh False
    
    For i = 1 To PlotOnMap.Latitudes.Count
        plat = GetGPSValue(PlotOnMap.Latitudes.Item(i)) / 3600
        plon = GetGPSValue(PlotOnMap.Longitudes.Item(i)) / 3600
        plat = utm.DegToRad(plat)
        plon = utm.DegToRad(plon)
        If utm.GetPointXY(lat1, lon1, x1, y1, lat2, lon2, x2, y2, plat, plon, px, py) Then
            On Error Resume Next
            Set s = PlotOnMap.LegendSymbol.Duplicate()
            s.PositionX = px - s.SizeWidth / 2
            s.PositionY = py + s.SizeHeight / 2
            symbolshaperange.Add s
            If PlotOnMap.TextBox3.value <> "" Then 'name
                s.Name = PlotOnMap.Names.Item(i)
            Else
                s.Name = i
            End If
        End If
        'If i = 1 Then Set sp = crv.CreateSubPath(x, y) Else sp.AppendCurveSegment x, y 'line
    
        
    Next i
    
'    If DrawAt.CheckBox1.value Then  'draw polyline
'        If crv.Nodes.Count > 1 Then
'            If crv.Nodes.First.GetDistanceFrom(crv.Nodes.Last) < 0.00001 Then crv.Closed = True
'            allshaperange.Add ActiveLayer.CreateCurve(crv)
'        End If
'    End If
'
'    If DrawAt.CheckBox2.value Then 'draw spline
'        If crv.Nodes.Count > 1 Then
'            Dim k As Long
'            If crv.Nodes.First.GetDistanceFrom(crv.Nodes.Last) < 0.00001 Then crv.Closed = True
'            For k = 1 To crv.Nodes.Count
'                crv.Nodes(k).Type = cdrSmoothNode
'            Next k
'            allshaperange.Add ActiveLayer.CreateCurve(crv)
'        End If
'    End If

    symbolshaperange.OrderToFront
    Common.AutoRefresh True
    ActiveDocument.EndCommandGroup
    Exit Sub
finally:
    Common.AutoRefresh True
    ActiveDocument.EndCommandGroup
    MsgBox "The Excel file may be closed."
End Sub

Sub Draw1()
    Dim i As Long, j As Long, calibratepointrange As New ShapeRange, s As Shape, crv As Curve, sp As SubPath, v As Variant
    Dim symbolshaperange As New ShapeRange, nameshaperange As New ShapeRange, utm As New UTMConverter
    
    Dim x1 As Double, y1 As Double, lat1 As Double, lon1 As Double
    
    Dim x2 As Double, y2 As Double, lat2 As Double, lon2 As Double
    
    Dim px As Double, py As Double, plat As Double, plon As Double
    On Error GoTo finally
    For Each v In PlotOnMap.Latitudes
        If GetGPSValue(v) > 324000 Or GetGPSValue(v) < -288000 Then MsgBox "Some latitude data are out of range!": Exit Sub
    Next v
    
    For Each v In PlotOnMap.Longitudes
        If GetGPSValue(v) > 648000 Or GetGPSValue(v) < -648000 Then MsgBox "Some longitude data are out of range!": Exit Sub
    Next v
    
    Set calibratepointrange = Common.FindShapeByType(ActivePage.Shapes.All, "CalibratePoint")
    If calibratepointrange.Count <> 2 Then
       MsgBox "The map hasn't been calibrated legally. Please recalibrate it."
       Exit Sub
    End If
    If calibratepointrange.Item(1).Type <> cdrEllipseShape Or calibratepointrange.Item(2).Type <> cdrEllipseShape Then
       MsgBox "The map hasn't been calibrated legally. Please recalibrate it."
       Exit Sub
    End If
    
    Set crv = CreateCurve(ActiveDocument)
    
    x1 = calibratepointrange.Item(1).Ellipse.CenterX
    x2 = calibratepointrange.Item(2).Ellipse.CenterX
    y1 = calibratepointrange.Item(1).Ellipse.CenterY
    y2 = calibratepointrange.Item(2).Ellipse.CenterY
    'MsgBox y2
    lat1 = Val(VBA.Split(calibratepointrange.Item(1).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(0)) / 3600
    lat2 = Val(VBA.Split(calibratepointrange.Item(2).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(0)) / 3600
    lon1 = Val(VBA.Split(calibratepointrange.Item(1).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(1)) / 3600
    lon2 = Val(VBA.Split(calibratepointrange.Item(2).ObjectData("CalibrateInfo").value, ",", -1, vbTextCompare)(1)) / 3600
    lat1 = utm.DegToRad(lat1)
    lat2 = utm.DegToRad(lat2)
    lon1 = utm.DegToRad(lon1)
    lon2 = utm.DegToRad(lon2)
    
    ActiveDocument.BeginCommandGroup
    Common.AutoRefresh False
    
    For i = 1 To PlotOnMap.Latitudes.Count
        plat = GetGPSValue(PlotOnMap.Latitudes.Item(i)) / 3600
        plon = GetGPSValue(PlotOnMap.Longitudes.Item(i)) / 3600
        plat = utm.DegToRad(plat)
        plon = utm.DegToRad(plon)
        If utm.GetPointXY(lat1, lon1, x1, y1, lat2, lon2, x2, y2, plat, plon, px, py) Then
        If i = 1 Then Set sp = crv.CreateSubPath(px, py) Else sp.AppendLineSegment px, py 'line
        End If
    Next i
    

    If crv.Nodes.Count > 1 Then
        If crv.Nodes.First.GetDistanceFrom(crv.Nodes.Last) < 0.00001 Then crv.Closed = True
    End If

    Set s = ActiveLayer.CreateCurve(crv)
    s.OrderToFront
    Common.AutoRefresh True
    ActiveDocument.EndCommandGroup
    Exit Sub
finally:
    Common.AutoRefresh True
    ActiveDocument.EndCommandGroup
    MsgBox "The Excel file may be closed."
End Sub
