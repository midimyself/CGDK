Attribute VB_Name = "Projection"
Option Explicit

Const PI = 3.14159265358979

Function QuadrantAngle(Azimuth As Double) As Double
    While Azimuth >= 360
        Azimuth = Azimuth - 360
    Wend
    While Azimuth < 0
        Azimuth = Azimuth + 360
    Wend
    If Azimuth >= 0 And Azimuth < 90 Then
        QuadrantAngle = 90 - Azimuth
    Else
        QuadrantAngle = 450 - Azimuth
    End If
    If QuadrantAngle = 360 Then QuadrantAngle = 0
End Function

Sub DrawPolarStereographicProjection(boundary As Shape)
    Dim sr As New ShapeRange, s As Shape, size1 As Double, size2 As Double, i As Long
    If DrawProjection.OptionButton1.value Then 'planar structure
        ActiveDocument.BeginCommandGroup
        Common.AutoRefresh False
        sr.Add boundary.Duplicate
        For i = 1 To DrawProjection.Azimuths.Count
            On Error GoTo finally
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double
            boundary.Ellipse.GetRadius size1, size2
            x1 = boundary.Ellipse.CenterX
            y1 = boundary.Ellipse.CenterY + size1
            x2 = boundary.Ellipse.CenterX + size2 * Tan((45 - DrawProjection.Dips.Item(i) / 2) * PI / 180)
            y2 = boundary.Ellipse.CenterY
            x3 = boundary.Ellipse.CenterX
            y3 = boundary.Ellipse.CenterY - size1
            Set s = ThreePointsCircle(x1, y1, x2, y2, x3, y3)
            s.ConvertToCurves
            s.Curve.Closed = False
            Set s = boundary.Intersect(s, True, False)
            s.SetRotationCenter boundary.Ellipse.CenterX, boundary.Ellipse.CenterY
            s.Rotate QuadrantAngle(DrawProjection.Azimuths.Item(i))
            s.Name = "Planer" & i & " " & DrawProjection.Azimuths.Item(i) & "буб╧" & DrawProjection.Dips.Item(i) & "бу"
            s.Selected = False
            sr.Add s
            If DrawProjection.CheckBox1.value Then ' draw polar with planer
                Dim size3 As Double
                size3 = size1 / 50
                Set s = ActiveLayer.CreateEllipse2(0, 0, size3)
                s.Ellipse.CenterX = boundary.Ellipse.CenterX + size2 * Tan((45 - (90 - DrawProjection.Dips(i)) / 2) * PI / 180)
                s.Ellipse.CenterY = boundary.Ellipse.CenterY
                s.SetRotationCenter boundary.Ellipse.CenterX, boundary.Ellipse.CenterY
                s.Rotate QuadrantAngle(DrawProjection.Azimuths.Item(i) + 180)
                s.Fill.ApplyUniformFill CreateCMYKColor(0, 0, 0, 100)
                s.Name = "Polar" & i & " " & DrawProjection.Azimuths.Item(i) & "буб╧" & DrawProjection.Dips.Item(i) & "бу"
                s.Selected = False
                sr.Add s
            End If
        Next i
        Common.AutoRefresh True
        sr.group.Selected = False
        boundary.Selected = True
        ActiveDocument.EndCommandGroup
    Else 'linear structure
        ActiveDocument.BeginCommandGroup
        Common.AutoRefresh False
        sr.Add boundary.Duplicate
        For i = 1 To DrawProjection.Azimuths.Count
            On Error GoTo finally
            boundary.Ellipse.GetRadius size1, size2
            size1 = size1 / 50
            Set s = ActiveLayer.CreateEllipse2(0, 0, size1)
            s.Ellipse.CenterX = boundary.Ellipse.CenterX + size2 * Tan((45 - DrawProjection.Dips(i) / 2) * PI / 180)
            s.Ellipse.CenterY = boundary.Ellipse.CenterY
            s.SetRotationCenter boundary.Ellipse.CenterX, boundary.Ellipse.CenterY
            s.Rotate QuadrantAngle(DrawProjection.Azimuths.Item(i))
            s.Fill.ApplyUniformFill CreateCMYKColor(0, 0, 0, 100)
            s.Name = "Linear" & i & " " & DrawProjection.Azimuths.Item(i) & "буб╧" & DrawProjection.Dips.Item(i) & "бу"
            s.Selected = False
            sr.Add s
        Next i
        Common.AutoRefresh True
        sr.group.Selected = False
        boundary.Selected = True
        ActiveDocument.EndCommandGroup
    End If
    Exit Sub
finally:
    Common.AutoRefresh True
    ActiveDocument.EndCommandGroup
    MsgBox "The Excel file may be closed."
End Sub

Function ThreePointsCircle(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double) As Shape
    Dim x As Single, y As Single, R As Single
    x = ((y3 - y1) * (y2 * y2 - y1 * y1 + x2 * x2 - x1 * x1) + (y2 - y1) * (y1 * y1 - y3 * y3 + x1 * x1 - x3 * x3)) / (2 * (x2 - x1) * (y3 - y1) - 2 * (x3 - x1) * (y2 - y1))
    y = ((x3 - x1) * (x2 * x2 - x1 * x1 + y2 * y2 - y1 * y1) + (x2 - x1) * (x1 * x1 - x3 * x3 + y1 * y1 - y3 * y3)) / (2 * (y2 - y1) * (x3 - x1) - 2 * (y3 - y1) * (x2 - x1))
    R = VBA.Sqr((x1 - x) * (x1 - x) + (y1 - y) * (y1 - y))
    Set ThreePointsCircle = ActiveLayer.CreateEllipse2(x, y, R)
End Function

Sub DrawRoseDiagram(boundary As Shape)
    Dim i As Long, strike() As Double, strikegroup As New DataGroup, azimuthsgroup As New DataGroup, dipgroup As New DataGroup
    On Error GoTo finally
    For i = 1 To DrawProjection.Azimuths.Count
        While DrawProjection.Azimuths.Item(i) >= 360
            DrawProjection.Azimuths.Item(i) = DrawProjection.Azimuths.Item(i) - 360
        Wend
        If DrawProjection.Azimuths.Item(i) = 360 Then DrawProjection.Azimuths.Item(i) = 0
        While DrawProjection.Azimuths.Item(i) < 0
            DrawProjection.Azimuths.Item(i) = DrawProjection.Azimuths.Item(i) + 360
        Wend
    Next i
    ReDim strike(1 To DrawProjection.Azimuths.Count)
    For i = 1 To DrawProjection.Azimuths.Count
        strike(i) = DrawProjection.Azimuths.Item(i) + 90
        If strike(i) > 90 And strike(i) < 180 Then strike(i) = strike(i) + 180
        If strike(i) >= 180 And strike(i) < 270 Then strike(i) = strike(i) - 180
        strikegroup.Add strike(i)
        azimuthsgroup.Add DrawProjection.Azimuths.Item(i)
        If DrawProjection.CheckBox4.value Then
            dipgroup.Add DrawProjection.Dips.Item(i)
        End If
    Next i
    Dim s As Shape, sr As New ShapeRange
    ActiveDocument.BeginCommandGroup
    sr.Add boundary.Duplicate
    If DrawProjection.CheckBox3.value Then
        Set s = DrawStrikeCurve(strikegroup, boundary)
        s.Name = "Strike"
        s.Selected = False
        sr.Add s
    End If
    If DrawProjection.CheckBox2.value Then
        Set s = DrawAzimuthsCurve(azimuthsgroup, boundary)
        s.Name = "Azimuth"
        s.Selected = False
        sr.Add s
    End If
    If DrawProjection.CheckBox4.value Then
        Set s = DrawDipCurve(azimuthsgroup, dipgroup, boundary)
        s.Name = "Dip/Plunge"
        s.Selected = False
        sr.Add s
    End If
    sr.group.Selected = False
    boundary.Selected = True
    ActiveDocument.EndCommandGroup
    Exit Sub
finally:
    Common.AutoRefresh True
    ActiveDocument.EndCommandGroup
    MsgBox "The Excel file may be closed."
End Sub

Private Function FindMaxNum(group As DataGroup) As Long
    Dim i As Double, Max As Long
    For i = 0 To 350 Step 10
        If group.DataNumInRange(i, i + 10) > Max Then Max = group.DataNumInRange(i, i + 10)
    Next i
    FindMaxNum = Max
End Function

Private Function DrawAzimuthsCurve(azimuthsgroup As DataGroup, boundary As Shape) As Shape
    Dim x(1 To 36) As Double, y(1 To 36) As Double, MaxNum As Long, rad As Double, i As Long, n
    Dim crv As Curve, sp As SubPath
    MaxNum = FindMaxNum(azimuthsgroup)
    boundary.Ellipse.GetRadius rad, rad
    n = 1
    On Error GoTo finally
    For i = 0 To 350 Step 10
        If azimuthsgroup.DataNumInRange(i, i + 10) > 0 Then
            x(n) = boundary.Ellipse.CenterX + rad * azimuthsgroup.DataNumInRange(i, i + 10) / MaxNum * Cos(QuadrantAngle(azimuthsgroup.DataMeanInRange(i, i + 10)) * PI / 180)
            y(n) = boundary.Ellipse.CenterY + rad * azimuthsgroup.DataNumInRange(i, i + 10) / MaxNum * Sin(QuadrantAngle(azimuthsgroup.DataMeanInRange(i, i + 10)) * PI / 180)
        Else
            x(n) = boundary.Ellipse.CenterX
            y(n) = boundary.Ellipse.CenterY
        End If
        n = n + 1
    Next i
    Set crv = CreateCurve(ActiveDocument)
    For i = 1 To 36
        If i = 1 Then Set sp = crv.CreateSubPath(x(i), y(i)) Else sp.AppendCurveSegment x(i), y(i)
    Next i
    Set DrawAzimuthsCurve = ActiveLayer.CreateCurve(crv)
    DrawAzimuthsCurve.Curve.Closed = True
    Exit Function
finally:
    Common.AutoRefresh True
    ActiveDocument.EndCommandGroup
    MsgBox "The Excel file may be closed."
End Function

Private Function DrawStrikeCurve(strikegroup As DataGroup, boundary As Shape) As Shape
    Dim x(1 To 36) As Double, y(1 To 36) As Double, MaxNum As Long, rad As Double, i As Long, n
    Dim crv As Curve, sp As SubPath
    On Error GoTo finally
    MaxNum = FindMaxNum(strikegroup)
    boundary.Ellipse.GetRadius rad, rad
    n = 1
    For i = 0 To 350 Step 10
        If strikegroup.DataNumInRange(i, i + 10) > 0 Then
            'MsgBox strikegroup.DataNumInRange(i, i + 10)
            'MsgBox MaxNum
            'MsgBox QuadrantAngle(strikegroup.DataMeanInRange(i, i + 10))
            'MsgBox "n=" & n
            x(n) = boundary.Ellipse.CenterX + rad * strikegroup.DataNumInRange(i, i + 10) / MaxNum * Cos(QuadrantAngle(strikegroup.DataMeanInRange(i, i + 10)) * PI / 180)
            y(n) = boundary.Ellipse.CenterY + rad * strikegroup.DataNumInRange(i, i + 10) / MaxNum * Sin(QuadrantAngle(strikegroup.DataMeanInRange(i, i + 10)) * PI / 180)
        Else
            x(n) = boundary.Ellipse.CenterX
            y(n) = boundary.Ellipse.CenterY
        End If
        n = n + 1
    Next i
    Set crv = CreateCurve(ActiveDocument)
    For i = 1 To 36
        If i = 1 Then Set sp = crv.CreateSubPath(x(i), y(i)) Else sp.AppendCurveSegment x(i), y(i)
    Next i
    Set DrawStrikeCurve = ActiveLayer.CreateCurve(crv)
    DrawStrikeCurve.Curve.Closed = True
    Exit Function
finally:
    Common.AutoRefresh True
    ActiveDocument.EndCommandGroup
    MsgBox "The Excel file may be closed."
End Function

Private Function DrawDipCurve(azimuthsgroup As DataGroup, dipgroup As DataGroup, boundary As Shape) As Shape
    Dim x(1 To 36) As Double, y(1 To 36) As Double, rad As Double, i As Long, n
    Dim crv As Curve, sp As SubPath
    On Error GoTo finally
    boundary.Ellipse.GetRadius rad, rad
    n = 1
 
    For i = 0 To 350 Step 10
        If azimuthsgroup.DataNumInRange(i, i + 10) > 0 Then
            x(n) = boundary.Ellipse.CenterX + rad * azimuthsgroup.LinkedGroupDataMeanInRange(dipgroup, i, i + 10) / 90 * Cos(QuadrantAngle(azimuthsgroup.DataMeanInRange(i, i + 10)) * PI / 180)
            y(n) = boundary.Ellipse.CenterY + rad * azimuthsgroup.LinkedGroupDataMeanInRange(dipgroup, i, i + 10) / 90 * Sin(QuadrantAngle(azimuthsgroup.DataMeanInRange(i, i + 10)) * PI / 180)
        Else
            x(n) = boundary.Ellipse.CenterX
            y(n) = boundary.Ellipse.CenterY
        End If
        n = n + 1
    Next i
    Set crv = CreateCurve(ActiveDocument)
    For i = 1 To 36
        If i = 1 Then Set sp = crv.CreateSubPath(x(i), y(i)) Else sp.AppendCurveSegment x(i), y(i)
    Next i
    Set DrawDipCurve = ActiveLayer.CreateCurve(crv)
    DrawDipCurve.Curve.Closed = True
    Exit Function
finally:
    Common.AutoRefresh True
    ActiveDocument.EndCommandGroup
    MsgBox "The Excel file may be closed."
End Function
