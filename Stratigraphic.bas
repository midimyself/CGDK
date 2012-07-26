Attribute VB_Name = "Stratigraphic"
Option Explicit

Sub Draw()
    'On Error GoTo veryend
    
    Dim border As Shape
    Dim sr1 As New ShapeRange, sr2 As New ShapeRange
    Set border = ActiveSelectionRange.Shapes.Item(1)
    Dim i As Long, TotalThickness As Double, MaxWidth As Double, CurrentThickness As Double, v As Variant, crv As Curve, sp As SubPath, MinWidth As Double

    TotalThickness = Stratigraphic.GetTotalThickness

    If TotalThickness = 0 Then MsgBox "The total thickness can should be larger than zero.": Exit Sub
    MaxWidth = Stratigraphic.GetMaxWidth
    MinWidth = Stratigraphic.GetMinWidth / 10
    If MaxWidth < 0.00001 Then MaxWidth = 0.00001
    If MinWidth < 0.000001 Then MinWidth = 0.000001

    CurrentThickness = 0
    Set crv = CreateCurve(ActiveDocument)

    ActiveDocument.BeginCommandGroup
    Common.AutoRefresh False
    
    DrawStratigraphicColumns.legends.RemoveAll
    
    'draw legend list
    If DrawStratigraphicColumns.ComboBox1.value = "Common stratigraphic column" Or DrawStratigraphicColumns.ComboBox1.value = "Palo-geomagnetism column" Then
        If DrawStratigraphicColumns.TextBox3.value = "" Then GoTo end5
        Dim thecopyshaperange As New ShapeRange, thecopyshape As Shape
        'MsgBox DrawStratigraphicColumns.lithos.Count
        For i = 1 To DrawStratigraphicColumns.lithos.Count
            Set thecopyshape = DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.FindShape(Name:=DrawStratigraphicColumns.lithos.Item(i))
            If Not thecopyshape Is Nothing Then thecopyshaperange.Add thecopyshape
        Next i
        
        thecopyshaperange.Copy
        thecopyshaperange.RemoveAll
        
        ActiveSelectionRange.RemoveFromSelection
        'Common.AutoRefresh True
        ActiveLayer.Paste
        'Common.AutoRefresh False
        
        thecopyshaperange.AddRange ActiveSelectionRange
        
        Dim ssss As Shape, ts As Shape, tsc As Double
        'MsgBox thecopyshaperange.Count
        For i = 1 To thecopyshaperange.Count
            Set ssss = ActiveLayer.CreateRectangle2(border.PositionX + 3 / 2 * border.SizeWidth, border.PositionY - border.SizeWidth * 0.5 * 2 / 3 - border.SizeWidth * 0.5 * 2 / 3 * (i - 1) * 1.5, border.SizeWidth * 0.5, border.SizeWidth * 0.5 * 2 / 3)
            ssss.Fill.CopyAssign thecopyshaperange.Item(i).Fill
            DrawStratigraphicColumns.legends.AddElement ssss, thecopyshaperange.Item(i).Name
            On Error GoTo end4:
            Set ts = ActiveLayer.CreateArtisticText(0, 0, thecopyshaperange.Item(i).Name)
            tsc = ts.SizeHeight / ts.SizeWidth
            ts.SizeHeight = ssss.SizeHeight / 3 * 2
            ts.SizeWidth = ts.SizeHeight / tsc
            ts.PositionX = ssss.PositionX + ssss.SizeWidth * 3 / 2
            ts.PositionY = ssss.PositionY - (ssss.SizeHeight - ts.SizeHeight) / 2
            sr2.Add ts
end4:
            sr2.Add ssss
        Next i
        
        thecopyshaperange.Delete
end5:
    End If
    
    For i = 1 To DrawStratigraphicColumns.Thickness.Count
        Dim cx As Double, cy As Double, height As Double, width As Double, rectshape As Shape

        height = VBA.Abs(Val(DrawStratigraphicColumns.Thickness.Item(i))) / TotalThickness * border.SizeHeight
        If DrawStratigraphicColumns.TextBox2.Text = "" Then
            width = border.SizeWidth
        Else
            If DrawStratigraphicColumns.CheckBox1.value Then ' log width
                If VBA.Abs(Val(DrawStratigraphicColumns.Grainsize.Item(i))) < 0.000001 Then
                    width = 0.000001
                Else
                    width = (Log(VBA.Abs(Val(DrawStratigraphicColumns.Grainsize.Item(i)))) - Log(MinWidth)) / (Log(MaxWidth) - Log(MinWidth)) * border.SizeWidth
                End If
            Else 'normal width
                width = VBA.Abs(Val(DrawStratigraphicColumns.Grainsize.Item(i))) / MaxWidth * border.SizeWidth
            End If
        End If
'
        If DrawStratigraphicColumns.OptionButton1.value Then 'from bottom to top
                cx = border.PositionX
                cy = border.PositionY + CurrentThickness + height - border.SizeHeight

                If DrawStratigraphicColumns.ComboBox1.value = "Common stratigraphic column" And height > 0 Then
                    Set rectshape = ActiveLayer.CreateRectangle(cx, cy, cx + width, cy - height)
                    On Error Resume Next
                    If DrawStratigraphicColumns.TextBox3 <> "" Then
                        rectshape.Fill.CopyAssign DrawStratigraphicColumns.legends.GetElement(DrawStratigraphicColumns.Lith_Pos_Negative.Item(i).value).Fill
                    End If
                    On Error GoTo veryend
                    sr1.Add rectshape
                End If
'
                If DrawStratigraphicColumns.ComboBox1.value = "Palo-geomagnetism column" And height > 0 Then
                    Set rectshape = ActiveLayer.CreateRectangle(cx, cy, cx + width, cy - height)
                    On Error Resume Next
                    If DrawStratigraphicColumns.TextBox3 <> "" Then
                        rectshape.Fill.CopyAssign DrawStratigraphicColumns.legends.GetElement(DrawStratigraphicColumns.Lith_Pos_Negative.Item(i).value).Fill
                    End If
                    On Error GoTo veryend
                    rectshape.Selected = False
                    sr1.Add rectshape
                    If i > 1 And DrawStratigraphicColumns.TextBox3.value <> "" Then
                        If DrawStratigraphicColumns.Lith_Pos_Negative.Item(i).value = DrawStratigraphicColumns.Lith_Pos_Negative.Item(i - 1).value Then
                            sr1.Add sr1.LastShape.Weld(sr1.Shapes.Item(sr1.Count - 1), False, False)
                        End If
                    End If
                End If
'
                cx = border.PositionX + width
                cy = border.PositionY + CurrentThickness + height / 2 - border.SizeHeight
'
                If i = 1 Then
                    Set sp = crv.CreateSubPath(cx, cy)
                Else
                    sp.AppendCurveSegment cx, cy
                End If

            CurrentThickness = CurrentThickness + height
'
        Else 'from top to bottom
                cx = border.PositionX
                cy = border.PositionY - CurrentThickness

                If DrawStratigraphicColumns.ComboBox1.value = "Common stratigraphic column" And height > 0 Then
                    Set rectshape = ActiveLayer.CreateRectangle(cx, cy, cx + width, cy - height)
                    On Error Resume Next
                    If DrawStratigraphicColumns.TextBox3 <> "" Then
                        rectshape.Fill.CopyAssign DrawStratigraphicColumns.legends.GetElement(DrawStratigraphicColumns.Lith_Pos_Negative.Item(i).value).Fill
                    End If
                    On Error GoTo veryend
                    rectshape.Selected = False
                    sr1.Add rectshape
                End If
'
                If DrawStratigraphicColumns.ComboBox1.value = "Palo-geomagnetism column" And height > 0 Then
                    Set rectshape = ActiveLayer.CreateRectangle(cx, cy, cx + width, cy - height)
                    On Error Resume Next
                    If DrawStratigraphicColumns.TextBox3 <> "" Then
                        rectshape.Fill.CopyAssign DrawStratigraphicColumns.legends.GetElement(DrawStratigraphicColumns.Lith_Pos_Negative.Item(i).value).Fill
                    End If
                    On Error GoTo veryend
                    rectshape.Selected = False
                    sr1.Add rectshape
                    If i > 1 And DrawStratigraphicColumns.TextBox3.value <> "" Then
                        If DrawStratigraphicColumns.Lith_Pos_Negative.Item(i).value = DrawStratigraphicColumns.Lith_Pos_Negative.Item(i - 1).value Then
                            sr1.Add sr1.LastShape.Weld(sr1.Shapes.Item(sr1.Count - 1), False, False)
                        End If
                    End If
                End If
'

                cx = border.PositionX + width
                cy = border.PositionY - CurrentThickness - height / 2
'
                If i = 1 Then
                    Set sp = crv.CreateSubPath(cx, cy)
                Else
                    sp.AppendCurveSegment cx, cy
                End If
'
            CurrentThickness = CurrentThickness + height
'
        End If
    Next i

    On Error GoTo veryend

    If DrawStratigraphicColumns.ComboBox1.value = "Polyline column" Then 'polyline smoothline
        For i = 2 To crv.Nodes.Count - 1
            crv.Nodes(i).Type = cdrCuspNode
        Next i
        If crv.Nodes.Count > 1 Then ActiveLayer.CreateCurve(crv).Selected = False
    End If

    If DrawStratigraphicColumns.ComboBox1.value = "Smooth line column" Then 'draw smoothline
        For i = 2 To crv.Nodes.Count - 1
            crv.Nodes(i).Type = cdrSmoothNode
        Next i
        If crv.Nodes.Count > 1 Then ActiveLayer.CreateCurve(crv).Selected = False
    End If


    Common.AutoRefresh True
    
    If sr1.Count > 1 Then sr1.group
    If sr2.Count > 1 Then sr2.group
    ActiveSelectionRange.RemoveFromSelection
    border.Selected = True
    ActiveDocument.EndCommandGroup
    Exit Sub
veryend:
    Common.AutoRefresh True
    ActiveDocument.EndCommandGroup
    MsgBox "Can't draw the stratigraphic columns!"
End Sub

Function GetTotalThickness() As Double
If DrawStratigraphicColumns.TextBox1.value = "" Then GetTotalThickness = 0: Exit Function
Dim i As Long, total As Double
total = 0
For i = 1 To DrawStratigraphicColumns.Thickness.Count
    total = total + VBA.Abs(Val(DrawStratigraphicColumns.Thickness.Item(i)))
Next i
GetTotalThickness = total
End Function

Function GetMaxWidth() As Double
If DrawStratigraphicColumns.TextBox2.value = "" Then GetMaxWidth = 0: Exit Function
Dim i As Long, Max As Double
Max = Val(DrawStratigraphicColumns.Grainsize.Item(1))
For i = 2 To DrawStratigraphicColumns.Grainsize.Count
    If Max < VBA.Abs(Val(DrawStratigraphicColumns.Grainsize.Item(i))) Then Max = VBA.Abs(Val(DrawStratigraphicColumns.Grainsize.Item(i)))
Next i
GetMaxWidth = Max
End Function

Function GetMinWidth() As Double
If DrawStratigraphicColumns.TextBox2.value = "" Then GetMinWidth = 0.00001: Exit Function
Dim i As Long, Min As Double
Min = Val(DrawStratigraphicColumns.Grainsize.Item(1))
For i = 2 To DrawStratigraphicColumns.Grainsize.Count
    If Min > VBA.Abs(Val(DrawStratigraphicColumns.Grainsize.Item(i))) Then Min = VBA.Abs(Val(DrawStratigraphicColumns.Grainsize.Item(i)))
Next i
GetMinWidth = Min
End Function

Sub DrawYAxis()
    Dim s As Shape, a As String, v As Variant, s_x As Double, s_y As Double, height As Double, width As Double, textshape As Shape, textscale As Double
    Dim d As Double, x As Double, y As Double
    Dim left As Double, right As Double, top As Double, bottom As Double, xlog As Boolean, ylog As Boolean
    Set s = ActiveSelectionRange.Shapes(1)
    s.GetBoundingBox s_x, s_y, width, height
    
    If width > height Then d = height Else d = width
    
    If DrawStratigraphicColumns.OptionButton3.value Then
        top = Val(DrawStratigraphicColumns.TextBox5.value)
        bottom = Val(DrawStratigraphicColumns.TextBox4.value)
    Else
        bottom = Val(DrawStratigraphicColumns.TextBox5.value)
        top = Val(DrawStratigraphicColumns.TextBox4.value)
    End If

    Geochemistry.GetXYCoordinateInfo left, right, top, bottom, xlog, ylog

    Common.AutoRefresh False
    ActiveDocument.BeginCommandGroup
    
    For Each v In VBA.Split(DrawStratigraphicColumns.TextBox6.value, ",", -1, vbTextCompare)
        y = s_y + (Val(v) - bottom) / (top - bottom) * height
        Call ActiveLayer.CreateLineSegment(s_x, y, s_x - 0.06 * d, y)
        Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
        textscale = textshape.SizeHeight / textshape.SizeWidth
        textshape.SizeHeight = d * 0.06 * 2
        textshape.SizeWidth = textshape.SizeHeight / textscale
        textshape.PositionX = s_x - textshape.SizeWidth - textshape.SizeHeight / 2
        textshape.PositionY = y + textshape.SizeHeight / 2
    Next v
    ActiveSelectionRange.RemoveFromSelection
    s.Selected = True
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True

    Exit Sub
End Sub

Sub DrawXAxis()
    Dim s As Shape, a As String, v As Variant, s_x As Double, s_y As Double, height As Double, width As Double, textshape As Shape, textscale As Double
    Dim d As Double, x As Double, y As Double
    Dim left As Double, right As Double, top As Double, bottom As Double, xlog As Boolean, ylog As Boolean

    Set s = ActiveSelectionRange.Shapes(1)
    s.GetBoundingBox s_x, s_y, width, height

    
    If width > height Then d = height Else d = width
    
    If DrawStratigraphicColumns.CheckBox1.value Then
        left = Stratigraphic.GetMinWidth / 10
        right = Stratigraphic.GetMaxWidth
    Else
        left = 0
        right = Stratigraphic.GetMaxWidth
    End If

    Common.AutoRefresh False
    ActiveDocument.BeginCommandGroup
    
    For Each v In VBA.Split(DrawStratigraphicColumns.TextBox7, ",", -1, vbTextCompare)
        If DrawStratigraphicColumns.CheckBox1.value Then
            If Val(v) > 0.00001 Then
            x = s_x + (Log(Val(v)) - Log(left)) / (Log(right) - Log(left)) * width
            
            Call ActiveLayer.CreateLineSegment(x, s_y + s.SizeHeight, x, s_y + s.SizeHeight + 0.06 * d)
            Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
            textscale = textshape.SizeHeight / textshape.SizeWidth
            textshape.SizeHeight = d * 0.06 * 2
            textshape.SizeWidth = textshape.SizeHeight / textscale
            textshape.PositionX = x - textshape.SizeWidth / 2
            textshape.PositionY = s_y + s.SizeHeight + textshape.SizeHeight + 0.06 * d
            End If
        Else 'x is normal axis
            x = s_x + (Val(v) - left) / (right - left) * width
            Call ActiveLayer.CreateLineSegment(x, s_y + s.SizeHeight, x, s_y + s.SizeHeight + 0.06 * d)
            Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
            textscale = textshape.SizeHeight / textshape.SizeWidth
            textshape.SizeHeight = d * 0.06 * 2
            textshape.SizeWidth = textshape.SizeHeight / textscale
            textshape.PositionX = x - textshape.SizeWidth / 2
            textshape.PositionY = s_y + s.SizeHeight + textshape.SizeHeight + 0.06 * d
        End If
    Next v
    
    ActiveSelectionRange.RemoveFromSelection
    s.Selected = True
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
        
    Exit Sub
End Sub
