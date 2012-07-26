Attribute VB_Name = "Geochemistry"
Option Explicit

Public DiagramCurve As Shape
Public CurrentDiagramType As String
Public TriangularSeries As New Collection
Public XYScatterSeries As New Collection
Public SpiderSeries As New Collection

Sub EstablishTriangularCoordinate(s As Shape)
    Dim d As Document
    Set d = ActiveDocument
    If Not d.DataFields.IsPresent("Type") Then d.DataFields.AddEx "Type", cdrDataTypeString, "", , , "0", False, False, False
    s.ObjectData("Type").value = "TriangularCoordinate"
End Sub

Sub Establish_XY_ScatterCoordinate(s As Shape, left As Double, right As Double, bottom As Double, top As Double, xlog As Boolean, ylog As Boolean)
    Dim d As Document
    Set d = ActiveDocument
    If Not d.DataFields.IsPresent("Type") Then d.DataFields.AddEx "Type", cdrDataTypeString, "", , , "0", False, False, False
    If Not d.DataFields.IsPresent("CoordinateInfo") Then d.DataFields.AddEx "CoordinateInfo", cdrDataTypeString, "", , , "0", False, False, False
    s.ObjectData("Type").value = "XYScatterCoordinate"
    s.ObjectData("CoordinateInfo").value = left & "," & right & "," & bottom & "," & top & "," & xlog & "," & ylog
End Sub

Sub EstablishSpiderCoordinate(s As Shape, bottom As Double, top As Double, ylog As Boolean, normalizedvalues As Object)
    Dim d As Document
    Set d = ActiveDocument
    If Not d.DataFields.IsPresent("Type") Then d.DataFields.AddEx "Type", cdrDataTypeString, "", , , "0", False, False, False
    If Not d.DataFields.IsPresent("CoordinateInfo") Then d.DataFields.AddEx "CoordinateInfo", cdrDataTypeString, "", , , "0", False, False, False
    s.ObjectData("Type").value = "SpiderCoordinate"
    s.ObjectData("CoordinateInfo").value = bottom & "," & top & "," & ylog
    If TemplateDesigner.TextBox5.value <> "" Then
        Dim i As Long, v As Variant
        For Each v In normalizedvalues
            If Val(v) = 0 Then
                s.ObjectData("CoordinateInfo").Clear
                s.ObjectData("Type").Clear
                MsgBox "Some illegal normalized data are detected!"
                TemplateDesigner.TextBox5.value = ""
                Exit For
            End If
            s.ObjectData("CoordinateInfo").value = s.ObjectData("CoordinateInfo").value & "," & Val(v)
        Next v
    End If
End Sub

Function HowManyColumns() As Long
    On Error GoTo veryend
    If CurrentDiagramType <> "SpiderCoordinate" Then HowManyColumns = 0: Exit Function
    Dim str As String, n As Long, i As Long
    str = DiagramCurve.ObjectData("CoordinateInfo").value
    For i = 1 To Len(str)
       If Mid(str, i, 1) = "," Then n = n + 1
    Next i
    HowManyColumns = n - 2
    Exit Function
veryend:
    HowManyColumns = -1
End Function

Function GetXYCoordinateInfo(left As Double, right As Double, top As Double, bottom As Double, xlog As Boolean, ylog As Boolean) As Boolean
    On Error GoTo veryend
    Dim str1 As String, str2 As String
    left = Val(VBA.Split(DiagramCurve.ObjectData("CoordinateInfo").value, ",", -1, 1)(0))
    right = Val(VBA.Split(DiagramCurve.ObjectData("CoordinateInfo").value, ",", -1, 1)(1))
    bottom = Val(VBA.Split(DiagramCurve.ObjectData("CoordinateInfo").value, ",", -1, 1)(2))
    top = Val(VBA.Split(DiagramCurve.ObjectData("CoordinateInfo").value, ",", -1, 1)(3))
    str1 = VBA.Split(DiagramCurve.ObjectData("CoordinateInfo").value, ",", -1, 1)(4)
    str2 = VBA.Split(DiagramCurve.ObjectData("CoordinateInfo").value, ",", -1, 1)(5)
    If VBA.Abs(left - right) = 0 Then GoTo veryend
    If VBA.Abs(top - bottom) = 0 Then GoTo veryend
    
    If VBA.LCase(str1) = "true" Then xlog = True Else xlog = False
    If VBA.LCase(str2) = "true" Then ylog = True Else ylog = False
    If xlog Then
        If left < 0 Or right < 0 Then GoTo veryend
    End If
    If ylog Then
        If bottom < 0 Or top < 0 Then GoTo veryend
    End If
    GetXYCoordinateInfo = True
    Exit Function
veryend:
    GetXYCoordinateInfo = False
End Function

Sub Ontime()
    Dim i As Integer, j As Integer, k As Integer
    i = Common.FindShapeByType(ActiveSelectionRange, "TriangularCoordinate").Count
    j = Common.FindShapeByType(ActiveSelectionRange, "XYScatterCoordinate").Count
    k = Common.FindShapeByType(ActiveSelectionRange, "SpiderCoordinate").Count
    If i > 0 Then
        Set DiagramCurve = FindShapeByType(ActiveSelectionRange, "TriangularCoordinate").Item(1)
        CurrentDiagramType = "TriangularCoordinate"
        PlotGeochemistryDiagram.Frame1.Visible = True
        PlotGeochemistryDiagram.Frame2.Visible = False
        PlotGeochemistryDiagram.Frame3.Visible = False
    ElseIf j > 0 Then
        Set DiagramCurve = FindShapeByType(ActiveSelectionRange, "XYScatterCoordinate").Item(1)
        CurrentDiagramType = "XYScatterCoordinate"
        PlotGeochemistryDiagram.Frame1.Visible = False
        PlotGeochemistryDiagram.Frame2.Visible = True
        PlotGeochemistryDiagram.Frame3.Visible = False
    ElseIf k > 0 Then
        Set DiagramCurve = FindShapeByType(ActiveSelectionRange, "SpiderCoordinate").Item(1)
        CurrentDiagramType = "SpiderCoordinate"
        PlotGeochemistryDiagram.Frame1.Visible = False
        PlotGeochemistryDiagram.Frame2.Visible = False
        PlotGeochemistryDiagram.Frame3.Visible = True
    End If
End Sub

Sub Ontime1()
    Dim i As Integer, j As Integer, k As Integer
    i = Common.FindShapeByType(ActiveSelectionRange, "TriangularCoordinate").Count
    j = Common.FindShapeByType(ActiveSelectionRange, "XYScatterCoordinate").Count
    k = Common.FindShapeByType(ActiveSelectionRange, "SpiderCoordinate").Count
    If i > 0 Then
        Set DiagramCurve = FindShapeByType(ActiveSelectionRange, "TriangularCoordinate").Item(1)
        CurrentDiagramType = "TriangularCoordinate"
        TemplateDesigner.Label8.Caption = "Top"
        TemplateDesigner.Label9.Caption = "Left"
        TemplateDesigner.Label10.Caption = "Right"
        TemplateDesigner.Label12.Caption = "Top"
        TemplateDesigner.Label13.Caption = "Left"
        TemplateDesigner.Label14.Caption = "Right"
        TemplateDesigner.TextBox8.Visible = True
        TemplateDesigner.TextBox6.Visible = True
        TemplateDesigner.TextBox7.Visible = True
        TemplateDesigner.TextBox9.Visible = True
        TemplateDesigner.TextBox10.Visible = True
        TemplateDesigner.TextBox11.Visible = True
        TemplateDesigner.CommandButton4.Visible = True
        TemplateDesigner.CommandButton6.Visible = True
        TemplateDesigner.CommandButton7.Visible = True
        TemplateDesigner.CommandButton8.Visible = True
        TemplateDesigner.Label11.Caption = "Triangular diagram"
        TemplateDesigner.Label15.Caption = "Triangular diagram"
    ElseIf j > 0 Then
        Set DiagramCurve = FindShapeByType(ActiveSelectionRange, "XYScatterCoordinate").Item(1)
        CurrentDiagramType = "XYScatterCoordinate"
        TemplateDesigner.Label8.Caption = "X Axis"
        TemplateDesigner.Label9.Caption = "Y Axis"
        TemplateDesigner.Label12.Caption = "X Values"
        TemplateDesigner.Label13.Caption = "Y Values"
        TemplateDesigner.Label10.Caption = ""
        TemplateDesigner.Label14.Caption = ""
        TemplateDesigner.TextBox8.Visible = False
        TemplateDesigner.TextBox6.Visible = True
        TemplateDesigner.TextBox7.Visible = True
        TemplateDesigner.TextBox11.Visible = False
        TemplateDesigner.TextBox10.Visible = True
        TemplateDesigner.TextBox9.Visible = True
        TemplateDesigner.CommandButton4.Visible = True
        TemplateDesigner.CommandButton6.Visible = True
        TemplateDesigner.CommandButton7.Visible = True
        TemplateDesigner.CommandButton8.Visible = True
        TemplateDesigner.Label11.Caption = "XYScatter diagram"
        TemplateDesigner.Label15.Caption = "XYScatter diagram"
    ElseIf k > 0 Then
        Set DiagramCurve = FindShapeByType(ActiveSelectionRange, "SpiderCoordinate").Item(1)
        CurrentDiagramType = "SpiderCoordinate"
        TemplateDesigner.Label8.Caption = "Normalized"
        TemplateDesigner.Label9.Caption = "Y Axis"
        TemplateDesigner.Label12.Caption = "Y Values"
        TemplateDesigner.Label13.Caption = ""
        TemplateDesigner.Label10.Caption = ""
        TemplateDesigner.Label14.Caption = ""
        TemplateDesigner.TextBox8.Visible = False
        TemplateDesigner.TextBox6.Visible = True
        TemplateDesigner.TextBox7.Visible = True
        TemplateDesigner.TextBox11.Visible = False
        TemplateDesigner.TextBox9.Visible = True
        TemplateDesigner.TextBox10.Visible = False
        TemplateDesigner.CommandButton4.Visible = True
        TemplateDesigner.CommandButton6.Visible = True
        TemplateDesigner.CommandButton7.Visible = True
        TemplateDesigner.CommandButton8.Visible = True
        TemplateDesigner.Label11.Caption = "Spider diagram"
        TemplateDesigner.Label15.Caption = "Spider diagram"
    End If
End Sub

Sub PlotTriangularDiagram()
    If CurrentDiagramType <> "TriangularCoordinate" Then Exit Sub
    Dim a() As Double, b() As Double, c() As Double, a1() As Double, b1() As Double, c1() As Double, ename() As String, _
        x() As Double, y() As Double, i As Long, j As Long, myseries As Series, s_y As Double, s_x As Double, height As Double, width As Double, _
        symbols As ShapeRange, connecting As Shape, seriesmarks As New ShapeRange, textshape As Shape
    If Not ActiveDocument.DataFields.IsPresent("Type") Then ActiveDocument.DataFields.AddEx "Type", cdrDataTypeString, "", , , "0", False, False, False
    If Not ActiveDocument.DataFields.IsPresent("SymbolInfo") Then ActiveDocument.DataFields.AddEx "SymbolInfo", cdrDataTypeString, "", , , "0", False, False, False
    On Error GoTo veryend1
        DiagramCurve.GetBoundingBox s_x, s_y, width, height
    On Error GoTo finally
    ActiveDocument.BeginCommandGroup
    For i = 1 To Geochemistry.TriangularSeries.Count
        Dim cr As Integer, cg As Integer, cb As Integer
        Randomize
        cr = Int(255 * Rnd) + 1
        Randomize
        cg = Int(255 * Rnd) + 1
        Randomize
        cb = Int(255 * Rnd) + 1
        Set myseries = Geochemistry.TriangularSeries.Item(i)
        
        ReDim a(1 To myseries.Dim1_Data.Count)
        ReDim b(1 To myseries.Dim1_Data.Count)
        ReDim c(1 To myseries.Dim1_Data.Count)
        ReDim ename(1 To myseries.Dim1_Data.Count)
        ReDim a1(1 To myseries.Dim1_Data.Count)
        ReDim b1(1 To myseries.Dim1_Data.Count)
        ReDim c1(1 To myseries.Dim1_Data.Count)
        ReDim x(1 To myseries.Dim1_Data.Count)
        ReDim y(1 To myseries.Dim1_Data.Count)
        
        Dim v As Variant, Index As Long
        Index = 1
        For Each v In myseries.Dim1_Data
            a(Index) = VBA.Abs(Val(v))
            Index = Index + 1
        Next v
        Index = 1
        For Each v In myseries.Dim2_Data
            b(Index) = VBA.Abs(Val(v))
            Index = Index + 1
        Next v
        Index = 1
        For Each v In myseries.Dim3_Data
            c(Index) = VBA.Abs(Val(v))
            Index = Index + 1
        Next v
        
        If Not myseries.ElementNames Is Nothing Then
            Index = 1
            For Each v In myseries.ElementNames
                ename(Index) = v
                Index = Index + 1
            Next v
        End If

        For j = 1 To myseries.Dim1_Data.Count
            a1(j) = a(j) / (a(j) + b(j) + c(j))
            b1(j) = b(j) / (a(j) + b(j) + c(j))
            c1(j) = c(j) / (a(j) + b(j) + c(j))
            y(j) = s_y + a1(j) * height
            x(j) = s_x + c1(j) * width + a1(j) * width / 2
        Next j
        
        Common.AutoRefresh False
        
        'draw shapes
        Dim s As Shape, d As Double
        If width > height Then d = height Else d = width
        Set s = ActiveLayer.CreateRectangle2(s_x + width + 5 * d * 0.03, s_y + height - 3 * i * d * 0.03, d * 0.03, d * 0.03)
        s.Name = "Legend"
        s.ObjectData("Type").value = "Symbol"
        s.Fill.ApplyUniformFill CreateRGBColor(cr, cg, cb)
        seriesmarks.Add s
        
        Set textshape = ActiveLayer.CreateArtisticText(0, 0, myseries.Name)
        Dim textscale As Double
        textscale = textshape.SizeHeight / textshape.SizeWidth
        textshape.SizeHeight = d * 0.03 * 2
        textshape.SizeWidth = textshape.SizeHeight / textscale
        textshape.PositionX = s_x + width + 10 * d * 0.03
        textshape.PositionY = s_y + height - 3 * i * d * 0.03 + textshape.SizeHeight / 2 + d * 0.03
        textshape.Name = "SeriesName"
        seriesmarks.Add textshape
        
        Set symbols = Common.DrawShapes(s, x, y)
        For j = 1 To symbols.Count
            If Not myseries.ElementNames Is Nothing Then symbols(j).Name = ename(j) Else symbols(j).Name = j
            symbols(j).ObjectData("Type").value = "Symbol"
            symbols(j).ObjectData("SymbolInfo").value = symbols(j).Name & "," & a(j) & "," & b(j) & "," & c(j)
        Next j
        seriesmarks.AddRange symbols
        
        'draw curve
        If PlotGeochemistryDiagram.TCheck1 = True Then
            Dim legendconnecting As Shape
            Set connecting = Common.DrawCurve(x, y)
            If Not connecting Is Nothing Then
                Set legendconnecting = ActiveLayer.CreateCurveSegment(s_x + width + d * 0.03, s_y + height - 3 * i * d * 0.03 + d * 0.03 / 2, s_x + width + 10 * d * 0.03, s_y + height - 3 * i * d * 0.03 + d * 0.03 / 2)
                connecting.ObjectData("Type").value = "ConnectingLine"
                connecting.Name = "ConnectingLine"
                legendconnecting.ObjectData("Type").value = "ConnectingLine"
                legendconnecting.Name = "LegendConnectingLine"
                connecting.Outline.Color.RGBAssign cr, cg, cb
                legendconnecting.Outline.Color.RGBAssign cr, cg, cb
                connecting.OrderToBack
                legendconnecting.OrderToBack
                seriesmarks.Add connecting
                seriesmarks.Add legendconnecting
            End If
        End If
        
        Common.AutoRefresh True
        
        If seriesmarks.Count > 1 Then seriesmarks.group.Name = myseries.Name
        
        seriesmarks.RemoveAll
        ActiveSelectionRange.RemoveFromSelection
    Next i
    ActiveDocument.EndCommandGroup
    
    Exit Sub
veryend1:
    MsgBox "The template may be removed from the current document"
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
    Exit Sub
finally:
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
    MsgBox "The Excel file may be closed."
End Sub

Sub PlotXYScatterDiagram()
    If CurrentDiagramType <> "XYScatterCoordinate" Then Exit Sub
    Dim a() As Double, b() As Double, ename() As String, _
        x() As Double, y() As Double, i As Long, j As Long, myseries As Series, s_y As Double, s_x As Double, height As Double, width As Double, _
        symbols As ShapeRange, connecting As Shape, seriesmarks As New ShapeRange, textshape As Shape
    If Not ActiveDocument.DataFields.IsPresent("Type") Then ActiveDocument.DataFields.AddEx "Type", cdrDataTypeString, "", , , "0", False, False, False
    If Not ActiveDocument.DataFields.IsPresent("SymbolInfo") Then ActiveDocument.DataFields.AddEx "SymbolInfo", cdrDataTypeString, "", , , "0", False, False, False
    On Error GoTo veryend1
        DiagramCurve.GetBoundingBox s_x, s_y, width, height
    On Error GoTo finally
    ActiveDocument.BeginCommandGroup

    For i = 1 To Geochemistry.XYScatterSeries.Count
        Dim cr As Integer, cg As Integer, cb As Integer
        Randomize
        cr = Int(255 * Rnd) + 1
        Randomize
        cg = Int(255 * Rnd) + 1
        Randomize
        cb = Int(255 * Rnd) + 1
        Set myseries = Geochemistry.XYScatterSeries.Item(i)
        ReDim a(1 To myseries.Dim1_Data.Count)
        ReDim b(1 To myseries.Dim1_Data.Count)
        ReDim x(1 To myseries.Dim1_Data.Count)
        ReDim y(1 To myseries.Dim1_Data.Count)
        ReDim ename(1 To myseries.Dim1_Data.Count)
        
        Dim v As Variant, Index As Long
        Index = 1
        For Each v In myseries.Dim1_Data
            a(Index) = Val(v)
            Index = Index + 1
        Next v
        Index = 1
        For Each v In myseries.Dim2_Data
            b(Index) = Val(v)
            Index = Index + 1
        Next v
        
        If Not myseries.ElementNames Is Nothing Then
            Index = 1
            For Each v In myseries.ElementNames
                ename(Index) = v
                Index = Index + 1
            Next v
        End If
        
        Dim left As Double, right As Double, top As Double, bottom As Double, xlog As Boolean, ylog As Boolean
        
        If Not Geochemistry.GetXYCoordinateInfo(left, right, top, bottom, xlog, ylog) Then GoTo veryend2
        If VBA.Abs(left - right) = 0 Then GoTo veryend2
        If VBA.Abs(top - bottom) = 0 Then GoTo veryend2
        If xlog Then
            If left <= 0 Or right <= 0 Then GoTo veryend2
        End If
        If ylog Then
            If top <= 0 Or right <= 0 Then GoTo veryend2
        End If
        
        For j = 1 To myseries.Dim1_Data.Count
            If xlog And a(j) <= 0 Then GoTo veryend3
            If ylog And b(j) <= 0 Then GoTo veryend3
            If xlog Then 'x is logarithmic axis
                x(j) = s_x + (Log(a(j)) - Log(left)) / (Log(right) - Log(left)) * width
            Else 'x is normal axis
                x(j) = s_x + (a(j) - left) / (right - left) * width
            End If
            
            If ylog Then 'y is logarithmic axis
                y(j) = s_y + (Log(b(j)) - Log(bottom)) / (Log(top) - Log(bottom)) * height
            Else 'y is normal axis
                y(j) = s_y + (b(j) - bottom) / (top - bottom) * height
            End If
        Next j
        
        Common.AutoRefresh False
        
        'draw shapes
        Dim s As Shape, d As Double
        If width > height Then d = height Else d = width
        Set s = ActiveLayer.CreateRectangle2(s_x + width + 5 * d * 0.03, s_y + height - 3 * i * d * 0.03, d * 0.03, d * 0.03)
        s.Name = "Legend"
        s.ObjectData("Type").value = "Symbol"
        s.Fill.ApplyUniformFill CreateRGBColor(cr, cg, cb)
        seriesmarks.Add s
        
        Set textshape = ActiveLayer.CreateArtisticText(0, 0, myseries.Name)
        Dim textscale As Double
        textscale = textshape.SizeHeight / textshape.SizeWidth
        textshape.SizeHeight = d * 0.03 * 2
        textshape.SizeWidth = textshape.SizeHeight / textscale
        textshape.PositionX = s_x + width + 10 * d * 0.03
        textshape.PositionY = s_y + height - 3 * i * d * 0.03 + textshape.SizeHeight / 2 + d * 0.03
        textshape.Name = "SeriesName"
        seriesmarks.Add textshape
        
        Set symbols = Common.DrawShapes(s, x, y)
        For j = 1 To symbols.Count
            If Not myseries.ElementNames Is Nothing Then symbols(j).Name = ename(j) Else symbols(j).Name = j
            symbols(j).ObjectData("Type").value = "Symbol"
            symbols(j).ObjectData("SymbolInfo").value = symbols(j).Name & "," & a(j) & "," & b(j)
        Next j
        seriesmarks.AddRange symbols
        
        'draw curve
        If PlotGeochemistryDiagram.XYCheck1 = True Then
            Dim legendconnecting As Shape
            Set connecting = Common.DrawCurve(x, y)
            If Not connecting Is Nothing Then
                Set legendconnecting = ActiveLayer.CreateCurveSegment(s_x + width + d * 0.03, s_y + height - 3 * i * d * 0.03 + d * 0.03 / 2, s_x + width + 10 * d * 0.03, s_y + height - 3 * i * d * 0.03 + d * 0.03 / 2)
                connecting.ObjectData("Type").value = "ConnectingLine"
                connecting.Name = "ConnectingLine"
                legendconnecting.ObjectData("Type").value = "ConnectingLine"
                legendconnecting.Name = "LegendConnectingLine"
                connecting.Outline.Color.RGBAssign cr, cg, cb
                legendconnecting.Outline.Color.RGBAssign cr, cg, cb
                connecting.OrderToBack
                legendconnecting.OrderToBack
                seriesmarks.Add connecting
                seriesmarks.Add legendconnecting
            End If
        End If
        
        Common.AutoRefresh True
        
        If seriesmarks.Count > 1 Then seriesmarks.group.Name = myseries.Name
        seriesmarks.RemoveAll
        ActiveSelectionRange.RemoveFromSelection
    Next i
    ActiveDocument.EndCommandGroup
    
    Exit Sub
veryend1:
    MsgBox "The template may be removed from the current document"
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
    Exit Sub
veryend2:
    MsgBox "The coordinate system is illegal."
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
    Exit Sub
veryend3:
    MsgBox "Some data are illegal."
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
    Exit Sub
finally:
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
    MsgBox "The Excel file may be closed."
End Sub

Sub PlotSpiderDiagram()
    If CurrentDiagramType <> "SpiderCoordinate" Then Exit Sub
    Dim a() As Double, k As Long, n As Long, normalized() As Double, ename() As String, _
        x() As Double, y() As Double, i As Long, j As Long, myseries As Series, s_y As Double, s_x As Double, height As Double, width As Double, _
        symbols As ShapeRange, connecting As Shape, seriesmarks As New ShapeRange, textshape As Shape
    If Not ActiveDocument.DataFields.IsPresent("Type") Then ActiveDocument.DataFields.AddEx "Type", cdrDataTypeString, "", , , "0", False, False, False
    If Not ActiveDocument.DataFields.IsPresent("SymbolInfo") Then ActiveDocument.DataFields.AddEx "SymbolInfo", cdrDataTypeString, "", , , "0", False, False, False
    
    On Error GoTo veryend1
    DiagramCurve.GetBoundingBox s_x, s_y, width, height
    
    If Geochemistry.HowManyColumns = -1 Then GoTo veryend1
    If Geochemistry.HowManyColumns = 0 Then GoTo veryend2
    
    ReDim normalized(1 To Geochemistry.HowManyColumns)
    i = 1
    Dim v As Variant, top As Double, bottom As Double, ylog As Boolean
    For Each v In VBA.Split(DiagramCurve.ObjectData("CoordinateInfo"), ",", -1, 1)
        If i = 1 Then bottom = Val(v)
        If i = 2 Then top = Val(v)
        If i = 3 Then
            If VBA.LCase(v) = "true" Then ylog = True Else ylog = False
        End If
        If i > 3 Then normalized(i - 3) = Val(v)
        i = i + 1
    Next v
    
    If VBA.Abs(bottom - top) = 0 Then GoTo veryend2
    If ylog Then
        If bottom <= 0 Or top <= 0 Then GoTo veryend2
    End If
    
    On Error GoTo finally
    
    ActiveDocument.BeginCommandGroup
    
    For i = 1 To Geochemistry.SpiderSeries.Count
        Dim cr As Integer, cg As Integer, cb As Integer
        Randomize
        cr = Int(255 * Rnd) + 1
        Randomize
        cg = Int(255 * Rnd) + 1
        Randomize
        cb = Int(255 * Rnd) + 1
        Set myseries = Geochemistry.SpiderSeries.Item(i)
        
        If PlotGeochemistryDiagram.SOption1.value Then
            ReDim x(1 To myseries.Dim1_Data.Columns.Count)
            ReDim y(1 To myseries.Dim1_Data.Columns.Count)
            n = myseries.Dim1_Data.Rows.Count
            ReDim ename(1 To n)
        Else
            ReDim x(1 To myseries.Dim1_Data.Rows.Count)
            ReDim y(1 To myseries.Dim1_Data.Rows.Count)
            n = myseries.Dim1_Data.Columns.Count
            ReDim ename(1 To n)
        End If
        
        Common.AutoRefresh False
        
        Dim s As Shape, d As Double
        If width > height Then d = height Else d = width
        Set s = ActiveLayer.CreateRectangle2(s_x + width + 5 * d * 0.03, s_y + height - 3 * i * d * 0.03, d * 0.03, d * 0.03)
        s.Name = "Legend"
        s.ObjectData("Type").value = "Symbol"
        s.Fill.ApplyUniformFill CreateRGBColor(cr, cg, cb)
        seriesmarks.Add s
        
        Set textshape = ActiveLayer.CreateArtisticText(0, 0, myseries.Name)
        Dim textscale As Double
        textscale = textshape.SizeHeight / textshape.SizeWidth
        textshape.SizeHeight = d * 0.03 * 2
        textshape.SizeWidth = textshape.SizeHeight / textscale
        textshape.PositionX = s_x + width + 10 * d * 0.03
        textshape.PositionY = s_y + height - 3 * i * d * 0.03 + textshape.SizeHeight / 2 + d * 0.03
        textshape.Name = "SeriesName"
        seriesmarks.Add textshape
        
        Dim l As Double
        
        If Not myseries.ElementNames Is Nothing Then
            Dim Index As Long
            Index = 1
            For Each v In myseries.ElementNames
                ename(Index) = v
                Index = Index + 1
            Next v
        End If
        
        For k = 1 To n
            If PlotGeochemistryDiagram.SOption1.value Then
                For j = 1 To myseries.Dim1_Data.Columns.Count
                    l = Val(myseries.Dim1_Data.Item(k, j))
                    If ylog Then
                        If l <= 0 Then GoTo veryend3
                        y(j) = s_y + (Log(l / normalized(j)) - Log(bottom)) / (Log(top) - Log(bottom)) * height
                    Else
                        y(j) = s_y + (l / normalized(j) - bottom) / (top - bottom) * height
                    End If
                    If myseries.Dim1_Data.Columns.Count = 1 Then x(j) = s_x Else _
                    x(j) = s_x + width / (myseries.Dim1_Data.Columns.Count - 1) * (j - 1)
                Next j
            Else
                For j = 1 To myseries.Dim1_Data.Rows.Count
                    l = Val(myseries.Dim1_Data.Item(j, k))
                    If ylog Then
                        y(j) = s_y + (Log(l / normalized(j)) - Log(bottom)) / (Log(top) - Log(bottom)) * height
                    Else
                        y(j) = s_y + (l / normalized(j) - bottom) / (top - bottom) * height
                    End If
                    If myseries.Dim1_Data.Rows.Count = 1 Then x(j) = s_x Else _
                    x(j) = s_x + width / (myseries.Dim1_Data.Rows.Count - 1) * (j - 1)
                Next j
            End If
                
                'draw shapes
                
                Set symbols = Common.DrawShapes(s, x, y)
                For j = 1 To symbols.Count
                    symbols(j).Name = "DataPoint"
                    symbols(j).ObjectData("Type").value = "Symbol"
                    If PlotGeochemistryDiagram.SOption1.value Then
                        symbols(j).ObjectData("SymbolInfo").value = symbols(j).Name & "," & Val(myseries.Dim1_Data.Item(k, j))
                    Else
                        symbols(j).ObjectData("SymbolInfo").value = symbols(j).Name & "," & Val(myseries.Dim1_Data.Item(j, k))
                    End If
                Next j
                seriesmarks.AddRange symbols
                
                'draw curve
                If PlotGeochemistryDiagram.SCheck1 = True Then
                    Dim legendconnecting As Shape
                    Set connecting = Common.DrawCurve(x, y)
                    If Not connecting Is Nothing Then
                        If k = 1 Then
                            Set legendconnecting = ActiveLayer.CreateCurveSegment(s_x + width + d * 0.03, s_y + height - 3 * i * d * 0.03 + d * 0.03 / 2, s_x + width + 10 * d * 0.03, s_y + height - 3 * i * d * 0.03 + d * 0.03 / 2)
                            legendconnecting.ObjectData("Type").value = "ConnectingLine"
                            legendconnecting.Name = "LegendConnectingLine"
                            legendconnecting.Outline.Color.RGBAssign cr, cg, cb
                            legendconnecting.OrderToBack
                            seriesmarks.Add legendconnecting
                        End If
                        connecting.ObjectData("Type").value = "ConnectingLine"
                        If Not myseries.ElementNames Is Nothing Then connecting.Name = ename(k) Else connecting.Name = k
                        connecting.Outline.Color.RGBAssign cr, cg, cb
                        connecting.OrderToBack
                        seriesmarks.Add connecting

                    End If
                End If
        Next k
        
        Common.AutoRefresh True
        
        If seriesmarks.Count > 1 Then seriesmarks.group.Name = myseries.Name
        seriesmarks.RemoveAll
        ActiveSelectionRange.RemoveFromSelection
        
    Next i
    ActiveDocument.EndCommandGroup
    
    Exit Sub
veryend1:
    MsgBox "The template may be removed from the current document"
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
    Exit Sub
veryend2:
    MsgBox "The coordinate system is illegal."
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
    Exit Sub
veryend3:
    MsgBox "Some data are illegal."
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
    Exit Sub
finally:
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
    MsgBox "The Excel file may be closed."
End Sub
