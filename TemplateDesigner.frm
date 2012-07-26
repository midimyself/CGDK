VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateDesigner 
   Caption         =   "Template Designer"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   OleObjectBlob   =   "TemplateDesigner.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "TemplateDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private normalizedvalues As Object

Private markshape As Shape

Private Sub Axises_Click()
    Me.CoordinateFrame.Visible = False
    Me.Axises2.Visible = True
    Me.Marks2.Visible = False
End Sub

Private Sub ComboBox1_Change()
    Select Case Me.ComboBox1.ListIndex
        Case 0 'Triangular diagram
            Me.Label2.Enabled = False
            Me.Label3.Enabled = False
            Me.Label4.Enabled = False
            Me.Label5.Enabled = False
            Me.Label6.Enabled = False
            Me.CheckBox1.Enabled = False
            Me.CheckBox2.Enabled = False
            Me.TextBox1.Enabled = False
            Me.TextBox2.Enabled = False
            Me.TextBox3.Enabled = False
            Me.TextBox4.Enabled = False
            Me.TextBox5.Enabled = False
            Me.CommandButton2.Enabled = False
            Me.CommandButton3.Enabled = False
        Case 1 'Scatter diagram
            Me.Label2.Enabled = True
            Me.Label3.Enabled = True
            Me.Label4.Enabled = True
            Me.Label5.Enabled = True
            Me.Label6.Enabled = False
            Me.CheckBox1.Enabled = True
            Me.CheckBox2.Enabled = True
            Me.TextBox1.Enabled = True
            Me.TextBox2.Enabled = True
            Me.TextBox3.Enabled = True
            Me.TextBox4.Enabled = True
            Me.TextBox5.Enabled = False
            Me.CommandButton2.Enabled = False
            Me.CommandButton3.Enabled = False
        Case 2 'Spider diagram
            Me.Label2.Enabled = False
            Me.Label3.Enabled = False
            Me.Label4.Enabled = True
            Me.Label5.Enabled = True
            Me.Label6.Enabled = True
            Me.CheckBox1.Enabled = False
            Me.CheckBox2.Enabled = True
            Me.TextBox1.Enabled = False
            Me.TextBox2.Enabled = False
            Me.TextBox3.Enabled = True
            Me.TextBox4.Enabled = True
            Me.TextBox5.Enabled = True
            Me.CommandButton2.Enabled = True
            Me.CommandButton3.Enabled = True
    End Select
End Sub

Private Sub CommandButton1_Click()
    If ActiveSelectionRange.Shapes.Count <> 1 Then MsgBox "Please select a shape!": Exit Sub
    If ActiveSelectionRange.Shapes(1).SizeHeight < 0 Then MsgBox "The height of the shape is too small.": Exit Sub
    If ActiveSelectionRange.Shapes(1).SizeWidth < 0 Then MsgBox "The width of the shape is too small.": Exit Sub
    Select Case Me.ComboBox1.ListIndex
        Case 0 'triangular diagram
            Geochemistry.EstablishTriangularCoordinate ActiveSelectionRange.Shapes(1)
        Case 1 'scatter diagram
            Geochemistry.Establish_XY_ScatterCoordinate ActiveSelectionRange.Shapes(1), Val(Me.TextBox1.value) _
            , Val(Me.TextBox2.value), Val(Me.TextBox4.value), Val(Me.TextBox3), Me.CheckBox1.value, Me.CheckBox2.value
        Case 2 'spider diagram
            If Me.TextBox5 = "" Then MsgBox "You should assign normalized values to the spider diagram!": Exit Sub
            On Error GoTo finally
            Geochemistry.EstablishSpiderCoordinate ActiveSelectionRange.Shapes(1), Val(Me.TextBox4.value) _
            , Val(Me.TextBox3.value), Me.CheckBox2.value, normalizedvalues
    End Select
    Exit Sub
finally:
    Me.TextBox5.value = ""
    Set normalizedvalues = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton2_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not normalizedvalues Is Nothing Then
        Dim temp As Double
        temp = Val(normalizedvalues.Count)
    End If
    If Me.TextBox5.value <> "" Then
        Set obj = GetDataRange.GetDataRange(normalizedvalues)
        If Not obj Is Nothing Then Set normalizedvalues = obj
    Else
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set normalizedvalues = obj
        On Error Resume Next
        Me.TextBox5 = normalizedvalues.Address
        On Error GoTo 0
    End If
    Exit Sub
finally:
    TextBox5.value = ""
    Set normalizedvalues = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton3_Click()
    Me.TextBox5.value = ""
End Sub

Private Sub CommandButton4_Click()
    Dim s As Shape, a As String, v As Variant, s_x As Double, s_y As Double, height As Double, width As Double, textshape As Shape, textscale As Double
    Dim d As Double, x As Double, y As Double
    Dim left As Double, right As Double, top As Double, bottom As Double, xlog As Boolean, ylog As Boolean
    On Error GoTo veryend1
    Set s = Geochemistry.DiagramCurve
    s.GetBoundingBox s_x, s_y, width, height
    On Error GoTo finally
    
    If width > height Then d = height Else d = width
    
    If Me.Label11.Caption = "Triangular diagram" Then 'draw triangular diagram axies
        Common.AutoRefresh False
        ActiveDocument.BeginCommandGroup
        For Each v In VBA.Split(Me.TextBox6, ",", -1, vbTextCompare)
            If Val(v) >= 100 Or Val(v) <= 0 Then
            Else
                Call ActiveLayer.CreateLineSegment(width / 200 * Val(v) + s_x, s_y + Val(v) / 100 * height, s_x + width - width / 200 * Val(v), s_y + Val(v) / 100 * height)
                Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
                textscale = textshape.SizeHeight / textshape.SizeWidth
                textshape.SizeHeight = d * 0.02 * 2
                textshape.SizeWidth = textshape.SizeHeight / textscale
                textshape.PositionX = width / 200 * Val(v) + s_x - textshape.SizeWidth - textshape.SizeHeight / 2
                textshape.PositionY = s_y + Val(v) / 100 * height + textshape.SizeHeight / 2
            End If
        Next v
        For Each v In VBA.Split(Me.TextBox7, ",", -1, vbTextCompare)
            If Val(v) >= 100 Or Val(v) <= 0 Then
            Else
                Call ActiveLayer.CreateLineSegment(width / 100 * (100 - Val(v)) + s_x, s_y, s_x + width / 200 * (100 - Val(v)), s_y + (100 - Val(v)) / 100 * height)
                Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
                textscale = textshape.SizeHeight / textshape.SizeWidth
                textshape.SizeHeight = d * 0.02 * 2
                textshape.SizeWidth = textshape.SizeHeight / textscale
                textshape.PositionX = (100 - Val(v)) / 100 * width + s_x - textshape.SizeWidth / 2
                textshape.PositionY = s_y - textshape.SizeHeight
            End If
        Next v
            For Each v In VBA.Split(Me.TextBox8, ",", -1, vbTextCompare)
            If Val(v) >= 100 Or Val(v) <= 0 Then
            Else
                Call ActiveLayer.CreateLineSegment(width / 100 * Val(v) + s_x, s_y, s_x + width / 2 + width / 200 * Val(v), s_y + (100 - Val(v)) / 100 * height)
                Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
                textscale = textshape.SizeHeight / textshape.SizeWidth
                textshape.SizeHeight = d * 0.02 * 2
                textshape.SizeWidth = textshape.SizeHeight / textscale
                textshape.PositionX = s_x + width / 2 + width / 200 * Val(v) + textshape.SizeHeight / 2
                textshape.PositionY = s_y + (100 - Val(v)) / 100 * height + textshape.SizeHeight / 2
            End If
        Next v
        ActiveSelectionRange.RemoveFromSelection
        ActiveDocument.EndCommandGroup
        Common.AutoRefresh True
    ElseIf Me.Label11.Caption = "XYScatter diagram" Then 'draw XYScatter diagram axies
        Geochemistry.GetXYCoordinateInfo left, right, top, bottom, xlog, ylog
        If VBA.Abs(left - right) = 0 Or VBA.Abs(top - bottom) = 0 Then GoTo veryend2
        If xlog Then
            If right <= 0 Or left <= 0 Then GoTo veryend2
        End If
        If ylog Then
            If top <= 0 Or bottom <= 0 Then GoTo veryend2
        End If
        Common.AutoRefresh False
        ActiveDocument.BeginCommandGroup
        
        For Each v In VBA.Split(Me.TextBox6, ",", -1, vbTextCompare)

            If xlog Then
                If Val(v) > 0 Then
                x = s_x + (Log(Val(v)) - Log(left)) / (Log(right) - Log(left)) * width
                
                Call ActiveLayer.CreateLineSegment(x, s_y, x, s_y - 0.02 * d)
                Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
                textscale = textshape.SizeHeight / textshape.SizeWidth
                textshape.SizeHeight = d * 0.02 * 2
                textshape.SizeWidth = textshape.SizeHeight / textscale
                textshape.PositionX = x - textshape.SizeWidth / 2
                textshape.PositionY = s_y - textshape.SizeHeight
                End If
            Else 'x is normal axis
                x = s_x + (Val(v) - left) / (right - left) * width
                Call ActiveLayer.CreateLineSegment(x, s_y, x, s_y - 0.02 * d)
                Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
                textscale = textshape.SizeHeight / textshape.SizeWidth
                textshape.SizeHeight = d * 0.02 * 2
                textshape.SizeWidth = textshape.SizeHeight / textscale
                textshape.PositionX = x - textshape.SizeWidth / 2
                textshape.PositionY = s_y - textshape.SizeHeight
            End If
        Next v
        
        For Each v In VBA.Split(Me.TextBox7, ",", -1, vbTextCompare)

            If ylog Then
                If Val(v) > 0 Then
                y = s_y + (Log(Val(v)) - Log(bottom)) / (Log(top) - Log(bottom)) * height
                Call ActiveLayer.CreateLineSegment(s_x, y, s_x - 0.02 * d, y)
                Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
                textscale = textshape.SizeHeight / textshape.SizeWidth
                textshape.SizeHeight = d * 0.02 * 2
                textshape.SizeWidth = textshape.SizeHeight / textscale
                textshape.PositionX = s_x - textshape.SizeWidth - textshape.SizeHeight / 2
                textshape.PositionY = y + textshape.SizeHeight / 2
                End If
            Else 'y is normal axis
                y = s_y + (Val(v) - bottom) / (top - bottom) * height
                Call ActiveLayer.CreateLineSegment(s_x, y, s_x - 0.02 * d, y)
                Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
                textscale = textshape.SizeHeight / textshape.SizeWidth
                textshape.SizeHeight = d * 0.02 * 2
                textshape.SizeWidth = textshape.SizeHeight / textscale
                textshape.PositionX = s_x - textshape.SizeWidth - textshape.SizeHeight / 2
                textshape.PositionY = y + textshape.SizeHeight / 2
            End If
        Next v
        ActiveSelectionRange.RemoveFromSelection
        ActiveDocument.EndCommandGroup
        Common.AutoRefresh True
        
        
    ElseIf Me.Label11.Caption = "Spider diagram" Then 'draw spider diagram axis
        ActiveDocument.BeginCommandGroup
        Common.AutoRefresh False
        
        Dim n As Long
        n = Geochemistry.HowManyColumns
        If n = -1 Then GoTo veryend2
        
        bottom = Val(VBA.Split(DiagramCurve.ObjectData("CoordinateInfo"), ",", -1, 1)(0))
        top = Val(VBA.Split(DiagramCurve.ObjectData("CoordinateInfo"), ",", -1, 1)(1))
        If VBA.LCase(VBA.Split(DiagramCurve.ObjectData("CoordinateInfo"), ",", -1, 1)(2)) = "true" Then ylog = True Else ylog = False
        
        If ylog Then
            If bottom <= 0 Or top <= 0 Then GoTo veryend2
        End If
         
         Dim i As Long, cell As Double
         i = 0
         If Geochemistry.HowManyColumns = 1 Then cell = width Else cell = width / (Geochemistry.HowManyColumns - 1)
         
         For Each v In VBA.Split(Me.TextBox6, ",", -1, vbTextCompare)
             x = s_x + cell * i
             Call ActiveLayer.CreateLineSegment(x, s_y, x, s_y - d * 0.02)
                Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
                textscale = textshape.SizeHeight / textshape.SizeWidth
                textshape.SizeHeight = d * 0.02 * 2
                textshape.SizeWidth = textshape.SizeHeight / textscale
                textshape.PositionX = x - textshape.SizeWidth / 2
                textshape.PositionY = s_y - textshape.SizeHeight
             i = i + 1
         Next v
         For Each v In VBA.Split(Me.TextBox7, ",", -1, vbTextCompare)
            If ylog Then
                If Val(v) > 0 Then
                y = s_y + (Log(Val(v)) - Log(bottom)) / (Log(top) - Log(bottom)) * height
                Call ActiveLayer.CreateLineSegment(s_x, y, s_x - 0.02 * d, y)
                Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
                textscale = textshape.SizeHeight / textshape.SizeWidth
                textshape.SizeHeight = d * 0.02 * 2
                textshape.SizeWidth = textshape.SizeHeight / textscale
                textshape.PositionX = s_x - textshape.SizeWidth - textshape.SizeHeight / 2
                textshape.PositionY = y + textshape.SizeHeight / 2
                End If
            Else 'y is normal axis
                y = s_y + (Val(v) - bottom) / (top - bottom) * height
                Call ActiveLayer.CreateLineSegment(s_x, y, s_x - 0.02 * d, y)
                Set textshape = ActiveLayer.CreateArtisticText(0, 0, v)
                textscale = textshape.SizeHeight / textshape.SizeWidth
                textshape.SizeHeight = d * 0.02 * 2
                textshape.SizeWidth = textshape.SizeHeight / textscale
                textshape.PositionX = s_x - textshape.SizeWidth - textshape.SizeHeight / 2
                textshape.PositionY = y + textshape.SizeHeight / 2
            End If
        Next v
        ActiveSelectionRange.RemoveFromSelection
        ActiveDocument.EndCommandGroup
        Common.AutoRefresh True
    End If
    Exit Sub
veryend1:
    MsgBox "The template may be removed from the document."
    Exit Sub
veryend2:
    MsgBox "The coordinate system is illegal."
    Exit Sub
finally:
    MsgBox Err.Description
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
End Sub

Private Sub CommandButton6_Click()
    If ActiveSelectionRange.Count = 0 Then MsgBox "Please select a shape!": Exit Sub
    If ActiveSelectionRange.Count > 1 Then
        Set markshape = ActiveSelectionRange.group
    Else
        Set markshape = ActiveSelectionRange.Shapes(1)
    End If
End Sub

Private Sub CommandButton7_Click()
    If markshape Is Nothing Then MsgBox "Please set a symbol shape!": Exit Sub
    On Error GoTo veryend1
        Dim temp1 As Double
        temp1 = markshape.PositionX
    Dim i As Long, s_x As Double, s_y As Double, height As Double, width As Double, sr As New ShapeRange, s As Shape
    Dim a As Double, b As Double, c As Double, a1 As Double, b1 As Double, c1 As Double, x As Double, y As Double
    Dim left As Double, right As Double, top As Double, bottom As Double, xlog As Boolean, ylog As Boolean
    On Error GoTo veryend2
    DiagramCurve.GetBoundingBox s_x, s_y, width, height
    On Error GoTo end4
    If Not ActiveDocument.DataFields.IsPresent("Type") Then ActiveDocument.DataFields.AddEx "Type", cdrDataTypeString, "", , , "0", False, False, False
    If Not ActiveDocument.DataFields.IsPresent("SymbolInfo") Then ActiveDocument.DataFields.AddEx "SymbolInfo", cdrDataTypeString, "", , , "0", False, False, False
    Common.AutoRefresh False
    ActiveDocument.BeginCommandGroup
    
    If Me.Label15.Caption = "Triangular diagram" Then 'draw triangular diagram marks
        For i = 1 To Common.HowManyElementsInArray(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare))
            On Error GoTo end1
                a = VBA.Abs(Val(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare)(i - 1)))
                b = VBA.Abs(Val(VBA.Split(Me.TextBox10, ",", -1, vbTextCompare)(i - 1)))
                c = VBA.Abs(Val(VBA.Split(Me.TextBox11, ",", -1, vbTextCompare)(i - 1)))
                a1 = a / (a + b + c)
                b1 = b / (a + b + c)
                c1 = c / (a + b + c)
                x = s_x + c1 * width + a1 * width / 2
                y = s_y + a1 * height
                Set s = markshape.Duplicate
                sr.Add s
                s.ObjectData("Type").value = "Symbol"
                s.ObjectData("SymbolInfo").value = i & "," & a & "," & b & "," & c
                s.SetPosition x - markshape.SizeWidth / 2, y + markshape.SizeHeight / 2
end1:
        Next i
    ElseIf Me.Label15.Caption = "XYScatter diagram" Then 'draw scatter diagram marks
        If Not Geochemistry.GetXYCoordinateInfo(left, right, top, bottom, xlog, ylog) Then GoTo end4
        
        For i = 1 To Common.HowManyElementsInArray(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare))
            On Error GoTo end2
                a = Val(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare)(i - 1))
                b = Val(VBA.Split(Me.TextBox10, ",", -1, vbTextCompare)(i - 1))
                If xlog Then
                    x = s_x + (Log(a) - Log(left)) / (Log(right) - Log(left)) * width
                Else
                    x = s_x + (a - left) / (right - left) * width
                End If
                If ylog Then
                    y = s_y + (Log(b) - Log(bottom)) / (Log(top) - Log(bottom)) * height
                Else
                    y = s_y + (b - bottom) / (top - bottom) * height
                End If
                Set s = markshape.Duplicate
                s.ObjectData("Type").value = "Symbol"
                s.ObjectData("SymbolInfo").value = i & "," & a & "," & b
                s.SetPosition x - markshape.SizeWidth / 2, y + markshape.SizeHeight / 2
                sr.Add s
end2:
        Next i
    ElseIf Me.Label15.Caption = "Spider diagram" Then 'draw spider diagram marks
        If Geochemistry.HowManyColumns = -1 Then GoTo end4
        If Geochemistry.HowManyColumns = 0 Then GoTo end4
        
        Dim normalized() As Double
        ReDim normalized(1 To Geochemistry.HowManyColumns)
        i = 1
        Dim v As Variant
        For Each v In VBA.Split(DiagramCurve.ObjectData("CoordinateInfo"), ",", -1, 1)
            If i = 1 Then bottom = Val(v)
            If i = 2 Then top = Val(v)
            If i = 3 Then
                If VBA.LCase(v) = "true" Then ylog = True Else ylog = False
            End If
            If i > 3 Then normalized(i - 3) = Val(v)
            i = i + 1
        Next v
        
        If VBA.Abs(bottom - top) = 0 Then GoTo end4
        If ylog Then
            If bottom <= 0 Or top <= 0 Then GoTo end4
        End If
        Dim n As Long
        If Common.HowManyElementsInArray(normalized) > Common.HowManyElementsInArray(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare)) Then
            n = Common.HowManyElementsInArray(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare))
        Else
            n = Common.HowManyElementsInArray(normalized)
        End If
        For i = 1 To n
            On Error GoTo end3
                a = Val(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare)(i - 1))
                If ylog Then
                    y = s_y + (Log(a / normalized(i)) - Log(bottom)) / (Log(top) - Log(bottom)) * height
                Else
                    y = s_y + (a / normalized(i) - bottom) / (top - bottom) * height
                End If
                If Common.HowManyElementsInArray(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare)) = 1 Then x = s_x Else _
                    x = s_x + width / (Common.HowManyElementsInArray(normalized) - 1) * (i - 1)
                Set s = markshape.Duplicate
                s.ObjectData("Type").value = "Symbol"
                s.ObjectData("SymbolInfo").value = i & "," & a
                s.SetPosition x - markshape.SizeWidth / 2, y + markshape.SizeHeight / 2
                sr.Add s
end3:
        Next i
    End If
end4:
    Common.AutoRefresh True
    If sr.Count > 1 Then sr.group.Name = "Marks"
    sr.RemoveAll
    ActiveDocument.EndCommandGroup
    Exit Sub
veryend1:
    MsgBox "The shape may be removed from the document! Please reset the sample shape!"
    Exit Sub
veryend2:
    MsgBox "The template may be removed from the document! Please select a new template!"
End Sub

Private Sub CommandButton8_Click()
    Dim i As Long, s_x As Double, s_y As Double, height As Double, width As Double, s As Shape, cur As Curve, sp As SubPath
    Dim a As Double, b As Double, c As Double, a1 As Double, b1 As Double, c1 As Double, x As Double, y As Double
    Dim left As Double, right As Double, top As Double, bottom As Double, xlog As Boolean, ylog As Boolean
    On Error GoTo veryend2
    DiagramCurve.GetBoundingBox s_x, s_y, width, height
    On Error GoTo end4
    If Not ActiveDocument.DataFields.IsPresent("Type") Then ActiveDocument.DataFields.AddEx "Type", cdrDataTypeString, "", , , "0", False, False, False
    If Not ActiveDocument.DataFields.IsPresent("SymbolInfo") Then ActiveDocument.DataFields.AddEx "SymbolInfo", cdrDataTypeString, "", , , "0", False, False, False
    Common.AutoRefresh False
    ActiveDocument.BeginCommandGroup
    
    Dim countnodes As Long
    countnodes = 0
    Set cur = CreateCurve(ActiveDocument)
    
    If Me.Label15.Caption = "Triangular diagram" Then 'draw triangular diagram curve
        For i = 1 To Common.HowManyElementsInArray(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare))
            On Error GoTo end1
                a = VBA.Abs(Val(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare)(i - 1)))
                b = VBA.Abs(Val(VBA.Split(Me.TextBox10, ",", -1, vbTextCompare)(i - 1)))
                c = VBA.Abs(Val(VBA.Split(Me.TextBox11, ",", -1, vbTextCompare)(i - 1)))
                a1 = a / (a + b + c)
                b1 = b / (a + b + c)
                c1 = c / (a + b + c)
                x = s_x + c1 * width + a1 * width / 2
                y = s_y + a1 * height
                countnodes = countnodes + 1
                If countnodes = 1 Then
                    Set sp = cur.CreateSubPath(x, y)
                Else
                    sp.AppendCurveSegment x, y
                End If
end1:
        Next i
    ElseIf Me.Label15.Caption = "XYScatter diagram" Then 'draw scatter diagram marks
        If Not Geochemistry.GetXYCoordinateInfo(left, right, top, bottom, xlog, ylog) Then GoTo end4
        
        For i = 1 To Common.HowManyElementsInArray(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare))
            On Error GoTo end2
                a = Val(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare)(i - 1))
                b = Val(VBA.Split(Me.TextBox10, ",", -1, vbTextCompare)(i - 1))
                If xlog Then
                    x = s_x + (Log(a) - Log(left)) / (Log(right) - Log(left)) * width
                Else
                    x = s_x + (a - left) / (right - left) * width
                End If
                If ylog Then
                    y = s_y + (Log(b) - Log(bottom)) / (Log(top) - Log(bottom)) * height
                Else
                    y = s_y + (b - bottom) / (top - bottom) * height
                End If
                countnodes = countnodes + 1
                If countnodes = 1 Then
                    Set sp = cur.CreateSubPath(x, y)
                Else
                    sp.AppendCurveSegment x, y
                End If
end2:
        Next i
    ElseIf Me.Label15.Caption = "Spider diagram" Then 'draw spider diagram marks
        If Geochemistry.HowManyColumns = -1 Then GoTo end4
        If Geochemistry.HowManyColumns = 0 Then GoTo end4
        
        Dim normalized() As Double
        ReDim normalized(1 To Geochemistry.HowManyColumns)
        i = 1
        Dim v As Variant
        For Each v In VBA.Split(DiagramCurve.ObjectData("CoordinateInfo"), ",", -1, 1)
            If i = 1 Then bottom = Val(v)
            If i = 2 Then top = Val(v)
            If i = 3 Then
                If VBA.LCase(v) = "true" Then ylog = True Else ylog = False
            End If
        If i > 3 Then normalized(i - 3) = Val(v)
        i = i + 1
        Next v
        
        If VBA.Abs(bottom - top) = 0 Then GoTo end4
        If ylog Then
            If bottom <= 0 Or top <= 0 Then GoTo end4
        End If
        Dim n As Long
        If Common.HowManyElementsInArray(normalized) > Common.HowManyElementsInArray(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare)) Then
            n = Common.HowManyElementsInArray(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare))
        Else
            n = Common.HowManyElementsInArray(normalized)
        End If
        For i = 1 To n
            On Error GoTo end3
                a = Val(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare)(i - 1))
                If ylog Then
                    y = s_y + (Log(a / normalized(i)) - Log(bottom)) / (Log(top) - Log(bottom)) * height
                Else
                    y = s_y + (a / normalized(i) - bottom) / (top - bottom) * height
                End If
                If Common.HowManyElementsInArray(VBA.Split(Me.TextBox9, ",", -1, vbTextCompare)) = 1 Then x = s_x Else _
                    x = s_x + width / (Common.HowManyElementsInArray(normalized) - 1) * (i - 1)
                
                countnodes = countnodes + 1
                If countnodes = 1 Then
                    Set sp = cur.CreateSubPath(x, y)
                Else
                    sp.AppendCurveSegment x, y
                End If
end3:
        Next i
    End If
end4:
    If countnodes > 1 Then
        Set s = ActiveLayer.CreateCurve(cur)
        s.ObjectData("Type").value = "ConnectingLine"
    End If
    ActiveDocument.EndCommandGroup
    Common.AutoRefresh True
    Exit Sub
veryend2:
    MsgBox "The template may be removed from the document! Please select a new template!"
End Sub

Private Sub Coordinate_Click()
    Me.CoordinateFrame.Visible = True
    Me.Axises2.Visible = False
    Me.Marks2.Visible = False
End Sub

Private Sub Marks_Click()
    Me.CoordinateFrame.Visible = False
    Me.Axises2.Visible = False
    Me.Marks2.Visible = True
End Sub

Private Sub UserForm_Initialize()
    Me.ComboBox1.AddItem "Triangular diagram"
    Me.ComboBox1.AddItem "Scatter diagram"
    Me.ComboBox1.AddItem "Spider diagram"
    Me.Label2.Enabled = False
    Me.Label3.Enabled = False
    Me.Label4.Enabled = False
    Me.Label5.Enabled = False
    Me.Label6.Enabled = False
    Me.CheckBox1.Enabled = False
    Me.CheckBox2.Enabled = False
    Me.TextBox1.Enabled = False
    Me.TextBox2.Enabled = False
    Me.TextBox3.Enabled = False
    Me.TextBox4.Enabled = False
    Me.TextBox5.Enabled = False
    Me.CommandButton2.Enabled = False
    Me.CommandButton3.Enabled = False
    API.StartTimer1 (150)
End Sub

Private Sub UserForm_Terminate()
    Unload data
    API.StopTimer1
End Sub
