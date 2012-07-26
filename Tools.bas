Attribute VB_Name = "Tools"
Option Explicit

Sub SmartFill()
    If Common.IsWaitingForAction Then Exit Sub
    Common.IsWaitingForAction = True
    On Error GoTo veryend
    Dim x As Double, y As Double, Shift As Long, b As Boolean, sel As Shape, s As Shape, filestyle As Fill, i As Long, outlinestyle As Outline
    b = False
    Do While Not b
begin:
        b = ActiveDocument.GetUserClick(x, y, Shift, 10, False, cdrCursorWinCross)
        If b Then Exit Do
        If Shift = 1 Then
            Set sel = ActiveDocument.ActivePage.SelectShapesAtPoint(x, y, False)
            If sel Is Nothing Then GoTo begin
            If sel.Shapes.Count > 0 Then Set s = sel.Shapes(1) Else GoTo begin
            ActiveSelectionRange.RemoveFromSelection
            s.Selected = True
            Set filestyle = s.Fill.GetCopy
            Set outlinestyle = s.Outline.GetCopy
        Else
            If filestyle Is Nothing Then GoTo begin
            If outlinestyle Is Nothing Then GoTo begin
            ActiveDocument.BeginCommandGroup
            For i = 1 To ActivePage.Shapes.Count
                If ActivePage.Shapes(i).Type = cdrCurveShape Then
                    If Not ActivePage.Shapes(i).Curve.Closed Then ActivePage.Shapes(i).Fill.ApplyNoFill
                End If
            Next i
            ActiveDocument.EndCommandGroup
            If CommonTools.CheckBox2 And CommonTools.CheckBox1 Then
                ActiveDocument.BeginCommandGroup
                Set s = ActivePage.CustomCommand("Boundary", "SmartFill", x, y, Nothing, 0, Nothing)
                s.Fill.CopyAssign filestyle
                s.Outline.CopyAssign outlinestyle
                s.OrderToBack
                ActiveDocument.EndCommandGroup
            End If
            If CommonTools.CheckBox2 And Not CommonTools.CheckBox1 Then
                Set sel = ActiveDocument.ActivePage.SelectShapesAtPoint(x, y, False)
                If sel Is Nothing Then GoTo begin
                If sel.Shapes.Count > 0 Then Set s = sel.Shapes(1) Else GoTo begin
                ActiveSelectionRange.RemoveFromSelection
                s.Selected = True
                s.Outline.CopyAssign outlinestyle
                s.OrderToFront
            End If
            If Not CommonTools.CheckBox2 And CommonTools.CheckBox1 Then
                ActiveDocument.BeginCommandGroup
                Set s = ActivePage.CustomCommand("Boundary", "SmartFill", x, y, Nothing, 0, Nothing)
                s.Fill.CopyAssign filestyle
                s.OrderToBack
                ActiveDocument.EndCommandGroup
            End If
        End If
    Loop
    Common.IsWaitingForAction = False
    Exit Sub
veryend:
    Common.IsWaitingForAction = False
End Sub

Sub SmartSelect()
    If Common.IsWaitingForAction Then Exit Sub
    Common.IsWaitingForAction = True
    Dim i As Long, j As Long, k As Long, p As Page, l As Layer, s As Shape, sr As New ShapeRange, object As Shape
    Dim x As Double, y As Double, Shift As Long, b As Boolean
    
begin:
    b = False
    Do While Not b
        b = ActiveDocument.GetUserClick(x, y, Shift, 10, False, cdrCursorWinCross)
        If b Then Exit Do
        sr.AddRange ActiveSelectionRange
        sr.RemoveFromSelection
        sr.RemoveAll
        Set object = ActiveDocument.ActivePage.SelectShapesAtPoint(x, y, False)
        On Error GoTo veryend
        Common.AutoRefresh False
        For i = 1 To ActiveDocument.Pages.Count
            Set p = ActiveDocument.Pages(i)
            For j = 1 To p.Layers.Count
                Set l = p.Layers(j)
                If l.Visible And l.Editable Then
                    For k = 1 To l.Shapes.Count
                        Set s = l.Shapes(k)
                        If s.Type = cdrCurveShape Then
                            If Not s.Curve.Closed Then s.Fill.ApplyNoFill
                        End If
                        If CommonTools.CheckBox1 And Not CommonTools.CheckBox2 Then
                            If s.Fill.CompareWith(object.Fill) Then
                                If s.Type = cdrCurveShape Then
                                    If s.Curve.Closed = True Then sr.Add s
                                Else
                                    sr.Add s
                                End If
                            End If
                        ElseIf Not CommonTools.CheckBox1 And CommonTools.CheckBox2 Then
                            If s.Outline.CompareWith(object.Outline) Then
                                sr.Add s
                            End If
                        ElseIf CommonTools.CheckBox1 And CommonTools.CheckBox2 Then
                            If s.Fill.CompareWith(object.Fill) And s.Outline.CompareWith(object.Outline) Then
                                If s.Type = cdrCurveShape Then
                                    If s.Curve.Closed = True Then sr.Add s
                                Else
                                    sr.Add s
                                End If
                            End If
                        End If
                    Next k
                End If
            Next j
        Next i
        Common.AutoRefresh True
        sr.AddToSelection
        On Error GoTo 0
    Loop
    sr.AddToSelection
    Common.IsWaitingForAction = False
    Exit Sub
veryend:
    sr.AddRange ActiveSelectionRange
    sr.RemoveFromSelection
    sr.RemoveAll
    Resume begin
End Sub
