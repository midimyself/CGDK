VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SeriesEditor 
   Caption         =   "Series Editor"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3720
   OleObjectBlob   =   "SeriesEditor.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "SeriesEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim markshape As Shape

Private Sub CommandButton1_Click()
    If ActiveSelectionRange.Count < 1 Then MsgBox "Please select a symbol shape!": Exit Sub
    If ActiveSelectionRange.Count = 1 Then
        Set markshape = ActiveSelectionRange(1)
    Else
        Set markshape = ActiveSelectionRange.group
    End If
End Sub

Private Sub CommandButton2_Click()
    If markshape Is Nothing Then MsgBox "Please set a symbol shape first!": Exit Sub
    On Error GoTo veryend1
    Dim temp1 As Double
    temp1 = markshape.PositionX
    On Error GoTo end4
    Dim delsr As New ShapeRange
    If Not ActiveDocument.DataFields.IsPresent("Type") Then ActiveDocument.DataFields.AddEx "Type", cdrDataTypeString, "", , , "0", False, False, False
    If Not ActiveDocument.DataFields.IsPresent("SymbolInfo") Then ActiveDocument.DataFields.AddEx "SymbolInfo", cdrDataTypeString, "", , , "0", False, False, False
    ActiveDocument.BeginCommandGroup
    Replaceshapes ActiveSelectionRange, delsr
end4:
    Common.AutoRefresh True
    delsr.Delete
    ActiveDocument.EndCommandGroup
    Exit Sub
veryend1:
    MsgBox "The symbol shape may be removed from the document! Please set a new one."
End Sub

Private Sub Replaceshapes(sr As ShapeRange, delsr As ShapeRange)
    Dim i As Long, sr1 As ShapeRange, s As Shape, shaperangename As String
    For i = 1 To sr.Count
        If sr.Shapes(i).Type = cdrGroupShape Then
            shaperangename = sr.Shapes(i).Name
            Set sr1 = sr.Shapes(i).UngroupEx
            Replaceshapes sr1, delsr
            Common.AutoRefresh True
            sr1.group.Name = shaperangename
            ActiveSelectionRange.RemoveFromSelection
        Else
            Common.AutoRefresh False
            If VBA.LCase(sr.Shapes(i).ObjectData("Type").value) = "symbol" Then
            Set s = markshape.Duplicate
            If Me.CheckBox1.value Then
                s.SizeHeight = sr.Shapes(i).SizeHeight
                s.SizeWidth = sr.Shapes(i).SizeWidth
            End If
            s.Name = sr.Shapes(i).Name
            s.ObjectData("Type").value = sr.Shapes(i).ObjectData("Type").value
            s.ObjectData("SymbolInfo").value = sr.Shapes(i).ObjectData("SymbolInfo").value
            Common.ShapeAlign s, sr.Shapes(i), 0, 0
            s.OrderToFront
            sr.Add s
            delsr.Add sr.Shapes(i)
            Common.AutoRefresh True
            End If
        End If
    Next i
End Sub

Private Sub CommandButton3_Click()
    If Not ActiveDocument.DataFields.IsPresent("Type") Then ActiveDocument.DataFields.AddEx "Type", cdrDataTypeString, "", , , "0", False, False, False
    ActiveDocument.BeginCommandGroup
    SmoothCurve ActiveSelectionRange, True
end4:
    Common.AutoRefresh True
    ActiveDocument.EndCommandGroup
End Sub
Private Sub SmoothCurve(sr As ShapeRange, ToSmooth As Boolean)
    Dim i As Long, sr1 As ShapeRange, s As Shape, shaperangename As String, cur As Curve, crv As Curve, j As Long, sp As SubPath
    For i = 1 To sr.Count
        If sr.Shapes(i).Type = cdrGroupShape Then
            shaperangename = sr.Shapes(i).Name
            Set sr1 = sr.Shapes(i).UngroupEx
            SmoothCurve sr1, ToSmooth
            sr1.group.Name = shaperangename
        Else
            Common.AutoRefresh False
            If VBA.LCase(sr.Shapes(i).ObjectData("Type").value) = "connectingline" Then
                Set cur = sr.Shapes(i).Curve
                If ToSmooth Then
                    cur.Nodes.All.SetType cdrSmoothNode
                    Set s = ActiveLayer.CreateCurve(cur)
                Else
                    Set crv = CreateCurve(ActiveDocument)
                    For j = 1 To sr.Shapes(i).Curve.Nodes.Count
                        If j = 1 Then
                            Set sp = crv.CreateSubPath(sr.Shapes(i).Curve.Nodes(j).PositionX, sr.Shapes(i).Curve.Nodes(j).PositionY)
                        Else
                            sp.AppendCurveSegment sr.Shapes(i).Curve.Nodes(j).PositionX, sr.Shapes(i).Curve.Nodes(j).PositionY
                        End If
                    Next j
                    Set s = ActiveLayer.CreateCurve(crv)
                End If
                s.Name = sr.Shapes(i).Name
                s.ObjectData("Type").value = sr.Shapes(i).ObjectData("Type").value
                s.Outline.CopyAssign sr.Shapes(i).Outline
                s.OrderToBack
                sr.Add s
                s.Selected = False
                Common.AutoRefresh True
                sr.Shapes(i).Delete
            End If
        End If
    Next i
End Sub

Private Sub CommandButton4_Click()
    If Not ActiveDocument.DataFields.IsPresent("Type") Then ActiveDocument.DataFields.AddEx "Type", cdrDataTypeString, "", , , "0", False, False, False
    ActiveDocument.BeginCommandGroup
    SmoothCurve ActiveSelectionRange, False
end4:
    Common.AutoRefresh True
    ActiveDocument.EndCommandGroup
End Sub
