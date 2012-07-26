Attribute VB_Name = "Common"
Option Explicit

Public Const cdrCenterAlign = 0
Public Const cdrUpperLeftAlign = 1
Public Const cdrUpperRightAlign = 2
Public Const cdrLowerLeftAlign = 3
Public Const cdrLowerRightAlign = 4
Public Const cdrSourceShapeSize = 0
Public Const cdrTargetShapeSize = 1
Public Const cdrCustomSize = 2

Public IsWaitingForAction As Boolean

'Public Array1(), Array2(), Array3()

Sub AutoRefresh(flag As Boolean) ' This method aims to improve the speed of drawing shapes
    If Not flag Then
        Application.Optimization = True
        Application.EventsEnabled = False
    Else
        Application.Optimization = False
        Application.EventsEnabled = True
        Application.ActiveWindow.Refresh
    End If
End Sub

'Function GetDataRange(Optional title = "CGDK data selection", Optional Prompt = "Please select data from Excel Workbook.", Optional DefaultRange) As Object
'On Error GoTo veryend:
'
'Dim ExcelSheet As Object, Excel As Object
'
'Set ExcelSheet = CreateObject("Excel.Sheet")
'ExcelSheet.Parent.Windows(1).WindowState = -4137
'
'Set Excel = ExcelSheet.Application
'
'Set ExcelSheet = Nothing
'
'If Excel.Workbooks.Count < 1 Then
'    Dim q As Integer
'    q = MsgBox("No Excel Workbook is available! Do you want to open one now?", vbYesNo)
'    If q = 6 Then
'        If File.OpenExcelFile = 0 Then Exit Function
'    Else
'        Exit Function
'    End If
'End If
'
''AppActivate Excel.Caption
'Excel.ActiveWindow.WindowState = -4137
'
'
'If IsMissing(DefaultRange) Then
'    Set GetDataRange = Excel.InputBox(Prompt, title, , , , , , 8)
'Else
'    DefaultRange.Parent.Activate
'    Set GetDataRange = Excel.InputBox(Prompt, title, DefaultRange.Address, , , , , 8)
'End If
'
'AppActivate Application.AppWindow.Caption
'
'Exit Function
'
'veryend:
'
'AppActivate Application.AppWindow.Caption
'
'If Err.Number = 424 Then
'ElseIf Err.Number = -2147417846 Then
'    MsgBox "Can not access the Excel workbook! Maybe another process is using the software."
'Else
'    MsgBox Err.Description
'End If
'End Function


Function DrawCurve(ByVal x, ByVal y, Optional category = cdrCuspNode) As Shape
    On Error GoTo veryend
    Dim crv As Curve, sp As SubPath
    If (UBound(x) - LBound(x)) <> (UBound(y) - LBound(y)) Then GoTo veryend1
    If (UBound(x) - LBound(x)) < 1 Then GoTo veryend2
    
    'Common.AutoRefresh False
    
    Dim i As Long, n As Long
    n = 1

    Set crv = CreateCurve(ActiveDocument)
    
    For i = LBound(x) To UBound(x)
        If n = 1 Then
            Set sp = crv.CreateSubPath(x(i), y(i))
            n = n + 1
        Else
            sp.AppendCurveSegment x(i), y(i)
            n = n + 1
        End If
    Next i
    
    For i = 1 To crv.Nodes.Count
        Select Case category
            Case cdrSmoothNode
                crv.Nodes(i).Type = cdrSmoothNode
            Case cdrSymmetricalNode
                crv.Nodes(i).Type = cdrSymmetricalNode
        End Select
    Next i
    
    Set DrawCurve = ActiveLayer.CreateCurve(crv)
    'Common.AutoRefresh True
    Exit Function
veryend1:
    'MsgBox "Can not construct a curve! The numbers of x value and y value should be equivalent!"
    Exit Function
veryend2:
    'MsgBox "Can not construct a curve! It needs at least two x and y values to construct a curve!"
    Set DrawCurve = Nothing
    Exit Function
veryend:
    'Common.AutoRefresh True
    MsgBox Err.Description
End Function


Function DrawShapes(s As Shape, ByVal x, ByVal y, Optional align = cdrCenterAlign) As ShapeRange
    On Error GoTo veryend
    
    Dim sr As New ShapeRange, s1 As Shape, i As Long
    
    If (UBound(x) - LBound(x)) <> (UBound(y) - LBound(y)) Then GoTo veryend1
    If (UBound(x) - LBound(x)) < 0 Then GoTo veryend2
    
    'Common.AutoRefresh False
    
    For i = LBound(x) To UBound(x)
        Set s1 = s.Duplicate
        Common.MoveTo s1, x(i), y(i), align
        sr.Add s1
    Next i
    Set DrawShapes = sr
    'Common.AutoRefresh True
    Exit Function
veryend1:
    'MsgBox "Can not draw shapes! The numbers of x value and y value should be equivalent!"
    Exit Function
veryend2:
    'MsgBox "Can not draw shapes! It needs at least one x and y values to construct a shape!"
    Set DrawShapes = Nothing
    Exit Function
veryend:
    'Common.AutoRefresh True
    If Err.Number = 91 Or Err.Number = -2147221248 Then
        MsgBox "You have to set a legend first!"
    Else
        MsgBox Err.Description
    End If
End Function


Sub MoveTo(s As Shape, x, y, Optional align = cdrCenterAlign)
Dim temp As Integer
temp = ActiveDocument.ReferencePoint
Select Case align
    Case 0
        ActiveDocument.ReferencePoint = cdrCenter
    Case 1
        ActiveDocument.ReferencePoint = cdrTopLeft
    Case 2
        ActiveDocument.ReferencePoint = cdrTopRight
    Case 3
        ActiveDocument.ReferencePoint = cdrBottomLeft
    Case 4
        ActiveDocument.ReferencePoint = cdrBottomRight
    Case Else
        ActiveDocument.ReferencePoint = cdrCenter
End Select

s.SetPosition x, y

ActiveDocument.ReferencePoint = temp
End Sub


Sub ShapeAlign(s1 As Shape, s2 As Shape, Optional align1 = cdrCenterAlign, Optional align2 = cdrCenterAlign)
    Dim x, y, temp As Integer
    temp = ActiveDocument.ReferencePoint
    
    Select Case align2
        Case 0
            ActiveDocument.ReferencePoint = cdrCenter
        Case 1
            ActiveDocument.ReferencePoint = cdrTopLeft
        Case 2
            ActiveDocument.ReferencePoint = cdrTopRight
        Case 3
            ActiveDocument.ReferencePoint = cdrBottomLeft
        Case 4
            ActiveDocument.ReferencePoint = cdrBottomRight
        Case Else
            ActiveDocument.ReferencePoint = cdrCenter
    End Select
    
     x = s2.PositionX
     y = s2.PositionY
    
    Common.MoveTo s1, x, y, align1
    
    ActiveDocument.ReferencePoint = temp
End Sub


Function Replaceshapes(s As Shape, sr As ShapeRange, Optional sizetype = cdrSourceShapeSize, Optional height, Optional width, Optional align1 = cdrCenterAlign, Optional align2 = cdrCenterAlign) As ShapeRange
    On Error GoTo veryend
    
    'Common.AutoRefresh False
    
    Dim i As Long, s1 As Shape, sr1 As New ShapeRange
    For i = 1 To sr.Count
        Set s1 = s.Duplicate
            Select Case sizetype
                Case Common.cdrTargetShapeSize
                    s1.SizeHeight = sr(i).SizeHeight
                    s1.SizeWidth = sr(i).SizeWidth
                Case Common.cdrCustomSize
                    If Not IsMissing(width) And Not IsMissing(height) Then
                        s1.SizeHeight = height
                        s1.SizeWidth = width
                    End If
            End Select
        ShapeAlign s1, sr(i), align1, align2
        sr1.Add s1
    Next i
    'Common.AutoRefresh True
    sr.Delete
    Set Replaceshapes = sr1
    Exit Function
veryend:
    'Common.AutoRefresh True
    If Err.Number = 91 Or Err.Number = -2147221248 Then
        MsgBox "You have to set a source shape so that the program can replace the target shapes with it!"
    Else
        MsgBox Err.Description
    End If
End Function

Function FindShapeInShapeGroup(ShapeGroup As Shape, Name As String) As ShapeRange ' To find a shape by its name. The name is not case sensitive.
    Dim i As Long, sr As New ShapeRange
    For i = 1 To ShapeGroup.Shapes.Count
        If ShapeGroup.Shapes(i).Type = cdrGroupShape Then
            sr.AddRange FindShapeInShapeGroup(ShapeGroup.Shapes(i), Name)
        Else
            If VBA.LCase(ShapeGroup.Shapes(i).Name) = VBA.LCase(Name) Then sr.Add ShapeGroup.Shapes(i)
        End If
    Next i
    Set FindShapeInShapeGroup = sr
End Function

Function FindShapeByType(mysr As ShapeRange, mytype As String) As ShapeRange ' To find a shape by its name. The name is not case sensitive.
    Dim i As Long, sr As New ShapeRange
    For i = 1 To mysr.Count
        If ActiveDocument.DataFields.IsPresent("Type") Then
            If VBA.LCase(mysr.Item(i).ObjectData("Type").value) = VBA.LCase(mytype) Then sr.Add mysr.Item(i)
            If mysr.Item(i).Type = cdrGroupShape Then sr.AddRange FindShapeByType(mysr.Item(i).Shapes.All, mytype)
        End If
    Next i
    Set FindShapeByType = sr
End Function

'Sub ScatterCoordinateTranslate(ByVal x1, ByVal y1, Field As Shape)
'    Dim i As Long, logx As Boolean, logy As Boolean, x As Double, y As Double, w As Double, h As Double, left As Double, right As Double, top As Double, bottom As Double
'    Field.GetBoundingBox x, y, w, h
'
'    If LCase(Field.ObjectData("logx").Value) = "true" Then logx = True Else logx = False
'    If LCase(Field.ObjectData("logy").Value) = "true" Then logy = True Else logy = False
'    left = Val(Field.ObjectData("left_value").Value)
'    right = Val(Field.ObjectData("right_value").Value)
'    top = Val(Field.ObjectData("top_value").Value)
'    bottom = Val(Field.ObjectData("bottom_value").Value)
'
'    ReDim Array1(LBound(x1) To UBound(x1))
'    ReDim Array2(LBound(x1) To UBound(x1))
'    For i = LBound(x1) To UBound(x1)
'        If logx Then
'            Array1(i) = x + (Log(x1(i)) - Log(left)) / (right - left) * w
'        Else
'            Array1(i) = x + (x1(i) - left) / (right - left) * w
'        End If
'        If logy Then
'            Array2(i) = y + (Log(y1(i)) - Log(bottom)) / (top - bottom) * h
'        Else
'            Array2(i) = y + (y1(i) - bottom) / (top - bottom) * h
'        End If
'    Next i
'End Sub

Function HowManyElementsInArray(Array1) As Long
    On Error GoTo veryend1:
        HowManyElementsInArray = UBound(Array1) - LBound(Array1) + 1
    Exit Function
veryend1:
        HowManyElementsInArray = 0
End Function
