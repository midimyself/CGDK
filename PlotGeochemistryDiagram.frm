VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlotGeochemistryDiagram 
   Caption         =   "Plot Geochemistry Diagram"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   OleObjectBlob   =   "PlotGeochemistryDiagram.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "PlotGeochemistryDiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TSeries As New Series
Dim XYSeries As New Series
Dim SSeries As New Series

Private Sub ListBox1_Change()
    On Error GoTo finally
    If Me.ListBox1.ListIndex < 0 Then Exit Sub
    Set TSeries = New Series
    TSeries.Name = Geochemistry.TriangularSeries.Item(Me.ListBox1.ListIndex + 1).Name
    Set TSeries.Dim1_Data = Geochemistry.TriangularSeries.Item(Me.ListBox1.ListIndex + 1).Dim1_Data
    Set TSeries.Dim2_Data = Geochemistry.TriangularSeries.Item(Me.ListBox1.ListIndex + 1).Dim2_Data
    Set TSeries.Dim3_Data = Geochemistry.TriangularSeries.Item(Me.ListBox1.ListIndex + 1).Dim3_Data
    Set TSeries.ElementNames = Geochemistry.TriangularSeries.Item(Me.ListBox1.ListIndex + 1).ElementNames
    Me.TT1.value = TSeries.Name
    If Not TSeries.ElementNames Is Nothing Then Me.TT2.value = TSeries.ElementNames.Address
    Me.TT3.value = TSeries.Dim1_Data.Address
    Me.TT4.value = TSeries.Dim2_Data.Address
    Me.TT5.value = TSeries.Dim3_Data.Address
    Exit Sub
finally:
    Dim i As Long
    Geochemistry.TriangularSeries.Remove (Me.ListBox1.ListIndex + 1)
    Me.ListBox1.Clear
    For i = 1 To Geochemistry.TriangularSeries.Count
        Me.ListBox1.AddItem Geochemistry.TriangularSeries.Item(i).Name
    Next i
    MsgBox "The Excel file may be closed."
End Sub

Private Sub ListBox2_Change()
    On Error GoTo finally
    If Me.ListBox2.ListIndex < 0 Then Exit Sub
    Set XYSeries = New Series
    XYSeries.Name = Geochemistry.XYScatterSeries.Item(Me.ListBox2.ListIndex + 1).Name
    Set XYSeries.Dim1_Data = Geochemistry.XYScatterSeries.Item(Me.ListBox2.ListIndex + 1).Dim1_Data
    Set XYSeries.Dim2_Data = Geochemistry.XYScatterSeries.Item(Me.ListBox2.ListIndex + 1).Dim2_Data
    Set XYSeries.ElementNames = Geochemistry.XYScatterSeries.Item(Me.ListBox2.ListIndex + 1).ElementNames
    Me.XYT1.value = XYSeries.Name
    If Not XYSeries.ElementNames Is Nothing Then Me.XYT2.value = XYSeries.ElementNames.Address
    Me.XYT3.value = XYSeries.Dim1_Data.Address
    Me.XYT4.value = XYSeries.Dim2_Data.Address
    Exit Sub
finally:
    Dim i As Long
    Geochemistry.XYScatterSeries.Remove (Me.ListBox2.ListIndex + 1)
    Me.ListBox2.Clear
    For i = 1 To Geochemistry.XYScatterSeries.Count
        Me.ListBox1.AddItem Geochemistry.XYScatterSeries.Item(i).Name
    Next i
    MsgBox "The Excel file may be closed."
End Sub

Private Sub ListBox3_Change()
    On Error GoTo finally
    If Me.ListBox3.ListIndex < 0 Then Exit Sub
    Set SSeries = New Series
    SSeries.Name = Geochemistry.SpiderSeries.Item(Me.ListBox3.ListIndex + 1).Name
    Set SSeries.Dim1_Data = Geochemistry.SpiderSeries.Item(Me.ListBox3.ListIndex + 1).Dim1_Data
    Set SSeries.ElementNames = Geochemistry.SpiderSeries.Item(Me.ListBox3.ListIndex + 1).ElementNames
    Me.ST1.value = SSeries.Name
    If Not SSeries.ElementNames Is Nothing Then Me.ST2.value = SSeries.ElementNames.Address
    Me.ST3.value = SSeries.Dim1_Data.Address
    Exit Sub
finally:
    Dim i As Long
    Geochemistry.SpiderSeries.Remove (Me.ListBox3.ListIndex + 1)
    Me.ListBox3.Clear
    For i = 1 To Geochemistry.SpiderSeries.Count
        Me.ListBox3.AddItem Geochemistry.SpiderSeries.Item(i).Name
    Next i
    MsgBox "The Excel file may be closed."
End Sub

Private Sub SA_Click()
    On Error GoTo finally
    If ST1.value = "" Then SSeries.Name = "untitled series"
    If ST3.value = "" Then MsgBox "Please input the Y values.": Exit Sub
    If SOption1.value Then 'elements in rows
        If SSeries.Dim1_Data.Columns.Count <> Geochemistry.HowManyColumns() Then MsgBox "The number of the columns of Y values(" & SSeries.Dim1_Data.Columns.Count & ") is not equivalent to the number of normalized values(" & Geochemistry.HowManyColumns() & ").": Exit Sub
        If ST2.value <> "" Then
            If SSeries.Dim1_Data.Rows.Count() <> SSeries.ElementNames.Count Then MsgBox "The number of the element names(" & SSeries.ElementNames.Count & ") is not equivalent to that of the rows of Y values(" & SSeries.Dim1_Data.Rows.Count & ").": Exit Sub
        End If
    Else 'elements in columns
        If SSeries.Dim1_Data.Rows.Count <> Geochemistry.HowManyColumns() Then MsgBox "The number of the rows of Y values(" & SSeries.Dim1_Data.Rows.Count & ") is not equivalent to the number of normalized values(" & Geochemistry.HowManyColumns() & ").": Exit Sub
        If ST2.value <> "" Then
            If SSeries.Dim1_Data.Columns.Count <> SSeries.ElementNames.Count Then MsgBox "The number of the element names(" & SSeries.ElementNames.Count & ") is not equivalent to that of the columns of Y values(" & SSeries.Dim1_Data.Columns.Count & ").": Exit Sub
        End If
    End If

    Geochemistry.SpiderSeries.Add SSeries

    ST1.value = ""
    ST2.value = ""
    ST3.value = ""
    SOption1.Enabled = False
    SOption2.Enabled = False
    Set SSeries = New Series

    Dim i As Long
    Me.ListBox3.Clear
    For i = 1 To Geochemistry.SpiderSeries.Count
        Me.ListBox3.AddItem Geochemistry.SpiderSeries.Item(i).Name
    Next i
    Exit Sub
finally:
    MsgBox "The Excel file may be closed."
End Sub

Private Sub SC1_Click()
    Dim obj As Object
    Set obj = GetDataRange.GetDataRange()
    If Not obj Is Nothing Then
        SSeries.Name = obj.Item(1)
        Me.ST1.value = SSeries.Name
    End If
End Sub
Private Sub SC2_Click()
    SSeries.Name = ""
    Me.ST1.value = ""
End Sub

Private Sub SC3_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not SSeries.ElementNames Is Nothing Then
        Dim temp As Double
        temp = Val(SSeries.ElementNames.Count)
    End If
    
    If SSeries.ElementNames Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set SSeries.ElementNames = obj
    Else
        Set obj = GetDataRange.GetDataRange(SSeries.ElementNames)
        If Not obj Is Nothing Then Set SSeries.ElementNames = obj
    End If
    If Not SSeries.ElementNames Is Nothing Then Me.ST2.value = SSeries.ElementNames.Address
    Exit Sub
finally:
    Me.ST2.value = ""
    Set SSeries.ElementNames = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub SC4_Click()
    Set SSeries.ElementNames = Nothing
    Me.ST2.value = ""
End Sub

Private Sub SC5_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not SSeries.Dim1_Data Is Nothing Then
        Dim temp As Double
        temp = Val(SSeries.Dim1_Data.Count)
    End If
    
    If SSeries.Dim1_Data Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then
            If obj.Areas.Count <> 1 Then GoTo veryend1
            Set SSeries.Dim1_Data = obj
        End If
    Else
        Set obj = GetDataRange.GetDataRange(SSeries.Dim1_Data)
        If Not obj Is Nothing Then
        If obj.Areas.Count <> 1 Then GoTo veryend1
        Set SSeries.Dim1_Data = obj
        End If
    End If
    If Not SSeries.Dim1_Data Is Nothing Then Me.ST3.value = SSeries.Dim1_Data.Address
    Exit Sub
veryend1:
    MsgBox "The data range can only contain one area."
    Exit Sub
finally:
    Me.ST3.value = ""
    Set SSeries.Dim1_Data = Nothing
    MsgBox "The Excel file may be closed."
End Sub
Private Sub SC6_Click()
    Set SSeries.Dim1_Data = Nothing
    Me.ST3.value = ""
End Sub

Private Sub SD_Click()
    If Me.ListBox3.ListIndex < 0 Then MsgBox "Please select a series.": Exit Sub
    Dim l As Long, i As Long
    l = Me.ListBox3.ListIndex
    Me.ST1.value = ""
    Me.ST2.value = ""
    Me.ST3.value = ""
    SSeries.Name = ""
    Set SSeries.ElementNames = Nothing
    Set SSeries.Dim1_Data = Nothing
    Geochemistry.SpiderSeries.Remove (l + 1)
    Me.ListBox3.Clear
    For i = 1 To Geochemistry.SpiderSeries.Count
        Me.ListBox3.AddItem Geochemistry.SpiderSeries.Item(i).Name
    Next i
    If Me.ListBox3.ListCount = 0 Then SOption1.Enabled = True: SOption2.Enabled = True
End Sub

Private Sub SM_Click()
    On Error GoTo finally
    If Me.ListBox3.ListIndex < 0 Then MsgBox "Please select a series.": Exit Sub
    If ST1.value = "" Then XYSeries.Name = "untitled series"
    If ST3.value = "" Then MsgBox "Please input the Y values.": Exit Sub
    
    If SOption1.value Then 'elements in rows
        If SSeries.Dim1_Data.Columns.Count <> Geochemistry.HowManyColumns() Then MsgBox "The number of the columns of Y values(" & SSeries.Dim1_Data.Columns.Count & ") is not equivalent to the number of normalized values(" & Geochemistry.HowManyColumns() & ").": Exit Sub
        If ST2.value <> "" Then
            If SSeries.Dim1_Data.Rows.Count() <> SSeries.ElementNames.Count Then MsgBox "The number of the element names(" & SSeries.ElementNames.Count & ") is not equivalent to that of the rows of Y values(" & SSeries.Dim1_Data.Rows.Count & ").": Exit Sub
        End If
    Else 'elements in columns
        If SSeries.Dim1_Data.Rows.Count <> Geochemistry.HowManyColumns() Then MsgBox "The number of the rows of Y values(" & SSeries.Dim1_Data.Rows.Count & ") is not equivalent to the number of normalized values(" & Geochemistry.HowManyColumns() & ").": Exit Sub
        If ST2.value <> "" Then
            If SSeries.Dim1_Data.Columns.Count <> SSeries.ElementNames.Count Then MsgBox "The number of the element names(" & SSeries.ElementNames.Count & ") is not equivalent to that of the columns of Y values(" & SSeries.Dim1_Data.Columns.Count & ").": Exit Sub
        End If
    End If
    
    Dim l As Long, i As Long
    l = Me.ListBox3.ListIndex + 1
    Geochemistry.SpiderSeries.Remove (l)
    Geochemistry.SpiderSeries.Add SSeries
    ST1.value = ""
    ST2.value = ""
    ST3.value = ""
    
    Set SSeries = New Series
    
    Me.ListBox3.Clear
    For i = 1 To Geochemistry.SpiderSeries.Count
        Me.ListBox3.AddItem Geochemistry.SpiderSeries.Item(i).Name
    Next i
    Exit Sub
finally:
    MsgBox "The Excel file may be closed."
End Sub

Private Sub SPlot_Click()
    Geochemistry.PlotSpiderDiagram
End Sub

Private Sub TA_Click()
    On Error GoTo finally
    If TT1.value = "" Then TSeries.Name = "untitled series"
    If TT3.value = "" Then MsgBox "Please input the top values.": Exit Sub
    If TT4.value = "" Then MsgBox "Please input the left values.": Exit Sub
    If TT5.value = "" Then MsgBox "Please input the right values.": Exit Sub
    If TSeries.Dim1_Data.Count <> TSeries.Dim2_Data.Count Then MsgBox "The number of top values(" & TSeries.Dim1_Data.Count & ") is not equivalent to that of left values(" & TSeries.Dim2_Data.Count & ").": Exit Sub
    If TSeries.Dim1_Data.Count <> TSeries.Dim3_Data.Count Then MsgBox "The number of top values(" & TSeries.Dim1_Data.Count & ") is not equivalent to that of right values(" & TSeries.Dim3_Data.Count & ").": Exit Sub
    If TT2.value <> "" Then
        If TSeries.Dim1_Data.Count <> TSeries.ElementNames.Count Then MsgBox "The number of top values(" & TSeries.Dim1_Data.Count & ") is not equivalent to that of element names(" & TSeries.ElementNames.Count & ").": Exit Sub
    End If
    
    Geochemistry.TriangularSeries.Add TSeries

    TT1.value = ""
    TT2.value = ""
    TT3.value = ""
    TT4.value = ""
    TT5.value = ""
    
    
    Set TSeries = New Series

    Dim i As Long
    Me.ListBox1.Clear
    For i = 1 To Geochemistry.TriangularSeries.Count
        Me.ListBox1.AddItem Geochemistry.TriangularSeries.Item(i).Name
    Next i
    Exit Sub
finally:
    MsgBox "The Excel file may be closed."
End Sub

Private Sub TC1_Click()
    Dim obj As Object
    Set obj = GetDataRange.GetDataRange()
    If Not obj Is Nothing Then
        TSeries.Name = obj.Item(1)
        Me.TT1.value = TSeries.Name
    End If
End Sub

Private Sub TC2_Click()
    TSeries.Name = ""
    Me.TT1.value = ""
End Sub

Private Sub TC3_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not TSeries.ElementNames Is Nothing Then
        Dim temp As Double
        temp = Val(TSeries.ElementNames.Count)
    End If
    
    If TSeries.ElementNames Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set TSeries.ElementNames = obj
    Else
        Set obj = GetDataRange.GetDataRange(TSeries.ElementNames)
        If Not obj Is Nothing Then Set TSeries.ElementNames = obj
    End If
    If Not TSeries.ElementNames Is Nothing Then Me.TT2.value = TSeries.ElementNames.Address
    Exit Sub
finally:
    Me.TT2.value = ""
    Set TSeries.ElementNames = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub TC4_Click()
    Set TSeries.ElementNames = Nothing
    Me.TT2.value = ""
End Sub

Private Sub TC5_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not TSeries.Dim1_Data Is Nothing Then
        Dim temp As Double
        temp = Val(TSeries.Dim1_Data.Count)
    End If
    If TSeries.Dim1_Data Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set TSeries.Dim1_Data = obj
    Else
        Set obj = GetDataRange.GetDataRange(TSeries.Dim1_Data)
        If Not obj Is Nothing Then Set TSeries.Dim1_Data = obj
    End If
    If Not TSeries.Dim1_Data Is Nothing Then Me.TT3.value = TSeries.Dim1_Data.Address
    Exit Sub
finally:
    Me.TT3.value = ""
    Set TSeries.Dim1_Data = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub TC6_Click()
    Set TSeries.Dim1_Data = Nothing
    Me.TT3.value = ""
End Sub

Private Sub TC7_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not TSeries.Dim2_Data Is Nothing Then
        Dim temp As Double
        temp = Val(TSeries.Dim2_Data.Count)
    End If
    If TSeries.Dim2_Data Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set TSeries.Dim2_Data = obj
    Else
        Set obj = GetDataRange.GetDataRange(TSeries.Dim2_Data)
        If Not obj Is Nothing Then Set TSeries.Dim2_Data = obj
    End If
    If Not TSeries.Dim2_Data Is Nothing Then Me.TT4.value = TSeries.Dim2_Data.Address
    Exit Sub
finally:
    Me.TT4.value = ""
    Set TSeries.Dim2_Data = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub TC8_Click()
    Set TSeries.Dim2_Data = Nothing
    Me.TT4.value = ""
End Sub

Private Sub TC9_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not TSeries.Dim3_Data Is Nothing Then
        Dim temp As Double
        temp = Val(TSeries.Dim3_Data.Count)
    End If
    If TSeries.Dim3_Data Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set TSeries.Dim3_Data = obj
    Else
        Set obj = GetDataRange.GetDataRange(TSeries.Dim3_Data)
        If Not obj Is Nothing Then Set TSeries.Dim3_Data = obj
    End If
    If Not TSeries.Dim3_Data Is Nothing Then Me.TT5.value = TSeries.Dim3_Data.Address
    Exit Sub
finally:
    Me.TT5.value = ""
    Set TSeries.Dim3_Data = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub TC10_Click()
    Set TSeries.Dim3_Data = Nothing
    Me.TT5.value = ""
End Sub

Private Sub TD_Click()
    If Me.ListBox1.ListIndex < 0 Then MsgBox "Please select a series.": Exit Sub
    Dim l As Long, i As Long
    l = Me.ListBox1.ListIndex
    Me.TT1.value = ""
    Me.TT2.value = ""
    Me.TT3.value = ""
    Me.TT4.value = ""
    Me.TT5.value = ""
    TSeries.Name = ""
    Set TSeries.ElementNames = Nothing
    Set TSeries.Dim1_Data = Nothing
    Set TSeries.Dim2_Data = Nothing
    Set TSeries.Dim3_Data = Nothing
    Geochemistry.TriangularSeries.Remove (l + 1)
    Me.ListBox1.Clear
    For i = 1 To Geochemistry.TriangularSeries.Count
        Me.ListBox1.AddItem Geochemistry.TriangularSeries.Item(i).Name
    Next i
End Sub

Private Sub TM_Click()
    On Error GoTo finally
    If Me.ListBox1.ListIndex < 0 Then MsgBox "Please select a series.": Exit Sub
    If TT1.value = "" Then TSeries.Name = "untitled series"
    If TT3.value = "" Then MsgBox "Please input the top values.": Exit Sub
    If TT4.value = "" Then MsgBox "Please input the left values.": Exit Sub
    If TT5.value = "" Then MsgBox "Please input the right values.": Exit Sub
    If TSeries.Dim1_Data.Count <> TSeries.Dim2_Data.Count Then MsgBox "The number of top values(" & TSeries.Dim1_Data.Count & ") is not equivalent to that of left values(" & TSeries.Dim2_Data.Count & ").": Exit Sub
    If TSeries.Dim1_Data.Count <> TSeries.Dim3_Data.Count Then MsgBox "The number of top values(" & TSeries.Dim1_Data.Count & ") is not equivalent to that of right values(" & TSeries.Dim3_Data.Count & ").": Exit Sub
    If TT2.value <> "" Then
        If TSeries.Dim1_Data.Count <> TSeries.ElementNames.Count Then MsgBox "The number of top values(" & TSeries.Dim1_Data.Count & ") is not equivalent to that of element names(" & TSeries.ElementNames.Count & ").": Exit Sub
    End If
    
    Dim l As Long, i As Long
    l = Me.ListBox1.ListIndex + 1
    Geochemistry.TriangularSeries.Remove (l)
    Geochemistry.TriangularSeries.Add TSeries
    TT1.value = ""
    TT2.value = ""
    TT3.value = ""
    TT4.value = ""
    TT5.value = ""
    
    Set TSeries = New Series
    
    Me.ListBox1.Clear
    For i = 1 To Geochemistry.TriangularSeries.Count
        Me.ListBox1.AddItem Geochemistry.TriangularSeries.Item(i).Name
    Next i
    Exit Sub
finally:
    MsgBox "The Excel file may be closed."
End Sub

Private Sub TPlot_Click()
    Geochemistry.PlotTriangularDiagram
End Sub

Private Sub UserForm_Initialize()
    API.StartTimer (150)
    Me.Frame1.Visible = False
    Me.Frame2.Visible = False
    Me.Frame3.Visible = False
    Dim i As Long
    For i = 1 To Geochemistry.SpiderSeries.Count
        Geochemistry.SpiderSeries.Remove (1)
    Next i
    For i = 1 To Geochemistry.TriangularSeries.Count
        Geochemistry.TriangularSeries.Remove (1)
    Next i
    For i = 1 To Geochemistry.XYScatterSeries.Count
        Geochemistry.XYScatterSeries.Remove (1)
    Next i
End Sub

Private Sub UserForm_Terminate()
    API.StopTimer
End Sub

Private Sub XYA_Click()
    On Error GoTo finally
    If XYT1.value = "" Then XYSeries.Name = "untitled series"
    If XYT3.value = "" Then MsgBox "Please input the X values.": Exit Sub
    If XYT4.value = "" Then MsgBox "Please input the Y values.": Exit Sub
    If XYSeries.Dim1_Data.Count <> XYSeries.Dim2_Data.Count Then MsgBox "The number of X values(" & XYSeries.Dim1_Data.Count & ") is not equivalent to that of Y values(" & XYSeries.Dim2_Data.Count & ").": Exit Sub
    If XYT2.value <> "" Then
        If XYSeries.Dim1_Data.Count <> XYSeries.ElementNames.Count Then MsgBox "The number of X values(" & XYSeries.Dim1_Data.Count & ") is not equivalent to that of element names(" & XYSeries.ElementNames.Count & ").": Exit Sub
    End If
    
    Geochemistry.XYScatterSeries.Add XYSeries

    XYT1.value = ""
    XYT2.value = ""
    XYT3.value = ""
    XYT4.value = ""
    Set XYSeries = New Series

    Dim i As Long
    Me.ListBox2.Clear
    For i = 1 To Geochemistry.XYScatterSeries.Count
        Me.ListBox2.AddItem Geochemistry.XYScatterSeries.Item(i).Name
    Next i
    Exit Sub
finally:
    MsgBox "The Excel file may be closed."
End Sub

Private Sub XYC1_Click()
    Dim obj As Object
    Set obj = GetDataRange.GetDataRange()
    If Not obj Is Nothing Then
        XYSeries.Name = obj.Item(1)
        Me.XYT1.value = XYSeries.Name
    End If
End Sub

Private Sub XYC2_Click()
    XYSeries.Name = ""
    Me.XYT1.value = ""
End Sub

Private Sub XYC3_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not XYSeries.ElementNames Is Nothing Then
        Dim temp As Double
        temp = Val(XYSeries.ElementNames.Count)
    End If
    If XYSeries.ElementNames Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set XYSeries.ElementNames = obj
    Else
        Set obj = GetDataRange.GetDataRange(XYSeries.ElementNames)
        If Not obj Is Nothing Then Set XYSeries.ElementNames = obj
    End If
    If Not XYSeries.ElementNames Is Nothing Then Me.XYT2.value = XYSeries.ElementNames.Address
    Exit Sub
finally:
    Me.XYT2.value = ""
    Set XYSeries.ElementNames = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub XYC4_Click()
    Set XYSeries.ElementNames = Nothing
    Me.XYT2.value = ""
End Sub

Private Sub XYC5_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not XYSeries.Dim1_Data Is Nothing Then
        Dim temp As Double
        temp = Val(XYSeries.Dim1_Data.Count)
    End If
    If XYSeries.Dim1_Data Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set XYSeries.Dim1_Data = obj
    Else
        Set obj = GetDataRange.GetDataRange(XYSeries.Dim1_Data)
        If Not obj Is Nothing Then Set XYSeries.Dim1_Data = obj
    End If
    If Not XYSeries.Dim1_Data Is Nothing Then Me.XYT3.value = XYSeries.Dim1_Data.Address
    Exit Sub
finally:
    Me.XYT3.value = ""
    Set XYSeries.Dim1_Data = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub XYC6_Click()
    Set XYSeries.Dim1_Data = Nothing
    Me.XYT3.value = ""
End Sub

Private Sub XYC7_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not XYSeries.Dim2_Data Is Nothing Then
        Dim temp As Double
        temp = Val(XYSeries.Dim2_Data.Count)
    End If
    If XYSeries.Dim2_Data Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set XYSeries.Dim2_Data = obj
    Else
        Set obj = GetDataRange.GetDataRange(XYSeries.Dim2_Data)
        If Not obj Is Nothing Then Set XYSeries.Dim2_Data = obj
    End If
    If Not XYSeries.Dim2_Data Is Nothing Then Me.XYT4.value = XYSeries.Dim2_Data.Address
    Exit Sub
finally:
    Me.XYT4.value = ""
    Set XYSeries.Dim2_Data = Nothing
    MsgBox "The Excel file may be closed."
End Sub
Private Sub XYC8_Click()
    Set XYSeries.Dim2_Data = Nothing
    Me.XYT4.value = ""
End Sub

Private Sub XYD_Click()
    If Me.ListBox2.ListIndex < 0 Then MsgBox "Please select a series.": Exit Sub
    Dim l As Long, i As Long
    l = Me.ListBox2.ListIndex
    Me.XYT1.value = ""
    Me.XYT2.value = ""
    Me.XYT3.value = ""
    Me.XYT4.value = ""
    XYSeries.Name = ""
    Set XYSeries.ElementNames = Nothing
    Set XYSeries.Dim1_Data = Nothing
    Set XYSeries.Dim2_Data = Nothing
    Geochemistry.XYScatterSeries.Remove (l + 1)
    Me.ListBox2.Clear
    For i = 1 To Geochemistry.XYScatterSeries.Count
        Me.ListBox2.AddItem Geochemistry.XYScatterSeries.Item(i).Name
    Next i
End Sub

Private Sub XYM_Click()
    On Error GoTo finally
    If Me.ListBox2.ListIndex < 0 Then MsgBox "Please select a series.": Exit Sub
    If XYT1.value = "" Then XYSeries.Name = "untitled series"
    If XYT3.value = "" Then MsgBox "Please input the X values.": Exit Sub
    If XYT4.value = "" Then MsgBox "Please input the Y values.": Exit Sub
    If XYSeries.Dim1_Data.Count <> XYSeries.Dim2_Data.Count Then MsgBox "The number of X values(" & XYSeries.Dim1_Data.Count & ") is not equivalent to that of Y values(" & XYSeries.Dim2_Data.Count & ").": Exit Sub
    If XYT2.value <> "" Then
        If XYSeries.Dim1_Data.Count <> XYSeries.ElementNames.Count Then MsgBox "The number of X values(" & XYSeries.Dim1_Data.Count & ") is not equivalent to that of element names(" & XYSeries.ElementNames.Count & ").": Exit Sub
    End If
    
    Dim l As Long, i As Long
    l = Me.ListBox2.ListIndex + 1
    Geochemistry.XYScatterSeries.Remove (l)
    Geochemistry.XYScatterSeries.Add XYSeries
    XYT1.value = ""
    XYT2.value = ""
    XYT3.value = ""
    XYT4.value = ""
    
    Set XYSeries = New Series
    
    Me.ListBox2.Clear
    For i = 1 To Geochemistry.XYScatterSeries.Count
        Me.ListBox2.AddItem Geochemistry.XYScatterSeries.Item(i).Name
    Next i
    Exit Sub
finally:
    MsgBox "The Excel file may be closed."
End Sub

Private Sub XYPlot_Click()
    Geochemistry.PlotXYScatterDiagram
End Sub
