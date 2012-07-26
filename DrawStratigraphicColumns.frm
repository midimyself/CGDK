VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DrawStratigraphicColumns 
   Caption         =   "Draw Stratigraphic Columns"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   OleObjectBlob   =   "DrawStratigraphicColumns.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "DrawStratigraphicColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CGDKLithologyLegends As Document
Public Thickness As Object, Grainsize As Object, Lith_Pos_Negative As Object, lithos As New Collection, legends As New KeyCollection


Private Sub CheckBox1_Click()
    If Grainsize Is Nothing Then
        Me.TextBox8.value = ""
        Me.TextBox9.value = ""
        Exit Sub
    End If
    If Me.CheckBox1.value Then
        Me.TextBox8.value = Stratigraphic.GetMinWidth / 10
        Me.TextBox9.value = Stratigraphic.GetMaxWidth
    Else
        Me.TextBox8.value = 0
        Me.TextBox9.value = Stratigraphic.GetMaxWidth
    End If
End Sub

Private Sub ComboBox1_Change()
Dim i As Long
Select Case Me.ComboBox1.ListIndex
    Case 0
        Me.CommandButton7.Enabled = True
        Label1.Caption = "Thickness:"
        Label2.Caption = "Grain size:"
        Label3.Caption = "Lithologic:"
        Label1.Visible = True
        Label2.Visible = True
        Label3.Visible = True
        TextBox1.Visible = True
        TextBox2.Visible = True
        TextBox3.Visible = True
        CommandButton1.Visible = True
        CommandButton2.Visible = True
        CommandButton3.Visible = True
        CommandButton4.Visible = True
        CommandButton5.Visible = True
        CommandButton6.Visible = True
        Set Grainsize = Nothing
        Set Lith_Pos_Negative = Nothing
        Me.TextBox2.value = ""
        Me.TextBox3.value = ""
        For i = 1 To lithos.Count
            lithos.Remove (1)
        Next i
    Case 1
        Me.CommandButton7.Enabled = False
        Label1.Caption = "Thickness:"
        Label2.Caption = "X value:"
        Label3.Caption = ""
        Label1.Visible = True
        Label2.Visible = True
        Label3.Visible = False
        TextBox1.Visible = True
        TextBox2.Visible = True
        TextBox3.Visible = False
        CommandButton1.Visible = True
        CommandButton2.Visible = True
        CommandButton3.Visible = True
        CommandButton4.Visible = True
        CommandButton5.Visible = False
        CommandButton6.Visible = False
        Set Grainsize = Nothing
        Set Lith_Pos_Negative = Nothing
        Me.TextBox2.value = ""
        Me.TextBox3.value = ""
        For i = 1 To lithos.Count
            lithos.Remove (1)
        Next i
    Case 2
        Me.CommandButton7.Enabled = False
        Label1.Caption = "Thickness:"
        Label2.Caption = "X value:"
        Label3.Caption = ""
        Label1.Visible = True
        Label2.Visible = True
        Label3.Visible = False
        TextBox1.Visible = True
        TextBox2.Visible = True
        TextBox3.Visible = False
        CommandButton1.Visible = True
        CommandButton2.Visible = True
        CommandButton3.Visible = True
        CommandButton4.Visible = True
        CommandButton5.Visible = False
        CommandButton6.Visible = False
        Set Lith_Pos_Negative = Nothing
        Set Grainsize = Nothing
        Set Lith_Pos_Negative = Nothing
        Me.TextBox2.value = ""
        Me.TextBox3.value = ""
        For i = 1 To lithos.Count
            lithos.Remove (1)
        Next i
    Case 3
        Me.CommandButton7.Enabled = True
        Label1.Caption = "Thickness:"
        Label2.Caption = ""
        Label3.Caption = "Polarity:"
        Label1.Visible = True
        Label2.Visible = False
        Label3.Visible = True
        TextBox1.Visible = True
        TextBox2.Visible = False
        TextBox3.Visible = True
        CommandButton1.Visible = True
        CommandButton2.Visible = True
        CommandButton3.Visible = False
        CommandButton4.Visible = False
        CommandButton5.Visible = True
        CommandButton6.Visible = True
        Set Grainsize = Nothing
        Set Lith_Pos_Negative = Nothing
        Me.TextBox2.value = ""
        Me.TextBox3.value = ""
        For i = 1 To lithos.Count
            lithos.Remove (1)
        Next i
End Select
End Sub

Private Sub CommandButton1_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not Thickness Is Nothing Then
        Dim temp As Double
        temp = Val(Thickness.Count)
    End If
    If Thickness Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set Thickness = obj
    Else
        Set obj = GetDataRange.GetDataRange(Thickness)
        If Not obj Is Nothing Then Set Thickness = obj
    End If
    If Not Thickness Is Nothing Then Me.TextBox1.value = Thickness.Address
    Exit Sub
finally:
    TextBox1.value = ""
    Set Thickness = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton10_Click()
    If Me.TextBox2.value = "" Then MsgBox "Please input the grain size data.": Exit Sub
    If ActiveSelectionRange.Count <> 1 Then MsgBox "Please select a rectangle shape.": Exit Sub
    If ActiveSelectionRange.Shapes(1).SizeWidth < 0.01 Or ActiveSelectionRange.Shapes(1).SizeHeight < 0.01 Then MsgBox "The rectangle is too small."
    Dim n As Long
    Stratigraphic.DrawXAxis
End Sub

Private Sub CommandButton2_Click()
    Set Thickness = Nothing
    Me.TextBox1.value = ""
End Sub

Private Sub CommandButton3_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not Grainsize Is Nothing Then
        Dim temp As Double
        temp = Val(Grainsize.Count)
    End If
    If Grainsize Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set Grainsize = obj
    Else
        Set obj = GetDataRange.GetDataRange(Grainsize)
        If Not obj Is Nothing Then Set Grainsize = obj
    End If
    If Not Grainsize Is Nothing Then Me.TextBox2.value = Grainsize.Address
    Exit Sub
finally:
    TextBox2.value = ""
    Set Grainsize = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton4_Click()
    Set Grainsize = Nothing
    Me.TextBox2.value = ""
End Sub

Private Sub CommandButton5_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not Lith_Pos_Negative Is Nothing Then
        Dim temp As Double
        temp = Val(Lith_Pos_Negative.Count)
    End If
    If Lith_Pos_Negative Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set Lith_Pos_Negative = obj
    Else
        Set obj = GetDataRange.GetDataRange(Lith_Pos_Negative)
        If Not obj Is Nothing Then Set Lith_Pos_Negative = obj
    End If
    If Not Lith_Pos_Negative Is Nothing Then Me.TextBox3.value = Lith_Pos_Negative.Address
    Exit Sub
finally:
    TextBox3.value = ""
    Set Lith_Pos_Negative = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton6_Click()
    Set Lith_Pos_Negative = Nothing
    Me.TextBox3.value = ""
    Dim i As Long
    For i = 1 To lithos.Count
        lithos.Remove 1
    Next i
End Sub

Private Sub CommandButton7_Click()
    Me.Hide
    Lithology.Show 0
End Sub

Private Sub CommandButton8_Click()
    Dim i As Long, s As Shape
    'On Error GoTo finally
    If Me.TextBox1.value = "" Then MsgBox "Please input the thickness data.": Exit Sub
    If ActiveSelectionRange.Count <> 1 Then MsgBox "Please select a rectangle shape.": Exit Sub
    If ActiveSelectionRange.Shapes(1).SizeWidth < 0.01 Or ActiveSelectionRange.Shapes(1).SizeHeight < 0.01 Then MsgBox "The rectangle is too small."
    Dim n As Long
    n = Thickness.Count
    If Me.TextBox2.value <> "" Then
        If n <> Grainsize.Count Then MsgBox "The number of thickness data is not equivalent to that of grain size.": Exit Sub
    End If

    If Me.TextBox3.value <> "" Then
        If n <> Lith_Pos_Negative.Count Then MsgBox "The number of thickness data is not equivalent to that of lithology.": Exit Sub
    End If
    Stratigraphic.Draw
    Exit Sub
finally:
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton9_Click()
    If Me.TextBox1.value = "" Then MsgBox "Please input the thickness data.": Exit Sub
    If ActiveSelectionRange.Count <> 1 Then MsgBox "Please select a rectangle shape.": Exit Sub
    If ActiveSelectionRange.Shapes(1).SizeWidth < 0.01 Or ActiveSelectionRange.Shapes(1).SizeHeight < 0.01 Then MsgBox "The rectangle is too small."
    Dim n As Long
    Stratigraphic.DrawYAxis
End Sub

Private Sub TextBox1_Change()
    On Error GoTo finally
    If Thickness Is Nothing Then
        Me.Label11.Caption = "Total"
        Me.Label12.Caption = "Number"
        Me.TextBox5.value = ""
        Exit Sub
    End If
    Dim i As Long, s As Double, n As Long
    s = 0
    n = 0
    For i = 1 To Thickness.Count
        s = s + VBA.Abs(Val(Thickness.Item(i).value))
        n = n + 1
    Next i
    Me.Label11.Caption = "Total=" & s
    Me.Label12.Caption = "Number=" & n
    Me.TextBox5.value = Val(Me.TextBox4.value) + Stratigraphic.GetTotalThickness
    Exit Sub
finally:
    TextBox1.value = ""
    Set Thickness = Nothing
    'MsgBox "The Excel file may be closed."
End Sub

Private Sub TextBox2_Change()
    On Error GoTo finally
    If Grainsize Is Nothing Then
        Me.TextBox8.value = ""
        Me.TextBox9.value = ""
        Exit Sub
    End If
    If Me.CheckBox1.value Then
        Me.TextBox8.value = Stratigraphic.GetMinWidth / 10
        Me.TextBox9.value = Stratigraphic.GetMaxWidth
    Else
        Me.TextBox8.value = 0
        Me.TextBox9.value = Stratigraphic.GetMaxWidth
    End If
    Exit Sub
finally:
    TextBox1.value = ""
    Set Grainsize = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub TextBox3_Change()
    On Error GoTo finally
    Dim i As Long, j As Long, existed As Boolean
    If Me.TextBox3.value = "" Then Exit Sub
    For i = 1 To lithos.Count
        lithos.Remove (1)
    Next i
    existed = False
    Dim v As Variant
    For Each v In Lith_Pos_Negative
        For j = 1 To lithos.Count
            If lithos.Item(j) = v Then existed = True: Exit For
        Next j
        If Not existed Then lithos.Add v
        existed = False
    Next v
    Exit Sub
finally:
    TextBox1.value = ""
    Set Lith_Pos_Negative = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub TextBox4_Change()
    Me.TextBox5.value = Val(Me.TextBox4.value) + Stratigraphic.GetTotalThickness
End Sub

Private Sub UserForm_Initialize()
    Me.ComboBox1.AddItem "Common stratigraphic column"
    Me.ComboBox1.AddItem "Polyline column"
    Me.ComboBox1.AddItem "Smooth line column"
    Me.ComboBox1.AddItem "Palo-geomagnetism column"
    Me.Label11.Caption = "Total"
    Me.Label12.Caption = "Number"
    Dim i As Long, window1 As Window
    Set window1 = Application.ActiveWindow
    On Error GoTo veryend1
    Set CGDKLithologyLegends = Application.OpenDocument(Application.Path & "gms\CGDKLithologyLegends.CDR")
    On Error GoTo 0
    window1.Activate
    Exit Sub
veryend1:
    Dim d As Document
    Set d = Application.CreateDocument
    d.SaveAs Application.Path & "gms\CGDKLithologyLegends.CDR"
    Resume 0
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    CGDKLithologyLegends.Close
End Sub
