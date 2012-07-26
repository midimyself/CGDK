VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlotOnMap 
   Caption         =   "Plot on map"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   OleObjectBlob   =   "PlotOnMap.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "PlotOnMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Names As Object
Public Latitudes As Object
Public Longitudes As Object
Public LegendSymbol As Shape

Private Sub CommandButton1_Click()
    If ActiveSelectionRange.Count <> 1 Then MsgBox "You have to select a shape!": Exit Sub
    Set LegendSymbol = ActiveSelectionRange.Item(1)
End Sub

Private Sub CommandButton10_Click()
    If Me.TextBox1.value = "" Then MsgBox "Insufficient data!": Exit Sub
    If Me.TextBox2.value = "" Then MsgBox "Insufficient data!": Exit Sub
    On Error GoTo finally
    If Latitudes.Count <> Longitudes.Count Then MsgBox "The number of latitudes (" & Latitudes.Count & ") is not equivalent to that of  longitudes (" & Longitudes.Count & ")": Exit Sub
    If Me.TextBox3.value <> "" Then
        If Latitudes.Count <> Names.Count Then MsgBox "The number of latitudes (" & Latitudes.Count & ") is not equivalent to that of  names (" & Names.Count & ")": Exit Sub
    End If
    Module_PlotOnMap.Draw1
    Exit Sub
veryend:
    MsgBox "The symbole may be removed from the document. Please reset one."
    Exit Sub
finally:
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton2_Click()
    If Me.TextBox1.value = "" Then MsgBox "Insufficient data!": Exit Sub
    If Me.TextBox2.value = "" Then MsgBox "Insufficient data!": Exit Sub
    On Error GoTo finally
    If Latitudes.Count <> Longitudes.Count Then MsgBox "The number of latitudes (" & Latitudes.Count & ") is not equivalent to that of  longitudes (" & Longitudes.Count & ")": Exit Sub
    If Me.TextBox3.value <> "" Then
        If Latitudes.Count <> Names.Count Then MsgBox "The number of latitudes (" & Latitudes.Count & ") is not equivalent to that of  names (" & Names.Count & ")": Exit Sub
    End If
    If LegendSymbol Is Nothing Then MsgBox "Please set a symbol shape.": Exit Sub
    On Error GoTo veryend
        Dim temp As String
        temp = LegendSymbol.Name
    On Error GoTo 0
    Module_PlotOnMap.Draw
    Exit Sub
veryend:
    MsgBox "The symbole may be removed from the document. Please reset one."
    Exit Sub
finally:
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton3_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not Names Is Nothing Then
        Dim temp As Double
        temp = Val(Names.Count)
    End If
        
    If Names Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set Names = obj
    Else
        Set obj = GetDataRange.GetDataRange(Names)
        If Not obj Is Nothing Then Set Names = obj
    End If
    If Not Names Is Nothing Then Me.TextBox3.value = Names.Address
    Exit Sub
finally:
    TextBox3.value = ""
    Set Names = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton4_Click()
    Me.TextBox3.value = ""
    Set Names = Nothing
End Sub

Private Sub CommandButton5_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not Latitudes Is Nothing Then
        Dim temp As Double
        temp = Val(Latitudes.Count)
    End If
    If Latitudes Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set Latitudes = obj
    Else
        Set obj = GetDataRange.GetDataRange(Latitudes)
        If Not obj Is Nothing Then Set Latitudes = obj
    End If
    If Not Latitudes Is Nothing Then Me.TextBox1.value = Latitudes.Address
    Exit Sub
finally:
    TextBox1.value = ""
    Set Latitudes = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton6_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not Longitudes Is Nothing Then
        Dim temp As Double
        temp = Val(Longitudes.Count)
    End If
    If Longitudes Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set Longitudes = obj
    Else
        Set obj = GetDataRange.GetDataRange(Longitudes)
        If Not obj Is Nothing Then Set Longitudes = obj
    End If
    If Not Longitudes Is Nothing Then Me.TextBox2.value = Longitudes.Address
    Exit Sub
finally:
    TextBox2.value = ""
    Set Longitudes = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton8_Click()
    Me.TextBox1.value = ""
    Set Latitudes = Nothing
End Sub

Private Sub CommandButton9_Click()
    Me.TextBox2.value = ""
    Set Longitudes = Nothing
End Sub

Private Sub UserForm_Click()

End Sub
