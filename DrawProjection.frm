VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DrawProjection 
   Caption         =   "Projection & diagram"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4320
   OleObjectBlob   =   "DrawProjection.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "DrawProjection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Azimuths As Object
Public Dips As Object
Public CurrentAction As Integer

Private Sub CommandButton1_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not Azimuths Is Nothing Then
        Dim temp As Double
        temp = Val(Azimuths.Count)
    End If
    If Azimuths Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set Azimuths = obj
    Else
        Set obj = GetDataRange.GetDataRange(Azimuths)
        If Not obj Is Nothing Then Set Azimuths = obj
    End If
    If Not Azimuths Is Nothing Then Me.TextBox2.value = Azimuths.Address
    Exit Sub
finally:
    TextBox2.value = ""
    Set Azimuths = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton2_Click()
    Set Azimuths = Nothing
    Me.TextBox2.value = ""
End Sub

Private Sub CommandButton3_Click()
    Dim obj As Object
    On Error GoTo finally
    If Not Dips Is Nothing Then
        Dim temp As Double
        temp = Val(Dips.Count)
    End If
    If Dips Is Nothing Then
        Set obj = GetDataRange.GetDataRange()
        If Not obj Is Nothing Then Set Dips = obj
    Else
        Set obj = GetDataRange.GetDataRange(Dips)
        If Not obj Is Nothing Then Set Dips = obj
    End If
    If Not Dips Is Nothing Then Me.TextBox3.value = Dips.Address
    Exit Sub
finally:
    TextBox3.value = ""
    Set Dips = Nothing
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton4_Click()
    Set Dips = Nothing
    Me.TextBox3.value = ""
End Sub

Private Sub CommandButton5_Click()
    Dim i As Long, d1 As Double, d2 As Double
    If Me.TextBox2.value = "" Then MsgBox "Insufficient data!": Exit Sub
    If Me.TextBox3.value = "" Then MsgBox "Insufficient data!": Exit Sub
    On Error GoTo finally
    If Azimuths.Count <> Dips.Count Then MsgBox "The number of azimuth (" & Azimuths.Count & ") is not equivalent to that of dips(" & Dips.Count & ".)": Exit Sub
    For i = 1 To Azimuths.Count
        If Dips.Item(i) >= 90 Or Dips.Item(i) <= 0 Then MsgBox "Some dips are out of range!": Exit Sub
    Next i
    If ActiveSelectionRange.Count <> 1 Then MsgBox "You have to select a circle!": Exit Sub
    On Error GoTo veryend1
    ActiveSelectionRange.Shapes(1).Ellipse.GetRadius d1, d2
    On Error GoTo 0
    If VBA.Abs(d1 - d2) > 0.00001 Then MsgBox "You have to select a circle!": Exit Sub
    Projection.DrawPolarStereographicProjection ActiveSelectionRange.Shapes(1)
    Exit Sub
veryend1:
    MsgBox "The shape is not a circle!"
    Exit Sub
finally:
    MsgBox "The Excel file may be closed."
End Sub

Private Sub CommandButton6_Click()
    Dim i As Long, d1 As Double, d2 As Double
    If Me.CheckBox4.value Then
        If Me.TextBox2.value = "" Then MsgBox "Insufficient data!": Exit Sub
        If Me.TextBox3.value = "" Then MsgBox "Insufficient data!": Exit Sub
    Else
        If Me.TextBox2.value = "" Then MsgBox "Insufficient data!": Exit Sub
    End If
    
    On Error GoTo finally
    
    If Not Me.TextBox3.value = "" Then
        If Azimuths.Count <> Dips.Count Then MsgBox "The numbers of azimuth (" & Azimuths.Count & ") is not equivalent to that of dips(" & Dips.Count & ".)": Exit Sub
        For i = 1 To Azimuths.Count
            If Dips.Item(i) >= 90 Or Dips.Item(i) <= 0 Then MsgBox "Some dips are out of range!": Exit Sub
        Next i
    End If
    
    If ActiveSelectionRange.Count <> 1 Then MsgBox "You have to select a circle!": Exit Sub
    On Error GoTo veryend1
    ActiveSelectionRange.Shapes(1).Ellipse.GetRadius d1, d2
    On Error GoTo 0
    If VBA.Abs(d1 - d2) > 0.00001 Then MsgBox "You have to select a circle!": Exit Sub
    Projection.DrawRoseDiagram ActiveSelectionRange.Shapes(1)
    Exit Sub
veryend1:
    MsgBox "The shape is not a circle!"
    Exit Sub
finally:
    MsgBox "The Excel file may be closed."
End Sub

Private Sub UserForm_Terminate()
    Unload data
End Sub
