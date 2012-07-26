VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lithology 
   Caption         =   "Lithology"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   OleObjectBlob   =   "Lithology.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Lithology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    If ActiveSelectionRange.Count <> 1 Then MsgBox "Please select a legend.": Exit Sub
    If ActiveSelectionRange.Shapes(1).Type = cdrGroupShape Then MsgBox "The legend can not be a group shape.": Exit Sub
    If Me.ListBox1.ListIndex = -1 Then MsgBox "Please select a rock type from the listbox.": Exit Sub
    
    If Not DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.FindShape(Name:=Me.ListBox1.List(Me.ListBox1.ListIndex)) Is Nothing Then _
        DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.FindShape(Name:=Me.ListBox1.List(Me.ListBox1.ListIndex)).Delete
    ActiveSelectionRange.Shapes(1).Name = Me.ListBox1.List(Me.ListBox1.ListIndex)
    ActiveSelectionRange.Shapes(1).Copy
    DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.ActiveLayer.Paste
    DrawStratigraphicColumns.CGDKLithologyLegends.Save
    Me.ListBox1.Clear
    Me.ListBox2.Clear
    UserForm_Initialize
    UserForm_Activate
End Sub

Private Sub CommandButton2_Click()
    If Me.ListBox1.ListIndex = -1 Then MsgBox "Please select a rock type from the listbox.": Exit Sub
    Dim s As Shape
    Set s = DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.Shapes.FindShape(Name:=Me.ListBox1.List(Me.ListBox1.ListIndex))
    s.Delete
    DrawStratigraphicColumns.CGDKLithologyLegends.Save
    Me.ListBox1.Clear
    Me.ListBox2.Clear
    UserForm_Initialize
    UserForm_Activate
End Sub

Private Sub CommandButton3_Click()
    Unload Me
End Sub

Private Sub CommandButton4_Click()
    Dim i As Long
    For i = 0 To Me.ListBox1.ListCount - 1
        If Not DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.FindShape(Name:=Me.ListBox1.List(i)) Is Nothing Then _
        DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.FindShape(Name:=Me.ListBox1.List(i)).Delete
    Next i
    DrawStratigraphicColumns.CGDKLithologyLegends.Save
    Me.ListBox1.Clear
    Me.ListBox2.Clear
    UserForm_Initialize
    UserForm_Activate
End Sub

Private Sub CommandButton5_Click()
    If ActiveSelectionRange.Shapes.Count <> 1 Then MsgBox "Please select a legend!": Exit Sub
    If ActiveSelectionRange.Shapes(1).Type = cdrGroupShape Then MsgBox "The legend can not be a group shape.": Exit Sub
    Dim lithologyname As String, s As Shape
    Set s = ActiveSelectionRange.Shapes(1)
    lithologyname = InputBox("Please input the lithology of the legend.")
    On Error GoTo veryend:
    If Not DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.Shapes.FindShape(Name:=lithologyname) Is Nothing Then MsgBox "The lithology has been added.": Exit Sub
    s.Name = lithologyname
    s.Copy
    DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.ActiveLayer.Paste
    DrawStratigraphicColumns.CGDKLithologyLegends.Save
    Me.ListBox1.Clear
    Me.ListBox2.Clear
    UserForm_Initialize
    UserForm_Activate
    Exit Sub
veryend:
End Sub

Private Sub CommandButton6_Click()
    If Me.ListBox2.ListIndex < 0 Then MsgBox "Please select a lithology": Exit Sub
    Dim s As Shape
    Set s = DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.Shapes.FindShape(Name:=Me.ListBox2.List(Me.ListBox2.ListIndex))
    s.Delete
    DrawStratigraphicColumns.CGDKLithologyLegends.Save
    Me.ListBox1.Clear
    Me.ListBox2.Clear
    UserForm_Initialize
    UserForm_Activate
End Sub

Private Sub CommandButton7_Click()
    Dim i As Long
    For i = DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.Shapes.Count To 1 Step -1
        DrawStratigraphicColumns.CGDKLithologyLegends.ActiveLayer.Shapes(i).Delete
    Next i
    DrawStratigraphicColumns.CGDKLithologyLegends.Save
    Me.ListBox1.Clear
    Me.ListBox2.Clear
    UserForm_Initialize
    UserForm_Activate
End Sub

Private Sub UserForm_Activate()
    Dim i As Long, s As Shape, temp As String
    For i = 0 To Me.ListBox1.ListCount - 1
        Dim str As String
        str = Me.ListBox1.List(i)
        Set s = DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.Shapes.FindShape(Name:=str)
        If Not s Is Nothing Then
            Me.ListBox1.List(i, 1) = "√"
        End If
    Next i
    Exit Sub
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    For i = 1 To DrawStratigraphicColumns.lithos.Count
        Me.ListBox1.AddItem DrawStratigraphicColumns.lithos.Item(i)
    Next i
    For i = DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.Shapes.Count To 1 Step -1
        Me.ListBox2.AddItem DrawStratigraphicColumns.CGDKLithologyLegends.ActivePage.Shapes(i).Name
    Next i
    Me.ListBox1.ListIndex = -1
    Me.ListBox2.ListIndex = -1
End Sub

Private Sub UserForm_Terminate()
    DrawStratigraphicColumns.Show 0
End Sub
