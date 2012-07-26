VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateManager 
   Caption         =   "Template Manager"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5010
   OleObjectBlob   =   "TemplateManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "TemplateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CGDKtemplates As Document

Public Function IsExisted(Name As String, LayerOrShape As String) As Boolean
Dim i As Long
If VBA.LCase(LayerOrShape) = "layer" Then
    For i = 1 To CGDKtemplates.Pages.Item(1).Layers.Count
        If VBA.LCase(CGDKtemplates.Pages.Item(1).Layers(i).Name) = VBA.LCase(Name) Then
            IsExisted = True
            Exit Function
        End If
    Next i
Else
    For i = 1 To CGDKtemplates.Pages.Item(1).Layers.Item(Me.ComboBox1.value).Shapes.Count
        If VBA.LCase(CGDKtemplates.Pages.Item(1).Layers.Item(Me.ComboBox1.value).Shapes(i).Name) = VBA.LCase(Name) Then
            IsExisted = True
            Exit Function
        End If
    Next i
End If
End Function

Private Sub ComboBox1_Change()
    Me.ListBox1.Clear
    If Me.ComboBox1.ListCount > 0 Then
        Dim i As Long
        For i = CGDKtemplates.Pages.Item(1).Layers.Item(Me.ComboBox1.value).Shapes.Count To 1 Step -1
            Me.ListBox1.AddItem CGDKtemplates.Pages.Item(1).Layers.Item(Me.ComboBox1.value).Shapes(i).Name
        Next i
        If Me.ListBox1.ListCount > 0 Then
            Me.ListBox1.ListIndex = 0
        End If
    End If
End Sub

Private Sub ComboBox2_Change()
    Me.CommandButton1.Caption = Me.ComboBox2.value
End Sub

Private Sub CommandButton1_Click()
    Dim i As Long
    Dim GroupName As String
    If Me.CommandButton1.Caption = "Add" Then
        GroupName = InputBox("Please input the name of the new group.")
        If GroupName = "" Then Exit Sub
        If IsExisted(GroupName, "layer") Then
            MsgBox "The group is existed! Please change a group name."
            Exit Sub
        End If
        CGDKtemplates.Pages.Item(1).CreateLayer GroupName
        CGDKtemplates.Save
        Me.ComboBox1.Clear
        For i = CGDKtemplates.Pages.Item(1).Layers.Count To 1 Step -1
            Me.ComboBox1.AddItem CGDKtemplates.Pages.Item(1).Layers(i).Name
        Next i
        Me.ComboBox1.value = GroupName
    ElseIf Me.CommandButton1.Caption = "Modify" Then
        If Me.ComboBox1.ListCount < 1 Then Exit Sub
        GroupName = InputBox("Please input the name of the group.")
        If GroupName = "" Then Exit Sub
        If IsExisted(GroupName, "layer") Then
            MsgBox "The group is existed! Please change a group name."
            Exit Sub
        End If
        CGDKtemplates.Pages.Item(1).Layers(Me.ComboBox1.value).Name = GroupName
        CGDKtemplates.Save
        Me.ComboBox1.Clear
        For i = CGDKtemplates.Pages.Item(1).Layers.Count To 1 Step -1
            Me.ComboBox1.AddItem CGDKtemplates.Pages.Item(1).Layers(i).Name
        Next i
        Me.ComboBox1.value = GroupName
    Else
        If Me.ComboBox1.ListCount < 1 Then Exit Sub
        Dim flag As Integer
        flag = MsgBox("If you delete the group, all the templates in this group will be deleted at the same time." & VBA.Chr(13) & "Are you sure to delete the group?", vbYesNo)
        If flag = 6 Then
            CGDKtemplates.Pages.Item(1).Layers(Me.ComboBox1.value).Delete
            CGDKtemplates.Save
            Me.ComboBox1.Clear
            For i = CGDKtemplates.Pages.Item(1).Layers.Count To 1 Step -1
                Me.ComboBox1.AddItem CGDKtemplates.Pages.Item(1).Layers(i).Name
            Next i
            If Me.ComboBox1.ListCount > 0 Then Me.ComboBox1.value = Me.ComboBox1.List(0)
        End If
    End If
End Sub

Private Sub CommandButton2_Click()
    Dim s As Shape
    If Me.ComboBox1.ListCount < 1 Then MsgBox "Please create a template group!": Exit Sub
    If ActiveSelectionRange.Count < 1 Then
        MsgBox "Please select a shape!"
        Exit Sub
    End If
    'TemplateManager.Enabled = False
    SaveTemplate.Caption = "Save Template"
    SaveTemplate.CommandButton1.Caption = "Save"
    SaveTemplate.Show
End Sub

Private Sub CommandButton3_Click()
    If Me.ListBox1.ListIndex < 0 Then MsgBox "Please select a template!": Exit Sub
    'TemplateManager.Enabled = False
    SaveTemplate.Caption = "Modify Template"
    SaveTemplate.CommandButton1.Caption = "Modify"
    SaveTemplate.Show
End Sub

Private Sub CommandButton4_Click()
    Dim i As Long
    If Me.ListBox1.ListIndex < 0 Then MsgBox "Please select a template!": Exit Sub
    i = MsgBox("Are you sure to delete the template?", vbYesNo)
    If i = 6 Then
        CGDKtemplates.Pages.Item(1).Layers.Item(Me.ComboBox1.value).FindShape(Me.ListBox1.List(Me.ListBox1.ListIndex)).Delete
        CGDKtemplates.Save
        TemplateManager.ListBox1.Clear
        For i = TemplateManager.CGDKtemplates.Pages.Item(1).Layers.Item(TemplateManager.ComboBox1.value).Shapes.Count To 1 Step -1
            TemplateManager.ListBox1.AddItem TemplateManager.CGDKtemplates.Pages.Item(1).Layers.Item(TemplateManager.ComboBox1.value).Shapes(i).Name
        Next i
        If TemplateManager.ListBox1.ListCount > 0 Then
            TemplateManager.ListBox1.ListIndex = 0
        Else
            Me.Label4.Caption = ""
        End If
    End If
End Sub

Private Sub CommandButton5_Click()
    Dim s As Shape
    If Me.ListBox1.ListIndex < 0 Then MsgBox "Please select a template": Exit Sub
    If Common.IsWaitingForAction Then Exit Sub
    Common.IsWaitingForAction = True
    Dim x As Double, y As Double, Shift As Long, b As Boolean

    CGDKtemplates.Pages.Item(1).Layers.Item(Me.ComboBox1.value).FindShape(Me.ListBox1.List(Me.ListBox1.ListIndex)).Copy
    b = ActiveDocument.GetUserClick(x, y, Shift, 5, False, cdrCursorSmallcrosshair)
    If Not b Then
    Set s = ActiveLayer.Paste
    s.SetPosition x, y
    End If
    Common.IsWaitingForAction = False
End Sub

Private Sub ListBox1_Change()
    If Me.ListBox1.ListCount > 0 Then
    On Error GoTo veryend:
    Me.Label4.Caption = CGDKtemplates.Pages.Item(1).Layers.Item(Me.ComboBox1.value).FindShape(Me.ListBox1.List(Me.ListBox1.ListIndex)).ObjectData("Description").value
    End If
    Exit Sub
veryend:
    Me.Label4.Caption = ""
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long, l As Layer, window1 As Window
    Set window1 = ActiveWindow
    Me.ComboBox2.AddItem "Add"
    Me.ComboBox2.AddItem "Modify"
    Me.ComboBox2.AddItem "Delete"
    Me.ComboBox2.value = "Add"
    On Error GoTo veryend1
    Set CGDKtemplates = Application.OpenDocument(Application.Path & "gms\CGDKtemplates.CDR")
    On Error GoTo 0
    window1.Activate
    For i = CGDKtemplates.Pages.Item(1).Layers.Count To 1 Step -1
        Me.ComboBox1.AddItem CGDKtemplates.Pages.Item(1).Layers(i).Name
    Next i
    If Me.ComboBox1.ListCount > 0 Then
        Me.ComboBox1.value = Me.ComboBox1.List(0)
    End If
    Exit Sub
veryend1:
    Dim d As Document
    Set d = Application.CreateDocument
    d.ActiveLayer.Name = "<default>"
    d.SaveAs Application.Path & "gms\CGDKtemplates.CDR"
    Resume 0
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    CGDKtemplates.Close
End Sub
