VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SaveTemplate 
   Caption         =   "Save Template"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   OleObjectBlob   =   "SaveTemplate.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "SaveTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CommandButton1_Click()
    Dim s As Shape, l As Layer, d As Document, i As Long, CurrentIndex As Long
    If Me.CommandButton1.Caption = "Save" Then
        Set d = ActiveDocument
        If Me.TextBox1.value = "" Then MsgBox "Please input the template name!": Exit Sub
        If TemplateManager.IsExisted(Me.TextBox1.value, "shape") Then MsgBox "The name has been used! Please change a name.": Exit Sub
        Set s = ActiveSelectionRange.group
        s.Name = Me.TextBox1.Text
        If Not d.DataFields.IsPresent("description") Then d.DataFields.AddEx "Description", cdrDataTypeString, "", , , "0", False, False, False
        s.ObjectData("Description").value = Me.TextBox2.Text
        s.Copy
        TemplateManager.CGDKtemplates.Pages.Item(1).Layers.Item(TemplateManager.ComboBox1.value).Paste
        TemplateManager.CGDKtemplates.Save
        TemplateManager.ListBox1.Clear
        For i = TemplateManager.CGDKtemplates.Pages.Item(1).Layers.Item(TemplateManager.ComboBox1.value).Shapes.Count To 1 Step -1
            TemplateManager.ListBox1.AddItem TemplateManager.CGDKtemplates.Pages.Item(1).Layers.Item(TemplateManager.ComboBox1.value).Shapes(i).Name
        Next i
        If TemplateManager.ListBox1.ListCount > 0 Then
            TemplateManager.ListBox1.ListIndex = TemplateManager.ListBox1.ListCount - 1
        End If
        Unload Me
    Else
        Set d = ActiveDocument
        CurrentIndex = TemplateManager.ListBox1.ListIndex
        Set s = TemplateManager.CGDKtemplates.Pages.Item(1).Layers.Item(TemplateManager.ComboBox1.value).Shapes.Item(TemplateManager.ListBox1.List(CurrentIndex))
        If Me.TextBox1.value = "" Then MsgBox "Please input the template name!": Exit Sub
        'If TemplateManager.IsExisted(Me.TextBox1.Value, "shape") Then MsgBox "The name has been used! Please change a name.": Exit Sub
        s.Name = Me.TextBox1.value
        If Not d.DataFields.IsPresent("description") Then d.DataFields.AddEx "Description", cdrDataTypeString, "", , , "0", False, False, False
        s.ObjectData("Description").value = Me.TextBox2.value
        TemplateManager.CGDKtemplates.Save
        TemplateManager.ListBox1.Clear
        For i = TemplateManager.CGDKtemplates.Pages.Item(1).Layers.Item(TemplateManager.ComboBox1.value).Shapes.Count To 1 Step -1
            TemplateManager.ListBox1.AddItem TemplateManager.CGDKtemplates.Pages.Item(1).Layers.Item(TemplateManager.ComboBox1.value).Shapes(i).Name
        Next i
        If TemplateManager.ListBox1.ListCount > 0 Then
            TemplateManager.ListBox1.ListIndex = CurrentIndex
        End If
        Unload Me
    End If
End Sub

Private Sub UserForm_Activate()
    If Me.Caption = "Modify Template" Then
        Me.TextBox1.value = TemplateManager.ListBox1.List(TemplateManager.ListBox1.ListIndex)
        Me.TextBox2.value = TemplateManager.CGDKtemplates.Pages.Item(1).Layers.Item(TemplateManager.ComboBox1.value).FindShape(Me.TextBox1.value).ObjectData("Description").value
    End If
End Sub

Private Sub UserForm_Terminate()
    Me.Caption = ""
    Me.CommandButton1.Caption = ""
End Sub
