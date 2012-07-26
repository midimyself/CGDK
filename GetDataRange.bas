Attribute VB_Name = "GetDataRange"
Option Explicit

Public Waiting As Boolean
Public IsOK As Boolean

Function GetDataRange(Optional DefaultRange) As Object
    Waiting = True
    
    On Error GoTo veryend:
    Dim ExcelSheet As Object, Excel As Object
    Set ExcelSheet = CreateObject("Excel.Sheet")
    ExcelSheet.Parent.Windows(1).WindowState = -4137
    Set Excel = ExcelSheet.Application
    Set ExcelSheet = Nothing
    
    If Excel.WorkBooks.Count < 1 Then
        Dim q As Integer
        q = MsgBox("No Excel Workbook is available! Do you want to open one now?", vbYesNo)
        If q = 6 Then
            If File.OpenExcelFile = 0 Then Exit Function
        Else
            Exit Function
        End If
    End If
    
    data.Show 0
    
    If Not IsMissing(DefaultRange) Then
        DefaultRange.Parent.Parent.Activate
        DefaultRange.Parent.Activate
        Excel.Range(DefaultRange.Address).Select
    End If
    
    AppActivate Excel.Caption
    Excel.ActiveWindow.WindowState = -4137
        
    Do While Waiting
        waitsec 0.01
    Loop
    
    If IsOK Then
        Set GetDataRange = Excel.Selection
        IsOK = False
    Else
        Set GetDataRange = Nothing
    End If
    
    AppActivate Application.AppWindow.Caption

Exit Function

veryend:

Waiting = False
Set GetDataRange = Nothing

AppActivate Application.AppWindow.Caption

If Err.Number = 424 Then
ElseIf Err.Number = -2147417846 Then
    MsgBox "Can not access the Excel workbook! Maybe another process is using the software."
Else
    MsgBox Err.Description
End If
End Function
Private Sub waitsec(ByVal sec As Double)
Dim sTimer As Date
sTimer = Timer
Do
DoEvents
Loop While VBA.Format((Timer - sTimer), "0.000") < sec
End Sub
