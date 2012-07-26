Attribute VB_Name = "File"
Option Explicit

Private Type OPENFILENAME
          lStructSize   As Long
          hwndOwner   As Long
          hInstance   As Long
          lpstrFilter   As String
          lpstrCustomFilter   As String
          nMaxCustFilter   As Long
          nFilterIndex   As Long
          lpstrFile   As String
          nMaxFile   As Long
          lpstrFileTitle   As String
          nMaxFileTitle   As Long
          lpstrInitialDir   As String
          lpstrTitle   As String
          flags   As Long
          nFileOffset   As Integer
          nFileExtension   As Integer
          lpstrDefExt   As String
          lCustData   As Long
          lpfnHook   As Long
          lpTemplateName   As String
End Type

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Dim OFName As OPENFILENAME
Dim sOpenLastPath As String
Dim sSaveLastPath As String

Function OpenExcelFile() As Integer
    Dim FileName As String
    FileName = File.ShowOpen
    Dim ExcelSheet As Object, Excel As Object
    Set ExcelSheet = CreateObject("Excel.Sheet")
    ExcelSheet.Parent.Windows(1).WindowState = -4137
    Set Excel = ExcelSheet.Application
    Set ExcelSheet = Nothing
    If FileName <> "" Then
        Excel.WorkBooks.Open FileName
        Excel.Visible = True
        OpenExcelFile = 1
    Else
        OpenExcelFile = 0
    End If
End Function

Function ShowOpen(Optional sFilter As String = "", Optional sDir As String = "LastPath", Optional sTitle As String = "Open Excel File") As String
          Dim x() As String
          OFName.lStructSize = Len(OFName)
          OFName.hwndOwner = 0
          OFName.hInstance = 0
          OFName.lpstrDefExt = ""

          If sFilter = "" Then sFilter = "Excel 2003 (*.xls)|*.xls|Excel 2007 (*.xlsx)|*.xlsx"
          x = VBA.Split(sFilter, "|"): sFilter = Join(x, VBA.Chr(0))
          OFName.lpstrFilter = sFilter
          OFName.lpstrFile = VBA.Space$(254)
          OFName.nMaxFile = 255
          OFName.lpstrFileTitle = VBA.Space$(254)
          OFName.nMaxFileTitle = 255

          If sDir = "LastPath" Then sDir = IIf(VBA.Len(sOpenLastPath) = 0, VBA.CreateObject("Shell.Application").NameSpace(5).Self.Path, sOpenLastPath)
          OFName.lpstrInitialDir = sDir

          OFName.lpstrTitle = sTitle
          OFName.flags = 0
    
          If GetOpenFileName(OFName) Then
                ShowOpen = VBA.Trim$(Replace(OFName.lpstrFile, vbNullChar, ""))
          Else
                ShowOpen = ""
          End If
          sOpenLastPath = VBA.left(ShowOpen, InStrRev(ShowOpen, "\"))
End Function
