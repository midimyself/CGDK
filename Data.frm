VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Data 
   Caption         =   "CGDKData"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1785
   OleObjectBlob   =   "Data.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub OK_Click()
    GetDataRange.IsOK = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
Dim formhandle As Long
formhandle = API.FindWindow(vbNullString, "CGDKData")
API.SetWindowPos formhandle, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub

Private Sub UserForm_Terminate()
    GetDataRange.Waiting = False
End Sub
