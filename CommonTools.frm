VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CommonTools 
   Caption         =   "Common tools"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3285
   OleObjectBlob   =   "CommonTools.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "CommonTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Tools.SmartFill
End Sub

Private Sub CommandButton2_Click()
    Tools.SmartSelect
End Sub
