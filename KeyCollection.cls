VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Collection1 As New Collection
Private Collection2 As New Collection

Public Function ElementNum() As Long
    ElementNum = Collection1.Count
End Function

Public Function AddElement(Element As Object, Key As String)
    Collection1.Add Element
    Collection2.Add Key
End Function

Public Function GetElement(Key As String) As Object
    On Error GoTo veryend:
    Dim i As Long
    For i = 1 To Collection1.Count
        If Key = Collection2.Item(i) Then Set GetElement = Collection1.Item(i): Exit Function
    Next i
veryend:
    Set GetElement = Nothing
End Function

Public Function RemoveElement(Key As String)
    Dim i As Long
    For i = 1 To Collection1.Count
        If Key = Collection2.Item(i) Then
            Collection1.Remove i
            Collection2.Remove i
            Exit Function
        End If
    Next i
End Function

Public Function RemoveAll()
    Dim i As Long
    For i = 1 To Collection1.Count
        Collection1.Remove 1
        Collection2.Remove 1
    Next i
End Function

Public Function KeyAt(Index As Long) As String
    KeyAt = Collection2.Item(Index)
End Function

Public Function ElementAt(Index As Long) As Object
    On Error GoTo veryend
    Set ElementAt = Collection1.Item(Index)
    Exit Function
veryend:
    ElementAt = Nothing
End Function
