VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private data() As Double
Public mean As Double

Public Sub Add(ByVal a As Double)
    ReDim Preserve data(1 To GetArrayElementNumber(data) + 1)
    data(GetArrayElementNumber(data)) = a
End Sub

Private Function GetArrayElementNumber(a() As Double) As Long
    On Error GoTo veryend
    GetArrayElementNumber = UBound(a) - LBound(a) + 1
    Exit Function
veryend:
    GetArrayElementNumber = 0
End Function

Public Function GetItem(ByVal Index As Long) As Double
    GetItem = data(Index)
End Function

Public Sub SetItem(ByVal Index As Long, ByVal value As Double)
    data(Index) = value
End Sub

Public Function DataNumInRange(ByVal lower As Double, ByVal upper As Double) As Double
    Dim n As Long, i As Long
    For i = 1 To GetArrayElementNumber(data)
        If data(i) >= lower And data(i) < upper Then n = n + 1
    Next i
    DataNumInRange = n
End Function

Public Function LinkedGroupDataMeanInRange(group As DataGroup, ByVal lower As Double, ByVal upper As Double) As Double
    Dim i As Long, sum As Double, n As Long
    n = DataNumInRange(lower, upper)
    If n > 0 Then
        For i = 1 To GetArrayElementNumber(data)
            If data(i) >= lower And data(i) < upper Then sum = sum + group.GetItem(i)
        Next i
        LinkedGroupDataMeanInRange = sum / n
    Else
        LinkedGroupDataMeanInRange = 0
    End If
End Function

Public Function DataSumInRange(ByVal lower As Double, ByVal upper As Double) As Double
    Dim sum As Double, i As Long
    For i = 1 To GetArrayElementNumber(data)
        If data(i) >= lower And data(i) < upper Then sum = data(i) + sum
    Next i
    DataSumInRange = sum
End Function

Public Function DataMeanInRange(ByVal lower As Double, ByVal upper As Double) As Double
    DataMeanInRange = DataSumInRange(lower, upper) / DataNumInRange(lower, upper)
    Exit Function
End Function

Public Function DataNum() As Long
    DataNum = GetArrayElementNumber(data)
End Function

