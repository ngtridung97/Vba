Attribute VB_Name = "NumberArray"
Option Explicit

Function NumbArr(ByVal input_arr As Variant)
    Dim dict As Scripting.Dictionary
    Dim i As Long
    
    Set dict = New Scripting.Dictionary
    dict.CompareMode = BinaryCompare
    
    For i = LBound(input_arr) To UBound(input_arr)
        If Not dict.Exists(input_arr(i)) Then
            dict.Add input_arr(i), 0
        End If
    Next i
    
    For i = LBound(input_arr) To UBound(input_arr)
        dict(input_arr(i)) = dict(input_arr(i)) + 1
        input_arr(i) = input_arr(i) & dict(input_arr(i))
    Next i
    
    NumbArr = input_arr
End Function
