Attribute VB_Name = "TransposeArray"
Option Explicit

Function TransArr(input_arr As Variant) As Variant
    Dim i, j As Long
    Dim output_arr As Variant
    
    'Create New Array
    ReDim output_arr(LBound(input_arr, 1) To UBound(input_arr, 2), LBound(input_arr, 2) To UBound(input_arr, 1))
    
    'Transpose Array
    For i = LBound(input_arr, 1) To UBound(input_arr, 1)
        For j = LBound(input_arr, 2) To UBound(input_arr, 2)
            output_arr(j, i) = input_arr(i, j)
        Next j
    Next i
    
    'Output Array
    TransposeArray = output_arr
End Function
