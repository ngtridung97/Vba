Attribute VB_Name = "RegEx"
Option Explicit

Sub Main()
    Dim input_str, pattern_str As String
    
    input_str = "Edit the Expression & Text"
    pattern_str = "([A-Z])\w+"
    Debug.Print ExtractRegEx(input_str, pattern_str)
End Sub

Function ExtractRegEx(ByVal input_str As String, ByVal partern_str As String) As String
    Dim RegEx As Object
    
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx
        .pattern = partern_str 'Change pattern to apply
        .Global = True
        If .Test(input_str) Then ExtractRegEx = CStr(.Execute(input_str)(0))
    End With
End Function
