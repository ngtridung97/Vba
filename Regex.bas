Attribute VB_Name = "RegEx"
Sub RegExString()

Dim Filename As String
Filename = InputBox("Please input Text", "")

Sheet1.Range("A1") = Extract(Filename)

End Sub

Function Extract(Filename As String) As String

    Dim RegEx As Object
    Dim Pattern As String
    
    Pattern = InputBox("Please input Expression", "")
    
    Set RegEx = CreateObject("VBScript.RegExp")
    
    With RegEx
        .Pattern = Pattern 'Change pattern to apply
        .Global = True
        If .Test(Filename) Then
            Extract = CStr(.Execute(Filename)(0))
        End If
    End With
    
End Function