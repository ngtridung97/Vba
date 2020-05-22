Attribute VB_Name = "RemoveAllMacros"
Sub RemoveAllMacros()

Dim otmp As Object

    With ActiveWorkbook.VBProject
    
        For Each otmp In .VBComponents
        
            If otmp.Type = 100 Then
                otmp.CodeModule.DeleteLines 1, otmp.CodeModule.CountOfLines
                otmp.CodeModule.CodePane.Window.Close
            Else: .VBComponents.Remove otmp
            End If
    
        Next otmp
        
    End With

End Sub