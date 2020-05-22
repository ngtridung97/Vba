Attribute VB_Name = "LoopFolder"
Sub LoopFolder()

Dim Directory As String

With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Select a location containing the files you want to list."
    .Show
    
    If .SelectedItems.Count = 0 Then
        Exit Sub
    Else
        Directory = .SelectedItems(1) & "\"
    End If
    
End With

End Sub