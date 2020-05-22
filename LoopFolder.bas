Attribute VB_Name = "LoopFolder"
Sub LoopFolder()

Dim Directory As String

'Add folder picker
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

Sub RecursiveDir(ByVal CurrDir As String)

Dim Dirs() As String
Dim NumDirs As Long
Dim Filename As String
Dim PathAndName As String
Dim Filesize As Double

'Put column headings on active sheet
Sheet1.Select
Cells(1, 1) = "Path"
Cells(1, 2) = "Filename"
Cells(1, 3) = "FullPath"
Cells(1, 4) = "Size"
Cells(1, 5) = "Date/Time"
Range("A1:G1").Font.Bold = True
    
'Make sure path ends in backslash
If Right(CurrDir, 1) <> "\" Then CurrDir = CurrDir & "\"
    
'Get files
Filename = Dir(CurrDir & "*.*", vbDirectory)
Do While Len(Filename) <> 0
    
    If Left(Filename, 1) <> "." Then 'Current dir
        PathAndName = CurrDir & Filename
        
        If (GetAttr(PathAndName) And vbDirectory) = vbDirectory Then
            'Store found directories
            ReDim Preserve Dirs(0 To NumDirs) As String
            Dirs(NumDirs) = PathAndName
            NumDirs = NumDirs + 1
           
        Else
            'Write the path and file to the sheet
            Cells(WorksheetFunction.CountA(Range("A:A")) + 1, 1) = Left(CurrDir, Len(CurrDir) - 1)
            Cells(WorksheetFunction.CountA(Range("B:B")) + 1, 2) = Filename
            Cells(WorksheetFunction.CountA(Range("C:C")) + 1, 3) = CurrDir & "" & Filename
        
            'Adjust for filesize > 2 gigabytes
            Filesize = FileLen(PathAndName)
            If Filesize < 0 Then
                Filesize = Filesize + 4294967296#
            End If
            
            Cells(WorksheetFunction.CountA(Range("D:D")) + 1, 4) = Filesize
            Cells(WorksheetFunction.CountA(Range("E:E")) + 1, 5) = FileDateTime(PathAndName)
        
        End If
        
    End If
    
Filename = Dir()
Loop
    
End Sub