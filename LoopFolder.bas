Attribute VB_Name = "LoopFolder"
Option Explicit

Dim Dirs() As String
Dim NumDirs As Long
Dim Filename As String
Dim PathAndName As String
Dim i As Long
Dim Filesize As Double
Dim Directory As String

Sub Main()
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

    ActiveSheet.Cells.ClearContents
    Call RecursiveDir(Directory)
End Sub

Sub RecursiveDir(ByVal CurrDir As String)
    'Heading
    With ActiveSheet.Range("A1:E1")
        .Value2 = Array("Path", "Filename", "FullPath", "Size", "Date/Time")
        .Font.Bold = True
    End With
        
    'Make sure path ends in backslash
    If Right(CurrDir, 1) <> "\" Then CurrDir = CurrDir & "\"
        
    'Get files
    Filename = Dir(CurrDir & "*.*", vbDirectory)
    i = 2
    Do While Len(Filename) <> 0
        If Left(Filename, 1) <> "." Then 'Current dir
            PathAndName = CurrDir & Filename
            
            If (GetAttr(PathAndName) And vbDirectory) = vbDirectory Then
                'Store found directories
                ReDim Preserve Dirs(0 To NumDirs) As String
                Dirs(NumDirs) = PathAndName
                NumDirs = NumDirs + 1
            Else
                'Adjust for filesize > 2 gigabytes
                Filesize = FileLen(PathAndName)
                If Filesize < 0 Then
                    Filesize = Filesize + 4294967296#
                End If
                
                'Write to sheet
                With ActiveSheet
                    .Cells(i, 1) = Left(CurrDir, Len(CurrDir) - 1)
                    .Cells(i, 2) = Filename
                    .Cells(i, 3) = CurrDir & "" & Filename
                    .Cells(i, 4) = Filesize
                    .Cells(i, 5) = FileDateTime(PathAndName)
                End With
                i = i + 1
            End If
        End If
        
        Filename = Dir()
    Loop
        
    'Process the found directories recursively
    For i = 0 To NumDirs - 1
        RecursiveDir Dirs(i)
    Next i
End Sub
