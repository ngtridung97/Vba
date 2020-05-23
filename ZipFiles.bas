Attribute VB_Name = "ZipFiles"
Sub ZipFiles()

Dim Fso As Object, ZipFile As Object, ObjShell As Object
Dim FsoFolder As Object, FsoFile As Object
Dim TimerStart As Single
Dim FolderPath As String, ZipName As String

'Folder Picker
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Select a location containing the files you want to list."
    .Show
    
    If .SelectedItems.Count = 0 Then
        Exit Sub
    Else
        Directory = .SelectedItems(1) & "\"
    End If
    
End With

'Folder to zip and Zip file name
FolderPath = Directory
ZipName = "ZipFile.zip"

'Create FSO to loop
Set Fso = CreateObject("Scripting.FileSystemObject")

'Create Zip file
Set ZipFile = Fso.CreateTextFile(FolderPath & ZipName)
ZipFile.WriteLine Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
ZipFile.Close

Set ObjShell = CreateObject("Shell.Application")
Set FsoFolder = Fso.GetFolder(FolderPath)

'Start a For Loop
For Each FsoFile In FsoFolder.Files

    Debug.Print FsoFile.Name
    If FsoFile.Name <> ZipName Then ' Check it's not the zip file before adding

        ObjShell.Namespace("" & FolderPath & ZipName).CopyHere FsoFile.Path

        TimerStart = Timer
        Do While Timer < TimerStart + 5
            Application.StatusBar = "Zipping, please wait..."
            DoEvents
        Loop

    End If

Next

'Clean up
Application.StatusBar = ""
Set FsoFile = Nothing
Set FsoFolder = Nothing
Set ObjShell = Nothing
Set ZipFile = Nothing
Set Fso = Nothing

MsgBox "Zipped Completed!", vbInformation

End Sub