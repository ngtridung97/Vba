Attribute VB_Name = "FileArrangement"
Sub MakeFolder(EmailPath As String)
Attribute MakeFolder.VB_ProcData.VB_Invoke_Func = "D\n14"

'Variables for Loop
Dim i As Long
Dim a, b As Range

'Variables for Window Explorer
Dim StrPath, EmailName As String
Dim Fso As Object
Dim SourceFileName As String, DestFileName As String

Set a = Selection

If a.Count = 1 Then

    i = ActiveCell.Row()
    
    'Get Email FileName
    EmailName = GetFilenameFromPath(EmailPath)
    
    StrPath = "D:\EmailPath\" & i & "\"
    
    If Not FolderExists(StrPath) Then FolderCreate (StrPath)
    
    Set Fso = CreateObject("Scripting.Filesystemobject")
    
    SourceFileName = EmailPath
    DestFileName = StrPath & "\" & EmailName
    
    Fso.CopyFile Source:=SourceFileName, Destination:=DestFileName

Else

    Set a = Selection.SpecialCells(xlCellTypeVisible)
    
    For Each b In a.Rows
    
        i = b.Row
        
        'Get Email FileName
        EmailName = GetFilenameFromPath(EmailPath)
        
        StrPath = "D:\EmailPath\" & i & "\"
        
        If Not FolderExists(StrPath) Then FolderCreate (StrPath)
        
        Set Fso = CreateObject("Scripting.Filesystemobject")
        
        SourceFileName = EmailPath
        DestFileName = StrPath & "\" & EmailName
        
        Fso.CopyFile Source:=SourceFileName, Destination:=DestFileName
    
    Next b

End If

End Sub

Function FolderCreate(ByVal Path As String) As Boolean

FolderCreate = True
Dim Fso As New FileSystemObject
Dim StrPath As String
Dim lCtr As Long

StrPath = Path

arrpath = Split(StrPath, "\")
StrPath = arrpath(LBound(arrpath)) & "\"

For lCtr = LBound(arrpath) + 1 To UBound(arrpath)
    StrPath = StrPath & arrpath(lCtr) & "\"
    If Dir(StrPath, vbDirectory) = "" Then
        MkDir StrPath
    End If
Next

End Function

Function FolderExists(ByVal Path As String) As Boolean

FolderExists = False
Dim Fso As New FileSystemObject

If Fso.FolderExists(Path) Then FolderExists = True

End Function

Function CleanName(strName As String) As String
'Will clean part # name so it can be made into valid folder name
'May need to add more lines to get rid of other characters

CleanName = Replace(strName, "/", "")
CleanName = Replace(CleanName, "*", "")
'etc...

End Function

Function GetFilenameFromPath(ByVal StrPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
' e.g. 'c:\winnt\win.ini' returns 'win.ini'

If Right$(StrPath, 1) <> "\" And Len(StrPath) > 0 Then
    GetFilenameFromPath = GetFilenameFromPath(Left$(StrPath, Len(StrPath) - 1)) + Right$(StrPath, 1)
End If

End Function
