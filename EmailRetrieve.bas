Attribute VB_Name = "EmailRetrieve"
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2

Sub EmailCheck(EmailPath As String)

'Variables for Loop
Dim i As Long
Dim a, b As Range

'Variables for Emails
Dim Myinspect, MyItem, MyAtt, OL As Object
Dim Subject, Sender, SendDate, AttCount As String
Dim Notes, CurrentNotes, UpdateNotes As String
    
If Dir(EmailPath) = "" Then
    MsgBox "File " & EmailPath & " does not exist"
Else
    ShellExecute 0, "Open", EmailPath, "", EmailPath, SW_SHOWNORMAL
End If
    
Application.Wait (Now + TimeValue("0:00:03"))
        
Set OL = CreateObject("Outlook.Application")
Set Myinspect = OL.ActiveInspector
Set MyItem = Myinspect.CurrentItem
Set MyAtt = MyItem.Attachments
    
Subject = MyItem.Subject
Sender = MyItem.SenderEmailAddress
SendDate = MyItem.SentOn
AttCount = MyAtt.Count
    
Set a = Selection
    
If a.Count = 1 Then
    
    i = ActiveCell.Row
    
    CurrentNotes = ActiveWorkbook.ActiveSheet.Range("A" & i).Text
        
    Notes = "Email on: " & SendDate & Chr(10) & "From: " & Sender & Chr(10) & "Subject: " & Subject & Chr(10) & "Attachments: " & AttCount
            
    If CurrentNotes = "" Then
        UpdatedNotes = Notes
    Else
        UpdatedNotes = Notes & Chr(10) & Chr(10) & CurrentNotes
    End If
            
    ActiveWorkbook.ActiveSheet.Range("A" & i).Value = UpdatedNotes
           
Else
    
    Set a = Selection.SpecialCells(xlCellTypeVisible)
    
    For Each b In a.Rows

        i = b.Row
        
        CurrentNotes = ActiveWorkbook.ActiveSheet.Range("A" & i).Text
        
        Notes = "Email on: " & SendDate & Chr(10) & "From: " & Sender & Chr(10) & "Subject: " & Subject & Chr(10) & "Attachments: " & AttCount
                
        If CurrentNotes = "" Then
            UpdatedNotes = Notes
        Else
            UpdatedNotes = Notes & Chr(10) & Chr(10) & CurrentNotes
        End If
                
        ActiveWorkbook.ActiveSheet.Range("A" & i).Value = UpdatedNotes
    
    Next b

End If
    
End Sub
