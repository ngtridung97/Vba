Attribute VB_Name = "OutlookControl"
Sub SendMessage(Optional AttachmentPath)
   
Dim ObjOutlook As Object
Dim ObjOutlookMsg As Object
Dim ObjOutlookRecip As Object
Dim ObjOutlookAttach As Object

Dim Address As String
Dim FilePath As String

Address = InputBox("Please input Address", "")
FilePath = Application.DefaultFilePath & "\" & ActiveWorkbook.Name

'Create Outlook session.
Set ObjOutlook = CreateObject("Outlook.Application")

'Create new message.
Set ObjOutlookMsg = ObjOutlook.CreateItem(0)

Application.DisplayAlerts = False

With ObjOutlookMsg

    'Add the To recipient(s) to the message.
    Set ObjOutlookRecip = .Recipients.Add(Address)

    'Set Subject, Body, or Importance of the message.
    .Subject = "Test Outlook Control" & " " & Now()
    
    'Add attachments to the message.
    Sheet1.Range("A1").Value = ObjOutlookMsg.Subject
    ActiveWorkbook.SaveAs FilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled  'Save and replace file in the Documents folder
    
    If Not IsMissing(FilePath) Then
        Set ObjOutlookAttach = .Attachments.Add(FilePath)
    End If

    'Resolve each Recipient's name.
    For Each ObjOutlookRecip In .Recipients
        ObjOutlookRecip.Resolve
        If Not ObjOutlookRecip.Resolve Then
        ObjOutlookMsg.Display
    End If
    
    Next
    
    .send
    
    Application.DisplayAlerts = True

End With

Set ObjOutlookMsg = Nothing
Set ObjOutlook = Nothing

End Sub

Sub ReplyMessage(Optional AttachmentPath)

Const OutlookFolderInbox = 6
Dim ObjOutlook As Object
Dim ObjOutlookMsg As Object
Dim ObjOutlookRecip As Object
Dim ObjOutlookAttach As Object
Dim ObjOutlookFolder As Object
Dim ObjOutlookNamespace As Object
Dim ObjOutlookReply As Object

Dim r As Long
Dim FilePath As String

FilePath = Application.DefaultFilePath & "\" & ActiveWorkbook.Name

'Create Outlook session.
Set ObjOutlook = CreateObject("Outlook.Application")
Set ObjOutlookNamespace = ObjOutlook.GetNamespace("MAPI")
Set ObjOutlookMsg = ObjOutlook.CreateItem(0)
Set ObjOutlookFolder = ObjOutlookNamespace.GetDefaultFolder(OutlookFolderInbox)

Application.DisplayAlerts = False

r = 1

For Each ObjOutlookMsg In ObjOutlookFolder.Items
    
    If InStr(ObjOutlookMsg.Subject, Sheet1.Range("A1").Value) <> 0 Then
    
    'Create reply message.
    Set ObjOutlookReply = ObjOutlookMsg.Reply
    
    'Add attachments to message.
    Sheet1.Range("A2").Value = "Receipt"
    ActiveWorkbook.SaveAs FilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    ObjOutlookReply.Attachments.Add (FilePath)
    
    'Reply
    ObjOutlookReply.Display
    ObjOutlookReply.send

    End If
    
r = r + 1

Next ObjOutlookMsg

Application.DisplayAlerts = True

Set ObjOutlookMsg = Nothing
Set ObjOutlook = Nothing

End Sub