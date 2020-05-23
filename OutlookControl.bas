Attribute VB_Name = "OutlookControl"
Sub SendMessage(Optional AttachmentPath)
   
Dim ObjOutlook As Object
Dim ObjOutlookMsg As Object
Dim ObjOutlookRecip As Object
Dim ObjOutlookAttach As Object

Dim FilePath As String

FilePath = Application.DefaultFilePath & "\" & ActiveWorkbook.Name

'Create Outlook session.
Set ObjOutlook = CreateObject("Outlook.Application")

'Create new message.
Set ObjOutlookMsg = ObjOutlook.CreateItem(0)

Application.DisplayAlerts = False

With ObjOutlookMsg

    'Add the To recipient(s) to the message.
    Set ObjOutlookRecip = .Recipients.Add("ng.tridung97@gmail.com")

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