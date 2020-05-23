Attribute VB_Name = "SAPControl"
Sub SimpleSAPExport()

Dim SapGuiAuto As Object
Dim ObjGui  As Object
Dim ObjConn As Object
Dim ObjSess As Object
Dim SavePath As String

Set SapGuiAuto = GetObject("SAPGUI") 'Get the SAP GUI Scripting object
Set ObjGui = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
Set ObjConn = ObjGui.Children(0) 'Get the first system that is currently connected
Set ObjSess = ObjConn.Children(0) 'Get the first session (window) on that connection

'Create folder picker    
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Select a location containing the files you want to zip."
    .Show
    
    If .SelectedItems.Count = 0 Then
        Exit Sub
    Else
        Directory = .SelectedItems(1) & "\"
    End If
    
End With

SavePath = Directory

For r = 2 To 10000

If Sheet1.Range("A" & r) = "" Then Exit For

With ObjSess

    .findById("wnd[0]").maximize
    .StartTransaction "FB03" 'Load the transaction you are after
    .findById("wnd[0]/usr/txtRF05L-BELNR").Text = Sheet1.Range("C" & r) 'Input Document Number into SAP
    .findById("wnd[0]/usr/ctxtRF05L-BUKRS").Text = Sheet1.Range("D" & r) 'Input Company Code into SAP
    .findById("wnd[0]/usr/txtRF05L-GJAHR").Text = Sheet1.Range("E" & r) 'Input Fiscal Year into SAP
    .findById("wnd[0]").sendVKey 0 'Execute transaction
    
    'The query runs and you select context menu and attachments
    .findById("wnd[0]/titl/shellcont/shell").pressContextButton "%GOS_TOOLBOX"
    .findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem "%GOS_VIEW_ATTA"
    
    .findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").currentCellColumn = "BITM_DESCR"
    
    'In case no hard copies
    On Error GoTo ErrCol
    
    .findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "0"
    .findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").contextMenu
    .findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectContextMenuItem "%ATTA_EXPORT"
    .findById("wnd[2]/usr/ctxtDY_FILENAME").Text = Sheet1.Range("A" & r) & "_" & Range("B" & r) & ".tif" 'Fix name into Supplier Number + Invoice Number
    .findById("wnd[2]/usr/ctxtDY_PATH").Text = SavePath 'Fix path for saving file
    
    'Close all current windows
    .findById("wnd[2]/tbar[0]/btn[0]").press
    .findById("wnd[1]").Close
    .findById("wnd[0]/tbar[0]/btn[3]").press
    
    End With
    
NextCol:
    Next r
    
Sheet1.Select
Range("A1").Select
    
MsgBox "COPY COMPLETED!"
    
Exit Sub

ErrCol:
    Sheets("Result").Range("F" & r).Value = "Document Not Copied" 'Note invoice copy status
    Resume NextCol

'Clean up
Set SapGuiAuto = Nothing
Set ObjGui = Nothing
Set ObjConn = Nothing
Set ObjSess = Nothing
    
End Sub