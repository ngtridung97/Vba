VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DragDrop 
   Caption         =   "Drag and Drop"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4485
   OleObjectBlob   =   "DragDrop.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DragDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TreeView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim StrPath As String
    StrPath = Data.Files(1)
    'Debug.Print (StrPath)
    Call EmailCheck(StrPath)
    Call MakeFolder(StrPath)
    UserForm_Initialize
    MsgBox ("File Moved to Evidences")
    
End Sub

Private Sub UserForm_Initialize()
    
End Sub

Private Sub UserForm_Activate()

    Me.StartUpPosition = 0
    Me.Top = Application.Top + 25
    Me.Left = Application.Left + Application.Width - Me.Width - 25

End Sub

