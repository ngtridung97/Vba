Attribute VB_Name = "EventMenu"
Option Explicit

Sub AddSubmenu()
    Dim Bar As CommandBar
    Dim NewMenu As CommandBarControl
    Dim NewSubmenu As CommandBarButton
    
Call RemoveMenu
    
Set Bar = CommandBars("Cell")

'Add first menu
Set NewMenu = Bar.Controls.Add(Type:=msoControlPopup, temporary:=True)
    NewMenu.Caption = "My &Command 1"
    NewMenu.BeginGroup = True
    
    'Add first submenu item
    Set NewSubmenu = NewMenu.Controls.Add(Type:=msoControlButton)
        With NewSubmenu
            .FaceId = 542
            .Caption = "&Action 1"
            .OnAction = "Sub1"
        End With
        
    'Add second submenu item
    Set NewSubmenu = NewMenu.Controls.Add(Type:=msoControlButton)
        With NewSubmenu
            .FaceId = 535
            .Caption = "&Action 2"
            .OnAction = "Sub2"
        End With
        
    'Add third submenu item
    Set NewSubmenu = NewMenu.Controls.Add(Type:=msoControlButton)
        With NewSubmenu
            .FaceId = 489
            .Caption = "&Action 3"
            .OnAction = "Sub3"
        End With

'Add second menu
Set NewMenu = Bar.Controls.Add(Type:=msoControlPopup, temporary:=True)
    NewMenu.Caption = "My &Command 2"
    NewMenu.BeginGroup = True
    
    'Add fourth submenu item
    Set NewSubmenu = NewMenu.Controls.Add(Type:=msoControlButton)
        With NewSubmenu
            .FaceId = 422
            .Caption = "&Action 4"
            .OnAction = "Sub4"
        End With
        
    'Add fifth submenu item
    Set NewSubmenu = NewMenu.Controls.Add(Type:=msoControlButton)
        With NewSubmenu
            .FaceId = 514
            .Caption = "&Action 5"
            .OnAction = "Sub5"
        End With
    
End Sub

Sub RemoveMenu()

Dim cb As Office.CommandBar

For Each cb In CommandBars

    If Not cb.BuiltIn Then
        cb.Delete
    Else
        cb.Reset
    End If
    
Next

End Sub