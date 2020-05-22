Attribute VB_Name = "VisibleCell"
Sub SelectPreviousVisibleCell1()

Dim Range As Range
Set Range = ActiveCell

Do
    Set Range = Range.Offset(-1, 0)
Loop Until Range.EntireRow.Hidden = False

Range.Select

End Sub

Sub CheckPreviousVisibleCell1()

If ActiveCell.Value = PreviousVisibleCell(ActiveCell) Then
    MsgBox ("Equal!")
Else
    MsgBox ("Not Equal!")
End If

End Sub

Private Function PreviousVisibleCell(Range As Range) As String

Do
    Set Range = Range.Offset(-1, 0)
Loop Until (Range.EntireRow.Hidden = False)

PreviousVisibleCell = Range

End Function