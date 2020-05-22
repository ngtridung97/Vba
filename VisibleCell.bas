Attribute VB_Name = "VisibleCell"
Sub SelectPreviousVisibleCell1()

Dim Range As Range
Set Range = ActiveCell

Do
    Set Range = Range.Offset(-1, 0)
Loop Until Range.EntireRow.Hidden = False

Range.Select

End Sub