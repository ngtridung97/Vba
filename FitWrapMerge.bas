Attribute VB_Name = "FitWrapMerge"
Sub AutoFitMergedCellRowHeight()

Dim CurrentRowHeight As Single, MergedCellRgWidth As Single
Dim ActiveCellWidth As Single, PossNewRowHeight As Single

If ActiveCell.MergeCells Then

    With ActiveCell.MergeArea
    
        If .Rows.Count = 1 And .WrapText = True Then
            
            CurrentRowHeight = .RowHeight
            ActiveCellWidth = ActiveCell.ColumnWidth
            .MergeCells = False
            .Cells(1).ColumnWidth = MergedCellRgWidth
            .EntireRow.AutoFit
            PossNewRowHeight = .RowHeight
            .Cells(1).ColumnWidth = ActiveCellWidth
            .MergeCells = True
            .RowHeight = IIf(CurrentRowHeight > PossNewRowHeight, CurrentRowHeight, PossNewRowHeight)
            
        End If
        
    End With
    
End If

End Sub

Sub AutoFitAll() 'Update autofit all, also for columns

Dim Range As Range

For Each Range In Selection
    MergeCellsFit Range
Next

End Sub

Sub MergeCellsFit(ByVal MergeCells As Range)

Dim Diff As Single
Dim FirstCell As Range, MergeCellArea As Range
Dim Col As Long, ColCount As Long, RowCount As Long
Dim FirstCellWidth As Double, FirstCellHeight As Double, MergeCellWidth As Double

If MergeCells.Count = 1 Then
    Set MergeCellArea = MergeCells.MergeArea
Else
    Set MergeCellArea = MergeCells
End If

With MergeCellArea
    ColCount = .Columns.Count
    RowCount = .Rows.Count
    .WrapText = True
    
    'Check if merge in only 1 cell, fit like normal
    If RowCount = 1 And ColCount = 1 Then
        .EntireRow.AutoFit
        
    Else
      
        Set FirstCell = .Cells(1, 1)
        FirstCellWidth = FirstCell.ColumnWidth
        Diff = 0.75
        
        For Col = 1 To ColCount
            MergeCellWidth = MergeCellWidth + .Cells(1, Col).ColumnWidth + Diff
        Next
        
        'Unmerge and adjust range
        .MergeCells = False
        FirstCell.ColumnWidth = MergeCellWidth - Diff
        .EntireRow.AutoFit
        FirstCellHeight = FirstCell.RowHeight
        .MergeCells = True
        FirstCell.ColumnWidth = FirstCellWidth
        FirstCellHeight = FirstCellHeight / RowCount
        .RowHeight = FirstCellHeight
            
    End If

End With

End Sub