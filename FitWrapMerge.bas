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