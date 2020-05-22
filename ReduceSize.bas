Attribute VB_Name = "ReduceSize"
Sub ReduceSize()
     
Dim ws As Worksheet
Dim LastRow As Long
Dim LastCol As Long
    
For Each ws In ActiveWorkbook.Worksheets
    
    With ws
        
        LastRow = .Cells.Find(What:="*", After:=.Range("A1"), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        LastCol = .Cells.Find(What:="*", After:=.Range("A1"), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        Range(.Cells(1, LastCol + 1), .Cells(.Rows.Count, .Columns.Count)).Delete
        Range(.Cells(LastRow + 1, 1), .Cells(.Rows.Count, .Columns.Count)).Delete
        LastRow = .UsedRange.Rows.Count
            
    End With
        
Next ws
     
End Sub