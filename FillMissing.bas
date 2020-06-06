Attribute VB_Name = "FillMissing"
Sub FillData()

Dim Rng As Range
Dim NextRng As Range
Dim i, j As Long

Call SampleData

'Get last row
i = ActiveWorkbook.ActiveSheet.Range("A" & ActiveWorkbook.ActiveSheet.Rows.Count).End(xlUp).Row

'Sort data
ActiveWorkbook.ActiveSheet.Range("A1", "B" & i).AutoFilter
ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Clear
ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range("A2", "A" & i), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Add Key:=Range("B2", "B" & i), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.ActiveSheet.AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
ActiveWorkbook.ActiveSheet.AutoFilter.Sort.SortFields.Clear

'Filter blank value
ActiveWorkbook.ActiveSheet.Range("A1", "B" & i).AutoFilter Field:=2, Criteria1:="="
On Error GoTo Cancel

    'Fill item_idnt
    Set Rng = ActiveWorkbook.ActiveSheet.Range("B2", "B" & i).SpecialCells(xlCellTypeVisible)
    Set NextRng = Range("B1")
    
    Do
        Set NextRng = NextRng.Offset(1, 0)
    Loop Until NextRng.EntireRow.Hidden = False

    NextRng.Select
    j = Selection.Row
    Selection.Formula = "=" & "IF(" & "A" & j & "=" & "A" & j - 1 & "," & "B" & j - 1 & "," & "IF(" & "A" & j & "=" & "A" & j + 1 & "," & "B" & j + 1 & "))"
    Rng.SpecialCells(xlCellTypeVisible).FillDown

Cancel:

'Remove filter and paste to value
ActiveWorkbook.ActiveSheet.Range("A1", "B" & i).AutoFilter
ActiveWorkbook.ActiveSheet.Range("A1", "B" & i).Copy
ActiveWorkbook.ActiveSheet.Range("A1", "B" & i).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
ActiveWorkbook.ActiveSheet.Range("A1").Select

End Sub

Sub SampleData()

ActiveWorkbook.ActiveSheet.Cells.ClearContents

'Header
ActiveWorkbook.ActiveSheet.Range("A1").Value = "ID"
ActiveWorkbook.ActiveSheet.Range("B1").Value = "Item"

'Detail
ActiveWorkbook.ActiveSheet.Range("A2").Value = "1"
ActiveWorkbook.ActiveSheet.Range("A3").Value = "2"
ActiveWorkbook.ActiveSheet.Range("A4").Value = "3"
ActiveWorkbook.ActiveSheet.Range("A5").Value = "1"
ActiveWorkbook.ActiveSheet.Range("A6").Value = "2"
ActiveWorkbook.ActiveSheet.Range("A7").Value = "3"
ActiveWorkbook.ActiveSheet.Range("A8").Value = "1"
ActiveWorkbook.ActiveSheet.Range("A9").Value = "2"
ActiveWorkbook.ActiveSheet.Range("A10").Value = "4"
ActiveWorkbook.ActiveSheet.Range("A11").Value = "4"

ActiveWorkbook.ActiveSheet.Range("B2").Value = "A"
ActiveWorkbook.ActiveSheet.Range("B3").Value = "B"
ActiveWorkbook.ActiveSheet.Range("B4").Value = "C"
ActiveWorkbook.ActiveSheet.Range("B10").Value = "D"

End Sub
