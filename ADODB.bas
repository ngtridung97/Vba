Attribute VB_Name = "ADODB"
'PostgesSQL Connection
Const postgreSQL_Connection As String = "Driver={PostgreSQL ANSI};Server=?;Port=?;UID=?;PWD=?; Database=postgres;READONLY=0;PROTOCOL=6.4;FAKEOIDINDEX=0;SHOWOIDCOLUMN=0;ROWVERSIONING=0;SHOWSYSTEMTABLES=1"

Sub ADODB()

Dim filterArray()
Dim CurrentFiltRange As String
Dim Col As Integer

Application.ScreenUpdating = False

'Capture current filter
With ActiveSheet.AutoFilter
    CurrentFiltRange = .Range.Address
    With .Filters
        ReDim filterArray(1 To .Count, 1 To 3)
        For f = 1 To .Count
            With .Item(f)
                If .On Then
                    filterArray(f, 1) = .Criteria1
                    If .Operator Then
                        filterArray(f, 2) = .Operator
                        filterArray(f, 3) = .Criteria2 'Delete this line to work in Excel 2010
                    End If
                End If
            End With
        Next f
    End With
End With

'Our Update code
Call Update
Call Query

'Restore current filter
For Col = 1 To UBound(filterArray(), 1)
    If Not IsEmpty(filterArray(Col, 1)) Then
        If filterArray(Col, 2) Then
            ActiveSheet.Range(CurrentFiltRange).AutoFilter Field:=Col, _
                Criteria1:=filterArray(Col, 1), _
                Operator:=filterArray(Col, 2), _
                Criteria2:=filterArray(Col, 3)
        Else
            ActiveSheet.Range(CurrentFiltRange).AutoFilter Field:=Col, _
                Criteria1:=filterArray(Col, 1)
        End If
    End If
Next Col
    
Application.ScreenUpdating = True

End Sub

Sub Update()

'Variables for Loop
Dim i As Long
Dim a, b As Range

'Variables for ADOdb
Dim Conn As ADODB.Connection
Dim RecSet As ADODB.Recordset
Dim Cmd As ADODB.Command
Dim SQLStringUpdate, Selected_Server As String

Set Conn = New ADODB.Connection
Set Cmd = New ADODB.Command

'Variables for Update Query
Dim Id, UserId As String
Dim TransDate As Date

Set a = Selection
    
If a.Count = 1 Then

    i = ActiveCell.Row
    
    Selected_Server = postgreSQL_Connection
    
    Id = ActiveWorkbook.ActiveSheet.Range("A" & i).Value
    UserId = ActiveWorkbook.ActiveSheet.Range("B" & i).Value
    TransDate = ActiveWorkbook.ActiveSheet.Range("C" & i).Value
    
    SQLStringUpdate = "UPDATE postgres.public.transaction" & _
        " SET user_id = '" & UserId & "'," & _
            " transaction_date = '" & TransDate & "'" & _
        " WHERE transaction_id = '" & Id & "';"
        
    Conn.CommandTimeout = 0
    Conn.ConnectionString = Selected_Server
    Conn.Open
    
    Cmd.ActiveConnection = Conn
    Cmd.CommandTimeout = 0
    Cmd.CommandText = SQLStringUpdate
    Set RecSet = Cmd.Execute
    
    Conn.Close
        
Else

    Set a = Selection.SpecialCells(xlCellTypeVisible)
    
    For Each b In a.Rows

    i = b.Row
    
        Selected_Server = postgreSQL_Connection
        
        Id = ActiveWorkbook.ActiveSheet.Range("A" & i).Value
        UserId = ActiveWorkbook.ActiveSheet.Range("B" & i).Value
        TransDate = ActiveWorkbook.ActiveSheet.Range("C" & i).Value
        
        SQLStringUpdate = "UPDATE postgres.public.transaction" & _
            " SET user_id = '" & UserId & "'," & _
                " transaction_date = '" & TransDate & "'" & _
            " WHERE transaction_id = '" & Id & "';"
            
        Conn.CommandTimeout = 0
        Conn.ConnectionString = Selected_Server
        Conn.Open
        
        Cmd.ActiveConnection = Conn
        Cmd.CommandTimeout = 0
        Cmd.CommandText = SQLStringUpdate
        Set RecSet = Cmd.Execute
        
        Conn.Close
        
    Next b

End If

End Sub

Sub Query()

'Variables for ADOdb
Dim Conn As ADODB.Connection
Dim RecSet As ADODB.Recordset
Dim Cmd As ADODB.Command
Dim SQLStringQuery, Selected_Server As String
Dim i As Long

Set Conn = New ADODB.Connection
Set RecSet = New ADODB.Recordset
Set Cmd = New ADODB.Command

Selected_Server = postgreSQL_Connection
SQLStringQuery = "SELECT * FROM postgres.public.transaction ORDER BY transaction_id"

Conn.CommandTimeout = 0
Conn.ConnectionString = Selected_Server
Conn.Open

Cmd.ActiveConnection = Conn
Cmd.CommandTimeout = 0
Cmd.CommandText = SQLStringQuery
Set RecSet = Cmd.Execute

ActiveWorkbook.ActiveSheet.Select
ActiveWorkbook.ActiveSheet.Cells.ClearContents

'Header
For i = 0 To RecSet.Fields.Count - 1
    ActiveWorkbook.ActiveSheet.Cells(1, i + 1).Value = RecSet.Fields(i).Name
Next

'Detail
ActiveWorkbook.ActiveSheet.Range("A2").CopyFromRecordset RecSet

RecSet.Close
Conn.Close

End Sub