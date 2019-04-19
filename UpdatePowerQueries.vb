Public Sub UpdatePowerQueriesOnly()
    Dim lTest As Long, cn As WorkbookConnection
    On Error Resume Next
    
    For Each cn In ThisWorkbook.Connections
        lTest = InStr(1, cn.OLEDBConnection.Connection, "Provider=Microsoft.Mashup.OleDb.1")
        If Err.Number <> 0 Then
        Err.Clear
    Exit For
    End If
    
    If lTest > 0 Then cn.Refresh
    Next cn

End Sub
