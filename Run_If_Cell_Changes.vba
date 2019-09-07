Private Sub Worksheet_Change(ByVal Target As Range)

' Source: https://docs.microsoft.com/en-us/office/troubleshoot/excel/run-macro-cells-change

    Dim KeyCells As Range

' The variable KeyCells contains the cells to verify if change
    Set KeyCells = Range("A1:C10")

If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then

  ' --------> Insert code here <--------------

End If

End Sub
