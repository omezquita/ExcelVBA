Sub fillvalues(inicell As Range, ByVal filas As Integer, ByVal val As String)
' Purpose: Fill a range with a value or formula starting from "inicell" and continuing through "filas" number of cell
' Assumption: actions will be executed in the active worksheet
' Parameters:
'            inicell = Starting cell, this will be the first cell where the value/formula is placed
'            filas = Number of cells where the value/formula will be placed (includes inicell)
'            val = Value or formula to paste. If is a formula it should be in R1C1 style

If val <> "" Then inicell.FormulaR1C1 = val
    inicell.Select
    Selection.AutoFill Destination:=Range(Selection, Selection.Offset(filas - 1, 0)), Type:=xlFillCopy
    inicell.Select
End Sub
