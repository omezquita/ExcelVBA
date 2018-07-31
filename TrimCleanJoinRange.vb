Function TrimCleanJoinRange(ByVal Sep As String, ByVal Rng As Range) As String

Dim cl As Range
Dim S As String

S = ""
For Each cl In Rng.Cells
        With Application.WorksheetFunction
            S = S & .Trim(.Clean(cl.Value)) & Sep
        End With
Next cl

TrimCleanJoinRange = Left(S, Len(S) - Len(Sep))

End Function
