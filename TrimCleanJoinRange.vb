Function TrimCleanJoinRange(ByVal Sep As String, _
                            ByVal Rng As Range, _
                            Optional ByVal prefix As String = "", _
                            Optional ByVal suffix As String = "") As String

Dim cl As Range
Dim S As String

S = ""
  'To use the VBA version of Trim remove the period in front of .Trim
For Each cl In Rng.Cells
        With Application.WorksheetFunction
            S = S & prefix & .Trim(.Clean(cl.Value)) & suffix & Sep
        End With
Next cl

TrimCleanJoinRange = Left(S, Len(S) - Len(Sep))

End Function
