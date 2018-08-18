Function TrimAndJoin(ByVal Sep As String, _
                     ByRef Arr() As String, _
                     Optional ByVal prefix As String = "", _
                     Optional ByVal suffix As String = "") As String

Dim c As Integer
Dim S As String

    c = LBound(Arr)
    S = ""
   Do While Arr(c) <> ""
    ' To use the VBA version of Trim remove the period in front of .Trim.
        With Application.WorksheetFunction
            S = S & prefix & .Trim(.Clean(Arr(c))) & suffix & Sep
            c = c + 1
        End With
    Loop

TrimAndJoin = Left(S, Len(S) - 2)

End Function
