Function TrimAndJoin(ByRef Arr() As String) As String

Dim c As Integer
Dim S As String

    c = LBound(Arr)
    S = ""
   Do While Arr(c) <> ""
        With Application.WorksheetFunction
            S = S & .Trim(.Clean(Arr(c))) & ", "
            c = c + 1
        End With
    Loop

TrimAndJoin = Left(S, Len(S) - 2)

End Function
