Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
' From: https://stackoverflow.com/questions/38267950/check-if-a-value-is-in-an-array-or-not-with-excel-vba/38268261
Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function
