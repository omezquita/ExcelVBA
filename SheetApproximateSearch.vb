Function searchsheet(ByVal txt2find As String) As String
'Returns the sheet name of the first sheet that contains the text 'txt2find

Dim ws As Worksheet
Dim shname As String
Dim found As Boolean
Dim searchtype As Integer, totalfound As Integer

found = False

    For Each ws In ActiveWorkbook.Worksheets
        If UCase(ws.Name) Like "*" & UCase(txt2find) & "*" Then
            searchsheet = ws.Name
            found = True
        End If
    Next ws
    
If found = False Then
    searchsheet = "Not Found"
End If

End Function
