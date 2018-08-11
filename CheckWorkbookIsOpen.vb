Function wkbisopen(ByVal wkbname As String) As Boolean
'Returns true if wkbname is open

	Dim wBook As Workbook
	On Error Resume Next
	
	Set wBook = Workbooks(wkbname)
	wkbisopen = Not (wBook Is Nothing)

End Function
