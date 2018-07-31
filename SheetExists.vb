 Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
 ' Source of code = https://stackoverflow.com/questions/6688131/test-or-check-if-sheet-exists
 
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     SheetExists = Not sht Is Nothing
 End Function
