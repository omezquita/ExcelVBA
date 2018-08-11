 Function FormatSheetNames(Wkb As Workbook, fn As String)
    
 ' To call function:
 ' Example:  Change all names of current workbook to UPPER case
 '        FormatSheetNames ActiveWorkbook, "UPPER"
 ' Parameters:
 '   *Wkb = Workbook
 '   *fn  = Name of Text Function (UPPER, PROPER, LOWER, TRIM, CLEAN)
 
    Dim wks As Worksheet
    
    For Each wks In Wkb.Worksheets
        wks.Name = Application.Evaluate("=" & fn & "(""" & wks.Name & """)")
    Next wks
 
 End Function
