 Function FormatSheetNames(Wkb As Workbook, fn As String)
    
 ' To call function:
 ' Example:  Change all names of current workbook to UPPER case
 '        FormatSheetNames ActiveWorkbook, "UPPER"
 ' Parameters:
 '   *Wkb = Workbook
 '   *fn  = Name of Text Function (UPPER, PROPER, LOWER, TRIM, CLEAN)
 
    Dim Wks As Worksheet
    
    For Each Wks In Wkb.Worksheets
        Wks.Name = Application.Evaluate("=" & fn & "(""" & Wks.Name & """)")
    Next Wks
 
 End Function
