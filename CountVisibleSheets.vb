Function CountVisibleSheets()
    Dim Wks As Worksheet
    Dim i As Long
    
    i = 0
    For Each xSht In ActiveWorkbook.Sheets
        If Wks.Visible Then i = i + 1
    Next
    
    VisibleSheetsCount = i
    
End Function
