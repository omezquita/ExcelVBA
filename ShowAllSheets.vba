Sub ShowAllSheets()

    Dim Wks As Worksheet
    
    For Each Wks In ThisWorkbook.Worksheets
        Wks.Visible = xlSheetVisible
    Next Wks

End Sub
