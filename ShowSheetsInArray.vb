Sub ShowSheets(shArray As Variant)

' To pass arguments use:
'    ShowSheets Array("Sheet1", "Sheet2")


Dim wks As Worksheet
Dim pos As Variant

For Each wks In ThisWorkbook.Worksheets
    pos = Application.Match(wks.Name, shArray, False)
    If IsError(pos) Then
        wks.Visible = xlSheetHidden
    Else
        wks.Visible = xlSheetVisible
    End If
Next wks

End Sub
