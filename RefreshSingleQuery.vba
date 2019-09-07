Sub Refresh_Division()
'
' Refresh a Power Query query whose result is on cell A2 of the worksheet Wks Name

    Worksheet("Wks Name").Range("A2").ListObject.QueryTable.Refresh BackgroundQuery:=False
End Sub
