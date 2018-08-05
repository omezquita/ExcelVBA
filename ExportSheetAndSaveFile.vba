Sub ExportSheetAndSaveFile(ByVal Wks_Name As String)
  ' Requires the function SaveFileAs located in this repository as well  
    Sheets(Wks_Name).Copy
    SaveFileAs "Enter the name and location of the Rules file"
    
End Sub
