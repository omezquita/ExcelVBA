Sub SaveFileAs(ByVal DialogTitle As String)
' Modified version of: https://stackoverflow.com/questions/29522278/save-as-dialog-excel-code

    Dim varResult As Variant
    Dim ActBook As Workbook

    'displays the save file dialog
    varResult = Application.GetSaveAsFilename(FileFilter:= _
             "Excel Files (*.xlsx), *.xlsx", Title:=DialogTitle)

    'checks to make sure the user hasn't canceled the dialog
    If varResult <> False Then
        ActiveWorkbook.SaveAs Filename:=varResult, _
        FileFormat:=xlWorkbookDefault
        Exit Sub
    End If
End Sub
