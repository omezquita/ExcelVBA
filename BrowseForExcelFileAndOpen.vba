Function Openfl(message2display As String, _
                Optional ByVal pathonly As Boolean = False, _
                Optional ByVal fileonly As Boolean = True)
'Parameters:
'   message2display: Message to display in a message box and when the file will be selected
'   pathonly: if true the file won't be opened and only the file path will be returned
'   fileonly: if true the filename will be returned instead of the full path

Dim filefiltertext As String
    'File filter for dialog box, modify appropriately
     filefiltertext = "Excel Files, *.xlsx"
     MsgBox message2display, vbInformation + vbOKOnly, "Browse for file"
     flpath = Application.GetOpenFilename(Title:=message2display, FileFilter:=filefiltertext)
     patharray = Split(flpath, "\")
     Openfl = patharray(UBound(patharray))
     If Openfl = "False" Then End
     
     If Not (fileonly) Then Openfl = flpath
     If Not (pathonly) Then Workbooks.Open Filename:=flpath, UpdateLinks:=0
End Function
