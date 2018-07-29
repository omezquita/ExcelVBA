Function Openfl(message2display As String)
' Purpose: Opens an Excel file selected by the user, returns the name of the file selected
' Created by: Orlando Mezquita
' Parameters:
'            message2display = Message to display in a message box and when the file will be selected

Dim filefiltertext As String
    'File filter for dialog box, modify appropriately
     filefiltertext = "Excel Files, *.xml"
     MsgBox message2display, vbInformation + vbOKOnly, "Browse for file"
     flpath = Application.GetOpenFilename(Title:=message2display, FileFilter:=filefiltertext)
     patharray = Split(flpath, "\")
     Openfl = patharray(UBound(patharray))
     Workbooks.Open Filename:=flpath, UpdateLinks:=0
End Function
