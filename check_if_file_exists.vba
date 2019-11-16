Function Check_if_file_exists(fullpath As String) As Boolean
' Returns true if file exists and false otherwise
    
    Check_if_file_exists = Dir(fullpath) <> ""
End Function
