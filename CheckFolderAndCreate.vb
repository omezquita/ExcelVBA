Function CheckFolderAndCreate(path2search As String) As Boolean
' Function that receives a path as argument, it checks if the path exists and if it doesn't it creates the folder
' The function will return False if the file didn't exist and True otherwise

    If Dir(path2search, vbDirectory) = "" Then
        MkDir Path:=path2search
        CheckPathAndCreate = False
    Else
        CheckPathAndCreate = True
    End If
End Function
