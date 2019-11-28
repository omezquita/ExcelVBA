Function Copy_Without_Open() As String
' Function to create a copy of a file with a new name and returns the path and name of the copy file.
    Dim PTT_Path As String
    Dim Project_Tracking_Tool As String
    Dim Original_Path As String, Copy_Path As String
    Dim Original_Name As String, Copy_Name As String
    Dim Original_Wkb As String, Copy_Wkb As String

    Original_Path = "Enter the path for the file to copy, end the path with \"
    Copy_Path = "Enter the destination path for the copy file, end the path with \"
       
    Original_Name = "Enter the filename of the file to copy including the extension"
    Copy_Name = "Enter the filename of copy file including the extension"
    
    Original_Wkb = Original_Path & Original_Name
    Copy_Wkb = Copy_Path & Copy_Name
    
    FileCopy Original_Wkb, Copy_Wkb
    
    Copy_Without_Open = Copy_Wkb

End Function
