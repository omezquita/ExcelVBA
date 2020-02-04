  

Sub Test()
  ExportImage_Func "A1:AA65", "WorksheetName", "FileName"
End Sub


Sub ExportImage_Func(ByVal Rng_Str As String, ByVal Wks_Name As String, ByVal FlName As String)
' Works for Excel 2016+
' Rng_Str: String with the A1 reference for the range to copy
' Wks_Name: String with the name of the worksheet where Rng_Str is located
' FlName: Name of picture file WITHOUT extension

Dim Sheet As Worksheet
Dim Rng As Range

Set Sheet = Worksheets(Wks_Name)
Set Rng = Sheet.Range(Rng_Str) '<-------- ENTER RANGE TO SAVE HERE

TempPicFile = CreateObject("WScript.Shell").specialfolders("Desktop") & "\" & FlName & ".png"

' convert snapshot to picture
Rng.CopyPicture xlPrinter, xlPicture

zoom_coef = 100 / Sheet.Parent.Windows(1).Zoom
lWidth = Rng.Width '* zoom_coef
lHeight = Rng.Height '* zoom_coef

Set Cht = ActiveSheet.ChartObjects.Add(Left:=0, Top:=0, Width:=lWidth, Height:=lHeight)
Cht.Activate

With Cht.Chart
  .Paste
  .Export Filename:=TempPicFile, Filtername:="PNG"
End With

Cht.Delete
End Sub

