Sub ExportIntroImage()
' Works for Excel 2016+
Dim Rng As Range
Dim Sheet As Worksheet

Set Sheet = ActiveSheet
Set Rng = Sheet.Range("AE4:GR123") '<-------- ENTER RANGE TO SAVE HERE

TempPicFile = CreateObject("WScript.Shell").specialfolders("Desktop") & "\" & "Intro" & ".png"

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
