Sub CheckVersion()

Dim Ver as Single

Ver = Application.Version

if Ver = 15 then
 msgbox "Excel 2013"
End if

'Possible Values:
' 11.0 = 2003
' 12.0 = 2007
' 14.0 = 2010
' 15.0 = 2013

End Sub
