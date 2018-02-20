Sub deactivateribbon()
' Purpose: Macro hide the ribbon, the formula bar and the worksheets
' Author: Orlando Mezquita
' Date: 19JAN15

Dim hide_worksheet_tabs As Boolean
Dim hide_headings As Boolean

' Modify these variables to hide the tabs and the headings
hide_worksheet_tabs = True
hide_headings = True

ActiveWindow.DisplayHeadings = Not (hide_headings)
Application.DisplayFullScreen = True
Application.DisplayFormulaBar = False
ActiveWindow.DisplayWorkbookTabs = Not (hide_worksheet_tabs)

End Sub
