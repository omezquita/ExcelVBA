Sub activateribbon()

' Purpose: Macro to show Excel in its regular state (Display ribbon, formula bar, and worksheets)
' Author: Orlando Mezquita
' Date: 19JAN15

    ActiveWindow.DisplayHeadings = True
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayWorkbookTabs = True

End Sub
