Option Explicit

Private Sub Workbook_Open()
    VbaRegexWorksheetFunctions.RegisterRxFunctions
End Sub

Private Sub Workbook_AddinInstall()
    VbaRegexWorksheetFunctions.RegisterRxFunctions
End Sub
