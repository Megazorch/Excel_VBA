Attribute VB_Name = "Deactivation"
Option Explicit

Private Sub Workbook_Deactivate()   '17.08.2022
    Word.Application.Quit
    Set Word.Application = Nothing
End Sub
