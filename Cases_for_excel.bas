Attribute VB_Name = "Cases_for_excel"
Sub UpperCase()
    Dim rng As Range
    Set rng = Selection
    For Each cell In rng
        cell.Value = UCase(cell)
    Next cell
End Sub
Sub LowerCase()
    Dim rng As Range
    Set rng = Selection
    For Each cell In rng
        cell.Value = LCase(cell)
    Next cell
End Sub
Sub ProperCase()
    Dim rng As Range
    Set rng = Selection
    For Each cell In rng
        cell.Value = WorksheetFunction.Proper(cell)
    Next cell
End Sub
Sub CapitalizeFirstLetter()
    Dim Sel As Range
    Set Sel = Selection
    For Each cell In Sel
        cell.Value = Application.WorksheetFunction.Replace(LCase(cell.Value), 1, 1, UCase(Left(cell.Value, 1)))
    Next cell
End Sub
