Attribute VB_Name = "First_step"
Option Explicit

Sub FillTable()
    Dim i As Variant
    Dim x As Variant
    Dim offRow As Integer, offCol As Integer
    Dim count As Integer
    
    offRow = 1
    offCol = 4
    count = 0
    
    Debug.Print count
    
    For Each i In Range("date")
        i.Select
        Debug.Print i
        If ActiveCell.Offset(1, 0).Range("A1").Value > 0 Then
            count = count + 1
            Sheets("Base").Range("A1").Offset(offRow, offCol).Value = count
            Sheets("Base").Range("A1").Offset(offRow, offCol + 1).Value = i _
            & "." & Application.WorksheetFunction.VLookup(Range("month"), Range("all_months"), 2, 1) & "." & Range("year")
            Sheets("Base").Range("A1").Offset(offRow, offCol + 2).Value = "газозварник/газорізальник"
            Sheets("Base").Range("A1").Offset(offRow, offCol + 3).Value = Range("first_guy")
            Sheets("Base").Range("A1").Offset(offRow, offCol + 5).Value = ActiveCell.Offset(1, 0).Range("A1")
            Sheets("Base").Range("A1").Offset(offRow, offCol + 6).Value = 1
            offRow = offRow + 1
        End If
        If ActiveCell.Offset(2, 0).Range("A1").Value > 0 Then
            count = count + 1
            Sheets("Base").Range("A1").Offset(offRow, offCol).Value = count
            Sheets("Base").Range("A1").Offset(offRow, offCol + 1).Value = i _
            & "." & Application.WorksheetFunction.VLookup(Range("month"), Range("all_months"), 2, 1) & "." & Range("year")
            Sheets("Base").Range("A1").Offset(offRow, offCol + 2).Value = "електрозварник ручного зварювання"
            Sheets("Base").Range("A1").Offset(offRow, offCol + 3).Value = Range("second_guy")
            Sheets("Base").Range("A1").Offset(offRow, offCol + 5).Value = ActiveCell.Offset(2, 0).Range("A1")
            Sheets("Base").Range("A1").Offset(offRow, offCol + 6).Value = 1
            offRow = offRow + 1
        End If
        If ActiveCell.Offset(3, 0).Range("A1").Value > 0 Then
            count = count + 1
            Sheets("Base").Range("A1").Offset(offRow, offCol).Value = count
            Sheets("Base").Range("A1").Offset(offRow, offCol + 1).Value = i _
            & "." & Application.WorksheetFunction.VLookup(Range("month"), Range("all_months"), 2, 1) & "." & Range("year")
            Sheets("Base").Range("A1").Offset(offRow, offCol + 2).Value = "електрогазозварник"
            Sheets("Base").Range("A1").Offset(offRow, offCol + 3).Value = Range("third_guy")
            Sheets("Base").Range("A1").Offset(offRow, offCol + 5).Value = ActiveCell.Offset(3, 0).Range("A1")
            Sheets("Base").Range("A1").Offset(offRow, offCol + 6).Value = 1
            offRow = offRow + 1
        End If
    Next i
    
    Sheets("Base").Range("A1").Offset(offRow, offCol).Value = "Stop"
    
    count = 0
    For Each i In Range("totall")
        count = count + 1
        Select Case count
            Case 3
                i.Value = "газозварник/газорізальник"
            Case 4
                i.Value = Range("first_guy")
            Case 6
                i.Value = Range("first_hours")
            Case 7
                i.Value = Range("first_days")
            Case 10
                i.Value = "електрозварник ручного зварювання"
            Case 11
                i.Value = Range("second_guy")
            Case 13
                i.Value = Range("second_hours")
            Case 14
                i.Value = Range("second_days")
            Case 17
                i.Value = "електрогазозварник"
            Case 18
                i.Value = Range("third_guy")
            Case 20
                i.Value = Range("third_hours")
            Case 21
                i.Value = Range("third_days")
        End Select
    Next i
End Sub
