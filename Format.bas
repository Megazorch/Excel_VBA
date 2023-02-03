Attribute VB_Name = "Format"
Option Explicit
'=И(F108<>"ВВ";F108<>"СВ";F108<>"РХП";F108<>"РВД";НЕ(ЕПУСТО(F108));F109<>"Х")
Sub Format_color()
Attribute Format_color.VB_ProcData.VB_Invoke_Func = " \n14"
' Format_color Макрос
    Dim i As Range
    
    Cells.FormatConditions.Delete
    
    Application.ScreenUpdating = False
    
    Dim n As Name
    Dim Count As Integer
    For Each n In ActiveWorkbook.Names
        If n.Name = "main_table" Then
            Count = Count + 1
        End If
    Next n
    If Count = 0 Then
        Dim sReturn As String
        Dim refer As Range
        sReturn = InputBox("Введите диапазон рабочей области:" & vbCrLf & "(например: H36:AJ400 или h36:aj400," & vbCrLf & "на англ. языке)", "Нужного диапазона не найдено")
        If StrPtr(sReturn) = 0 Then
            MsgBox "Вы нажали Cancel или ESC.", vbOKOnly, "Отмена действия"
        ElseIf sReturn = "" Then
            MsgBox "Вы нажали OK, данные не введены, нету начальных данных.", vbOKOnly, "Нет данных"
        Else
            MsgBox "Вы нажали OK, данные введены. Диапазону присвоено имя main_table.", vbOKOnly, "Данные введены"
            ActiveWorkbook.Names.Add Name:="main_table", RefersTo:=Range(sReturn), Visible:=True
        End If
    End If
    
    For Each i In Range("main_table")
        i.Select
        Select Case i
            Case Is = "ВВ" 'blue
                With Selection.Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.799981688894314
'                    .PatternTintAndShade = 0
                End With
                ActiveCell.Offset(1, 0).Range("A1").Value = "x"
            Case Is = "ВД"   'green
                With Selection.Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.799981688894314
'                    .PatternTintAndShade = 0
                End With
                ActiveCell.Offset(1, 0).Range("A1").Select
                With Selection.Font
                    .ThemeColor = xlThemeColorDark1
'                    .TintAndShade = 0
                End With
            Case Is = "РХП"  'green
                With Selection.Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.799981688894314
'                    .PatternTintAndShade = 0
                End With
            Case Is = "ВІ"   'pink
                With Selection.Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
                    .Color = 15049727
'                    .TintAndShade = 0
'                    .PatternTintAndShade = 0
                End With
                ActiveCell.Offset(1, 0).Range("A1").Select
                With Selection.Font
                    .ThemeColor = xlThemeColorDark1
'                    .TintAndShade = 0
                End With
            Case Is = "РВД"  'yellow
                With Selection.Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent4
                    .TintAndShade = 0.799981688894314
'                    .PatternTintAndShade = 0
                End With
            Case Is = "СВ"
                ActiveCell.Offset(1, 0).Range("A1").Select
                With Selection.Font
                    .ThemeColor = xlThemeColorDark1
'                    .TintAndShade = 0
                End With
            Case Is = "х"
                With Selection.Font         'отмена окраски шрифта
                    .ColorIndex = xlAutomatic
'                    .TintAndShade = 0
                End With
            Case "РС", "ВЧ", "РН", "НУ", "РВ", "ВДп", "ВДч", "В", "Д", "Ч", "ДІ", "ВУ", "ВЗ", "ТВ", "Н", "НБ", "ДБ", "НА", "ДО", "ВП", "ДД", "ІН", "ПК", "П", "ПР", "ТН", "НН", "НЗ", "ІВ", "І", "НПп", "С", "БЗ", "НД", "НП", "ДЛ", "ДВВ", "МО", "ПНМ", "ВН"
                ActiveCell.Offset(1, 0).Range("A1").Select  'font white
                With Selection.Font
                    .ThemeColor = xlThemeColorDark1
'                    .TintAndShade = 0
                End With
        End Select
    Next i
End Sub

Sub Clean()
    Dim n As Name
    Dim Count As Integer
    For Each n In ActiveWorkbook.Names
        If n.Name = "main_table" Then
            Count = Count + 1
        End If
    Next n
    If Count = 0 Then
        Dim sReturn As String
        Dim refer As Range
        sReturn = InputBox("Введите диапазон рабочей области:" & vbCrLf & "(например: H36:AJ400 или h36:aj400," & vbCrLf & "на англ. языке)", "Нужного диапазона не найдено")
        If StrPtr(sReturn) = 0 Then
            MsgBox "Вы нажали Cancel или ESC.", vbOKOnly, "Отмена действия"
        ElseIf sReturn = "" Then
            MsgBox "Вы нажали OK, данные не введены, нету начальных данных.", vbOKOnly, "Нет данных"
        Else
            MsgBox "Вы нажали OK, данные введены. Диапазону присвоено имя main_table.", vbOKOnly, "Данные введены"
            ActiveWorkbook.Names.Add Name:="main_table", RefersTo:=Range(sReturn), Visible:=True
        End If
    End If
    Dim i As Range
    Cells.FormatConditions.Delete
    Application.ScreenUpdating = False
    For Each i In Range("main_table")
        i.Select
        If Not ActiveCell.MergeCells Then
            With Selection.Interior     'отмена заливки
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With Selection.Font         'отмена окраски шрифта
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
            End With
        End If
    Next i
End Sub

Sub ChangeRefer()
    Dim sReturn As String
    Dim refer As Range
    sReturn = InputBox("Введите диапазон рабочей зоны:" & vbCrLf & "(например: H36:AJ400 или h36:aj400," & vbCrLf & "на англ. языке)", "Смена диапазона")
    If StrPtr(sReturn) = 0 Then
        MsgBox "Вы нажали Cancel или ESC.", vbOKOnly, "Отмена действия"
    ElseIf sReturn = "" Then
        MsgBox "Вы нажали OK, данные не введены.", vbOKOnly, "Нет данных"
    Else
        MsgBox "Вы нажали OK, данные введены. Ссылка на диапазон изменена.", vbOKOnly, "Данные введены"
        ActiveWorkbook.Names.Add Name:="main_table", RefersTo:=Range(sReturn), Visible:=True
    End If
End Sub
