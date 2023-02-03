Attribute VB_Name = "MakeReport_3"
Option Explicit


Sub MakeReport_3()
    'Розрахунок витрат на відрядження для групи працівників
    'macros by Ivakhiv Roman - megazorch@gmail.com
    Dim myCount As Integer
    Dim StartTime As Date, EndTime As Date
    Dim strA As Variant
    Dim strB As Variant
    Dim totTime As Variant
    Dim minutes As Integer
    Dim seconds As Double
    Dim result As String
        
    StartTime = Timer
    myCount = 0
    For Each strA In Range("P.I.B.")
        If strA = Empty Then
            Exit For
        Else
            myCount = myCount + 1
        End If
    Next strA
    
    Dim rgData As Range
    Dim objWord As Object
    Dim i As Integer
    Dim fileName As String
    Dim objRange As Variant
    Dim objDoc As Variant
    Dim objSelection As Variant
    Dim numOfRows As Single
    Dim numOfColumns As Single
    Dim pos As Integer
    Dim pos2 As Integer
    Dim j As Integer
    Dim savePTH As String
    Dim check As Variant
    
'    i = 0
'    For Each check In Range("sep_calc")  '16.09.2022 - #42
'        If check = True Then
'            i = i + 1
'        Else
'            i = i + 1
'            If i <= myCount Then
'                MsgBox "Виникла помилка у розрахунках витрах. Перевірте формули!", vbCritical, "Помилка!"
'                Exit Sub
'            Else
'                Exit For
'            End If
'        End If
'    Next check
        
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Add
    objWord.Visible = False
        
    'Поля
    With objDoc.PageSetup
        .LeftMargin = objWord.Application.CentimetersToPoints(Range("marg_left_5"))   '16.09.2022 - №34
        .RightMargin = objWord.Application.CentimetersToPoints(Range("marg_right_5"))
        .TopMargin = objWord.Application.CentimetersToPoints(Range("marg_top_5"))
        .BottomMargin = objWord.Application.CentimetersToPoints(Range("marg_bottom_5"))
        .Orientation = 1 'wdOrientLandscape = 1
    End With
        
    'ШАПКА
    With objWord.Selection
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1)
        .ParagraphFormat.SpaceAfter = objWord.Application.LinesToPoints(0)
        .Font.Size = 10
        .Font.name = "Times New Roman"
        .Font.Italic = True
        .ParagraphFormat.Alignment = 0 'Left
        .ParagraphFormat.LeftIndent = objWord.Application.CentimetersToPoints(21) 'отступ от края страницы
        .TypeText Text:="Додаток № 5"
        .TypeParagraph
        .TypeText Text:="до Положення про оформлення"
        .TypeParagraph
        .TypeText Text:="підзвітних сум працівників ТОВ"
        .TypeParagraph
        .TypeText Text:="""Оператор ГТС України"""
        .Font.Italic = False
        .TypeParagraph
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1.5)
        .ParagraphFormat.LeftIndent = objWord.Application.CentimetersToPoints(0)
        .ParagraphFormat.SpaceAfter = 10
        .TypeParagraph
        .Font.Bold = True 'Заголовок
        .Font.Size = 18
        .ParagraphFormat.Alignment = 1 'Center
        .TypeText Text:="Розрахунок витрат на відрядження для групи працівників"
        .Font.Bold = False 'Основний текст
        .Font.Size = 12
        .TypeParagraph
        .ParagraphFormat.Alignment = 0 'Left
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(1.25)
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1.15)
        .TypeText Text:="Згідно з наказом № " & Range("order") & " від " _
        & Range("order_date") & " та службової записки " & WorksheetFunction.Trim(Range("num_of_memo")) & " виданий працівникам, яких направлено у відрядження до " _
        & Range("one_place") & "." '17.08.2022 - №15; 13.09.2022 - №31
        .TypeParagraph
        .ParagraphFormat.Alignment = 1 'Center
    End With
    
    Set objRange = objDoc.Range
        
    numOfRows = myCount + 2
    numOfColumns = 10
    pos = objDoc.Paragraphs.Count   'узнаем сколько параграфов уже есть
    pos2 = objDoc.Paragraphs(pos).Range.End 'координаты конечной точки
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    
    With objDoc.Tables(1)
        .Borders.Enable = True
        .Columns(1).SetWidth objWord.Application.CentimetersToPoints(1.25), 0
        .Columns(2).SetWidth objWord.Application.CentimetersToPoints(4.25), 0
        .Columns(3).SetWidth objWord.Application.CentimetersToPoints(4.25), 0
        .Columns(4).SetWidth objWord.Application.CentimetersToPoints(3.25), 0
        .Columns(5).SetWidth objWord.Application.CentimetersToPoints(2.25), 0
        .Columns(6).SetWidth objWord.Application.CentimetersToPoints(3), 0
        .Columns(7).SetWidth objWord.Application.CentimetersToPoints(2.25), 0
        .Columns(8).SetWidth objWord.Application.CentimetersToPoints(2), 0
        .Columns(9).SetWidth objWord.Application.CentimetersToPoints(2), 0
        .Columns(10).SetWidth objWord.Application.CentimetersToPoints(2.25), 0
        .Rows(1).SetHeight objWord.Application.CentimetersToPoints(1.05), 0
    End With
    
    'Заполняем таблицу
    i = 1
    j = 1

    For Each strA In Range("dodatok_5_table")
        
        With objDoc.Tables(1).Cell(Row:=i, Column:=j)
            .Range.ParagraphFormat.Alignment = 1
            .VerticalAlignment = 1
            .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
            .Range.Font.Size = 10
'            .Range.InsertAfter CStr(strA)
        End With
        
        If j >= 5 And i > 1 Then
            objDoc.Tables(1).Cell(Row:=i, Column:=j).Range.InsertAfter Format(strA, "0.00")
        Else
            objDoc.Tables(1).Cell(Row:=i, Column:=j).Range.InsertAfter CStr(strA)
        End If

        j = j + 1   'смещяемся влево
        If j > numOfColumns Then
            j = 1
            i = i + 1 'сдел ряд, первая колона
            objDoc.Tables(1).Rows(i).SetHeight objWord.Application.CentimetersToPoints(1.05), 0
            If i = numOfRows Then
                For Each strB In Range("dodatok_5_total")
                    With objDoc.Tables(1).Cell(Row:=i, Column:=j)
                        .Range.Font.Bold = True
                        .Range.ParagraphFormat.Alignment = 1
                        .Range.Font.Size = 10
                        .VerticalAlignment = 1
                        .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
                    End With
                    If j >= 5 Then
                        objDoc.Tables(1).Cell(Row:=i, Column:=j).Range.InsertAfter Format(strB, "0.00")
                    Else
                        objDoc.Tables(1).Cell(Row:=i, Column:=j).Range.InsertAfter strB
                    End If
                    j = j + 1
                Next strB
                Exit For
            End If
            If i > numOfRows Then
                Exit For
            End If
        End If
    Next strA
    
    objDoc.Tables(1).Cell(Row:=numOfRows, Column:=j - 1).Select
    
    objWord.Selection.MoveDown Unit:=5, Count:=1
    
    With objWord.Selection
        .TypeParagraph
        .ParagraphFormat.Alignment = 0 'Left
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(1.25)
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1.5)
        .TypeText "Розрахунок склав:______________________" & Range("made_the_culc") '17.08.2022 - №13; 16.09.2022 - #44
        .TypeParagraph
        .TypeText "Розрахунок перевірено, бухгалтер:______________________" & Range("bookkeeper_5") '17.08.2022 - №14; 16.09.2022 - №35
    End With
    
    savePTH = ActiveWorkbook.pATH & "\" & "Розрахунок витрат для группи працівників.docx"
    objDoc.SaveAs savePTH
              
    EndTime = Timer
    totTime = Format(EndTime - StartTime, "0.0")
    If totTime >= 60 Then
        minutes = totTime \ 60
        seconds = Format((totTime / 60 - minutes) * 60, "0")
        result = CStr(minutes) & " хв. " & CStr(seconds) & " сек."
    Else
        result = CStr(totTime) & " сек."
    End If
    
    MsgBox "Звіт успішно стоворено і збереженно в папці: " & vbCrLf & ActiveWorkbook.pATH & "," & vbCrLf & "лише за " _
    & result, vbInformation, "Готово!"
    If MsgBox("Відкрити файл?", vbYesNo, "Готово!") = vbYes Then
        If Dir(savePTH) <> "" Then
            objWord.Visible = True
            objWord.Documents.Open savePTH
            MsgBox "Файл відкрито. Натисніть на іконку Word на панелі задач.", vbInformation, "Файл відкрито."
        Else
            MsgBox "Файл пропав О_о ?!", vbCritical, "Упс!"
        End If
    Else
        objWord.Quit
        Set objWord = Nothing
    End If
End Sub


