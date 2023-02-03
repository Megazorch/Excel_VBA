Attribute VB_Name = "CreatLog"
Option Explicit

Sub Create_Log()
    'macros by Ivakhiv Roman - megazorch@gmail.com
    Dim count As Integer
    Dim StartTime As Date, EndTime As Date
    Dim intA As Variant
    Dim strB As Variant
    Dim totTime As Variant
    Dim minutes As Integer
    Dim seconds As Double
    Dim result As String
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
    Dim myPgNum As Variant
    
    StartTime = Timer
    count = 0
    For Each intA In Range("table")
        If intA > 0 Then
            count = count + 1
        End If
    Next intA
    
    Call FillTable
    
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Add
    objWord.Visible = False
    
    Set myPgNum = objDoc.Sections(1).Footers(1) _
        .PageNumbers.Add(PageNumberAlignment:=1, FirstPage:=True) 'wdHeaderFooterPrimary,wdAlignPageNumberCenter
    myPgNum.Select
    With objWord.Selection.Range
        .Font.Name = "Times New Roman"
        .Font.Size = 12
    End With
    
    objWord.ActiveWindow.Panes(1).Activate
    
    'Поля
    With objDoc.PageSetup
        .LeftMargin = objWord.Application.CentimetersToPoints(1)
        .RightMargin = objWord.Application.CentimetersToPoints(1)
        .TopMargin = objWord.Application.CentimetersToPoints(1)
        .BottomMargin = objWord.Application.CentimetersToPoints(1)
        .Orientation = 1 'wdOrientLandscape = 1
    End With
        
    'ШАПКА
    With objWord.Selection
        .Font.Name = "Times New Roman"
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1)
        .ParagraphFormat.LeftIndent = objWord.Application.CentimetersToPoints(0)
        .ParagraphFormat.SpaceAfter = 0
        .TypeParagraph
        .Font.Bold = True 'Заголовок
        .Font.Size = 14
        .ParagraphFormat.Alignment = 1 'Center
        .TypeText Text:="ЖУРНАЛ"
        .TypeParagraph
        .TypeText "обліку часу роботи із шкідливими і важкими умовами праці"
        .TypeParagraph
        .TypeText "дільниці лінійно-експлуатаційної служби"
        .TypeParagraph
        .TypeText "Одеського промислового майданчика Миколаївського ЛВУМГ"
        .TypeParagraph
        .TypeText "за " & Range("month") & " місяць " & Range("year") & " року"
        .Font.Bold = False 'Основний текст
        .Font.Size = 12
        .TypeParagraph
        .ParagraphFormat.Alignment = 1 'Center
        .TypeParagraph
    End With
    
    Set objRange = objDoc.Range

    numOfRows = count + 5
    numOfColumns = 7
    pos = objDoc.Paragraphs.count   'узнаем сколько параграфов уже есть
    pos2 = objDoc.Paragraphs(pos).Range.End 'координаты конечной точки
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    objWord.Selection.Tables(1).Rows(1).HeadingFormat = True        'колонка переходит на страницы
    objWord.Selection.Tables(1).Rows.AllowBreakAcrossPages = False  'нельзя переносить строку на след страницу

    With objDoc.Tables(1)
        .Borders.Enable = True
        .Columns(1).SetWidth objWord.Application.CentimetersToPoints(0.99), 0
        .Columns(2).SetWidth objWord.Application.CentimetersToPoints(2.5), 0
        .Columns(3).SetWidth objWord.Application.CentimetersToPoints(5.25), 0
        .Columns(4).SetWidth objWord.Application.CentimetersToPoints(5.75), 0
        .Columns(5).SetWidth objWord.Application.CentimetersToPoints(6.25), 0
        .Columns(6).SetWidth objWord.Application.CentimetersToPoints(3), 0
        .Columns(7).SetWidth objWord.Application.CentimetersToPoints(3.75), 0
    End With

    'Заполняем таблицу
    i = 1
    j = 1

    For Each intA In Range("ready")
        If i = 1 Then
            With objDoc.Tables(1).Cell(Row:=i, Column:=j)
                .Range.ParagraphFormat.Alignment = 1
                .VerticalAlignment = 1
                .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
                .Range.Font.Size = 12
                .Range.Font.Bold = True
                .Range.InsertAfter CStr(intA)
            End With
        ElseIf intA = "Stop" Then
            Exit For
        Else
            With objDoc.Tables(1).Cell(Row:=i, Column:=j)
                .Range.ParagraphFormat.Alignment = 1
                .VerticalAlignment = 1
                .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
                .Range.Font.Size = 12
                .Range.Font.Bold = False
                .Range.InsertAfter CStr(intA)
            End With
        End If

        j = j + 1   'смещяемся влево
        If j > numOfColumns Then
            j = 1
            i = i + 1 'сдел ряд, первая колона
            If i > numOfRows Then
                Exit For
            End If
        End If
    Next intA
        
    i = i + 1
    For Each intA In Range("totall")
        With objDoc.Tables(1).Cell(Row:=i, Column:=j)
            .Range.ParagraphFormat.Alignment = 1
            .VerticalAlignment = 1
            .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
            .Range.Font.Size = 12
            .Range.Font.Bold = False
            .Range.Font.Italic = True
            .Range.InsertAfter CStr(intA)
        End With
        j = j + 1   'смещяемся влево
        If j > numOfColumns Then
            j = 1
            i = i + 1 'сдел ряд, первая колона
            If i > numOfRows Then
                Exit For
            End If
        End If
    Next intA
    
    objDoc.Tables(1).Rows(i - 4).Select 'обьединить ячейки
    objWord.Selection.Cells.Merge
    With objDoc.Tables(1).Cell(Row:=i - 4, Column:=j)
        .Range.ParagraphFormat.Alignment = 1
        .VerticalAlignment = 1
        .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
        .Range.Font.Size = 12
        .Range.Font.Bold = True
        .Range.Font.Italic = True
        .Range.InsertAfter "ВСЬОГО за " & Range("month") & " місяць " & Range("year") & " р.:"
    End With

    objDoc.Tables(1).Cell(Row:=numOfRows, Column:=j - 1).Select

    objWord.Selection.MoveDown Unit:=5, count:=1

    With objWord.Selection
        .TypeParagraph
        .TypeParagraph
        .ParagraphFormat.Alignment = 0 'Left
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(1.25)
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1.5)
        .ParagraphFormat.SpaceAfter = 10
        .TypeText "Начальник дільниці ЛЕС, відповідальна особа:______________________ " & Range("boss")
        .TypeParagraph
        .TypeText "Служба ОП, ПБ та НС :__________________________________________ Світлана БУЙЛУК"
    End With
    
    If objWord.ActiveWindow.View.SplitSpecial = 0 Then 'wdPaneNone
        objWord.ActiveWindow.ActivePane.View.Type = 3 'wdPrintView
    Else
        objWord.ActiveWindow.View.Type = 3
    End If
    
    savePTH = ActiveWorkbook.pATH & "\" & "Журнал ШУП за " & Range("month") & "_" & Range("year") & ".docx"
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
    
    MsgBox "Журнал успішно стоворено і збереженно в папці: " & vbCrLf & ActiveWorkbook.pATH & "," & vbCrLf & "лише за " _
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


