Attribute VB_Name = "TEO_Report"
Option Explicit

Sub ReportToWord()
    Dim myPgNm As Variant
    Dim objWord As Object
    Dim objRange As Variant
    Dim i As Integer
    Dim j As Integer
    Dim objDoc As Variant
    Dim objTable As Variant
    Dim numOfRows As Integer
    Dim numOfColumns As Integer
    Dim pos As Integer
    Dim pos2 As Long
    Dim strA As Variant
    Dim strB As Variant
    Dim count1 As Integer
    Dim count2 As Integer
    Dim nameArray() As Variant   'B - Найменування основного засобу
    Dim nameArray2() As Variant  'C - Родовий Відміннок - Кого?
    Dim sapNum() As Variant      'D - SAP Номер
    Dim reason2() As Variant     'F - Причина - для Тех Стану
    Dim reason() As Variant      'G - Причина "відновлювального ремонту" - Чого?
    Dim years() As Variant       'J - Лет обладнанню
    Dim explDate() As Variant    'K - Експлуатація (дата словами)
    Dim vartist() As Variant     'N - Вартість
    Dim vaga() As Variant        'O - Вага
    Dim typeBryht() As Variant   'P - Тип брухту
    Dim kindBryht() As Variant   'Q - Вид брухту
    Dim kilkist() As Variant     'R - Кількість
    Dim price() As Variant       'S - Ціна покупки брухту
    Dim serialN() As Variant     'T - Серійний номер
    Dim nameBryht() As Variant   'U - Назва обладнання (як брухт)
    
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Add
    objWord.Visible = True
    Set objRange = objDoc.Range
    numOfRows = 1
    numOfColumns = 2
    
    'Поля
    With objDoc.PageSetup
        .LeftMargin = 85.0394
        .RightMargin = 28.34646
        .TopMargin = 56.69291
        .BottomMargin = 56.69291
    End With
        
    'ШАПКА
    objDoc.Tables.Add objRange, numOfRows, numOfColumns
    Set objTable = objDoc.Tables(1)
    objTable.Borders.Enable = False
    objTable.Cell(1, 2).Range.Select
    objTable.Cell(1, 2).Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
    objTable.Cell(1, 2).SetWidth 212.5984, 0         'wdAdjustNone = 0
    With objWord.Selection
        .ParagraphFormat.LineSpacing = 18 'межстроч интервал 1,5
        .Font.Size = 12
        .Font.Name = "Times New Roman"
        .TypeText Text:="ЗАТВЕРДЖУЮ"
        .TypeParagraph
        .TypeText Text:="Начальник Миколаївського ЛВУМГ"
        .TypeParagraph
        .TypeText Text:="_______________"
        .Font.Bold = True
        .TypeText Text:="Євген ЛИТВИНЮК"
        .Font.Bold = False
        .TypeParagraph
        .TypeText Text:="""__""_________2021 р."
        .MoveDown Unit:=5, count:=2           'оглавление
        .ParagraphFormat.SpaceAfter = 0     'удалить интервал после абзаца
        .Font.Size = 12
        .Font.Name = "Times New Roman"
        .TypeParagraph
        .ParagraphFormat.Alignment = 1  'Center
        .Font.Bold = True
        .TypeText Text:="ТЕХНІКО-ЕКОНОМІЧНЕ ОБҐРУНТУВАННЯ ДОЦІЛЬНОСТІ СПИСАННЯ МАЙНА ТОВАРИСТВА"
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .ParagraphFormat.Alignment = 3    'Justify
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(1.25)
        .TypeText Text:="Постійно діюча комісія з питань списання державного та власного майна, яке обліковується на балансі Товариства, створена наказом Миколаївського ЛВУМГ ТОВ ""Оператор ГТС України"" від 15 липня 2021 р. №417 у складі:"
        .TypeParagraph
        .TypeParagraph
    End With
        
    'Вторая таблица с комисией
    numOfRows = 9
    numOfColumns = 3
    pos = objDoc.Paragraphs.count   'узнаем сколько параграфов уже есть
    pos2 = objDoc.Paragraphs(pos).Range.End 'координаты конечной точки
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    
    With objDoc.Tables(2)
        .Borders.Enable = False
        .Columns(1).SetWidth 99.2126, 0         '3.5cm
        .Columns(2).SetWidth 233.8583, 0        '8.25cm
        .Columns(3).SetWidth 148.8189, 0        '5.25cm
        .Columns(3).Cells.VerticalAlignment = 1 'wdAlignVerticalCenter
    End With
    
    'Заполняем таблицу
    i = 1
    j = 1

    For Each strA In Worksheets("Тех Акт (таблиці)").Range("AJ5:AL13")
        objDoc.Tables(2).Cell(Row:=i, Column:=j).Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
        objDoc.Tables(2).Cell(Row:=i, Column:=j).Range.InsertAfter strA
        j = j + 1   'смещяемся влево
        If j > numOfColumns Then
            j = 1
            i = i + 1 'сдел ряд, первая колона
            If i > numOfRows Then
                i = 1
            End If
        End If
    Next strA
    
    objWord.Selection.MoveDown Unit:=5, count:=9
    
    'продолжение вступительно части
    With objWord.Selection
        .TypeParagraph
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(0)
        .TypeText Text:="провела опрацювання матеріалів та обстеження наступних об'єктів: "
    End With
    
    count1 = 0   'считаю сколько елементов для запятой
    For Each strA In Worksheets("На списання").Range("C2:C24")
        If strA = Empty Then
            Exit For
        End If
        count1 = count1 + 1
    Next strA
    
    count2 = 0
    For Each strA In Worksheets("На списання").Range("C2:C24")
        If strA = "" Then
            Exit For
        End If
        count2 = count2 + 1
        If count2 < count1 Then
            objWord.Selection.TypeText CStr(strA) & ", "
        Else
            objWord.Selection.TypeText CStr(strA)
        End If
    Next strA
    
    With objWord.Selection
        .TypeText Text:=" для встановлення економічної доцільності списання."
        .TypeParagraph
        .TypeParagraph
        .ParagraphFormat.Alignment = 1  'Center
        .Font.Bold = True
        .TypeText Text:="Загальні відомості"
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .ParagraphFormat.Alignment = 3    'Justify
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(1.25)
    End With
    
    'Текст первой глави
    count2 = 0
    nameArray = Iteration(count1, Worksheets("На списання").Range("B2:B24")) 'Найменування основного засобу
    nameArray2 = Iteration(count1, Worksheets("На списання").Range("C2:C24")) 'Родовий Відміннок - Кого?
    sapNum = Iteration(count1, Worksheets("На списання").Range("D2:D24"))    'SAP Номер
    reason2 = Iteration(count1, Worksheets("На списання").Range("F2:F24"))   'Причина - для Тех Стану
    reason = Iteration(count1, Worksheets("На списання").Range("G2:G24"))    'Причина "відновлювального ремонту" - Чого?
    years = Iteration(count1, Worksheets("На списання").Range("J2:J24"))     'Лет обладнанню
    explDate = Iteration(count1, Worksheets("На списання").Range("K2:K24"))  'Експлуатація (дата словами)
    vartist = Iteration(count1, Worksheets("На списання").Range("N2:N24"))   'Вартість
    vaga = Iteration(count1, Worksheets("На списання").Range("O2:O24"))      'Вага
    typeBryht = Iteration(count1, Worksheets("На списання").Range("P2:P24")) 'Тип брухту
    kindBryht = Iteration(count1, Worksheets("На списання").Range("Q2:Q24")) 'Вид брухту
    kilkist = Iteration(count1, Worksheets("На списання").Range("R2:R24"))   'Кількість
    price = Iteration(count1, Worksheets("На списання").Range("S2:S24"))     'Ціна покупки брухту
    serialN = Iteration(count1, Worksheets("На списання").Range("T2:T24"))   'Серійний номер
    nameBryht = Iteration(count1, Worksheets("На списання").Range("U2:U24")) 'Назва обладнання (як брухт)
       
    For i = 1 To count1
        With objWord.Selection
            .TypeText nameArray(i) & ", інвентарний номер SAP " & CStr(sapNum(i)) & _
            " для встановлення економічної доцільності списання був введений в експлуатацію " & _
            explDate(i) & ". Через тривалий термін експлуатації понад " & CStr(years(i)) & _
            " років та інтенсивність його використання, знаходиться в неробочому стані: морально та фізично зношений та непридатний для подальшого використання. Знос становить 100%."
            .TypeParagraph
        End With
    Next i
    
    'Вторая глава
    With objWord.Selection
        .ParagraphFormat.Alignment = 1  'Center
        .Font.Bold = True
        .TypeParagraph
        .TypeText Text:="Технічна характеристика"
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .ParagraphFormat.Alignment = 3    'Justify
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(1.25)
    End With
    
    For i = 1 To count1
        With objWord.Selection
            .TypeText "Під час обстеження технічного стану " & nameArray2(i) & ", інвентарний номер " & CStr(sapNum(i)) & _
            ", для встановлення економічної доцільності списання було встановлено, що він знаходиться в неробочому стані та не підлягає проведенню відновлювального ремонту " & _
            reason(i) & ". У зв'язку з відсутністю запасних частин для даного обладнання на сьогодні у виробництві, а також в зв'язку з тим, що дане обладнання морально застаріле не відповідає сучасним технічним вимогам, подальша експлуатація недоцільна, що підтверджено актом технічного стану."
            .TypeParagraph
        End With
    Next i
    
    With objWord.Selection
        .TypeText Text:="Причина несправності подано в таблиці 1."
        .TypeParagraph
        .Font.Bold = True
        .TypeParagraph
        .TypeText Text:="Таблиця 1"
        .TypeParagraph
    End With
    
    'Третья таблица с тех характеристиками
    numOfRows = count1 + 2
    numOfColumns = 7
    pos = objDoc.Paragraphs.count   'узнаем сколько параграфов уже есть
    pos2 = objDoc.Paragraphs(pos).Range.End 'координаты конечной точки
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    
    With objDoc.Tables(3)
        .Borders.Enable = True
        .Columns(1).SetWidth objWord.Application.CentimetersToPoints(0.99), 0
        .Columns(2).SetWidth objWord.Application.CentimetersToPoints(3), 0
        .Columns(3).SetWidth objWord.Application.CentimetersToPoints(2.75), 0
        .Columns(4).SetWidth objWord.Application.CentimetersToPoints(2.25), 0
        .Columns(5).SetWidth objWord.Application.CentimetersToPoints(3.75), 0
        .Columns(6).SetWidth objWord.Application.CentimetersToPoints(2.5), 0
        .Columns(7).SetWidth objWord.Application.CentimetersToPoints(1.74), 0
    End With
    
    'Заполняем таблицу
    i = 1
    j = 1

    For Each strA In Worksheets("Тех Акт (таблиці)").Range("A2:G26")
        Select Case i
            Case Is = 2
                objDoc.Tables(3).Cell(Row:=i, Column:=j).Range.Font.Italic = True
                objDoc.Tables(3).Cell(Row:=i, Column:=j).Range.Font.Bold = False
            Case Is > 2
                objDoc.Tables(3).Cell(Row:=i, Column:=j).Range.Font.Bold = False
        End Select
                    
        With objDoc.Tables(3).Cell(Row:=i, Column:=j)
            .Range.ParagraphFormat.Alignment = 1
            .VerticalAlignment = 1
            .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
            .Range.Font.Size = 11
            .Range.InsertAfter strA
        End With

        j = j + 1   'смещяемся влево
        If j > numOfColumns Then
            j = 1
            i = i + 1 'сдел ряд, первая колона
            If i > numOfRows Then
                Exit For
            End If
        End If
    Next strA
    
    objDoc.Tables(3).Cell(Row:=numOfRows, Column:=j).Select
     
    'Глава 3 Економ показники
    With objWord.Selection
        .MoveDown Unit:=5, count:=1
        .Font.Size = 12
        .TypeParagraph
        .ParagraphFormat.Alignment = 1  'Center
        .Font.Bold = True
        .TypeText Text:="Економічні показники"
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .ParagraphFormat.Alignment = 3    'Justify
        .TypeText Text:="Комісією проведено інвентаризацію "
    End With
    
    For i = 1 To count1
        objWord.Selection.TypeText nameArray2(i)
        If count1 = 1 Then
            Exit For
        ElseIf i < count1 - 1 Then
            objWord.Selection.TypeText Text:=", "
        ElseIf i = count1 - 1 Then
            objWord.Selection.TypeText Text:=" та "
        End If
    Next i
    
    With objWord.Selection
        .TypeText Text:=" які пропонуються до списання, оформлені необхідні документи для проведення списання та підготовлено економічні показники."
        .TypeParagraph
        .TypeText "Балансову вартість та знос подано в таблиці 2 (станом на " & CStr(Worksheets("На списання").Range("AA2")) & ")."
        .TypeParagraph
        .ParagraphFormat.Alignment = 0  'Left
        .Font.Bold = True
        .TypeParagraph
        .TypeText Text:="Таблиця 2"
        .TypeParagraph
    End With
    
    'Четвертая таблица с економ показн
    numOfRows = count1 + 3
    numOfColumns = 6
    pos = objDoc.Paragraphs.count   'узнаем сколько параграфов уже есть
    pos2 = objDoc.Paragraphs(pos).Range.End 'координаты конечной точки
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    
    With objDoc.Tables(4)
        .Borders.Enable = True
        .Columns(1).SetWidth objWord.Application.CentimetersToPoints(1.64), 0
        .Columns(2).SetWidth objWord.Application.CentimetersToPoints(4.78), 0
        .Columns(3).SetWidth objWord.Application.CentimetersToPoints(3.06), 0
        .Columns(4).SetWidth objWord.Application.CentimetersToPoints(2.92), 0
        .Columns(5).SetWidth objWord.Application.CentimetersToPoints(2.64), 0
        .Columns(6).SetWidth objWord.Application.CentimetersToPoints(1.94), 0
    End With
    
    'Заполняем таблицу
    i = 1
    j = 1
    
    For Each strA In Worksheets("Тех Акт (таблиці)").Range("J2:O26")
        Select Case i
            Case Is = 2
                objDoc.Tables(4).Cell(Row:=i, Column:=j).Range.Font.Italic = True
                objDoc.Tables(4).Cell(Row:=i, Column:=j).Range.Font.Bold = False
            Case Is > 2
                objDoc.Tables(4).Cell(Row:=i, Column:=j).Range.Font.Bold = False
        End Select
                    
        With objDoc.Tables(4).Cell(Row:=i, Column:=j)
            .Range.ParagraphFormat.Alignment = 1
            .VerticalAlignment = 1
            .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
            .Range.Font.Size = 11
        End With
        
        If j >= 5 And i > 2 Then
            objDoc.Tables(4).Cell(Row:=i, Column:=j).Range.InsertAfter Format(strA, "0.00")
        Else
            objDoc.Tables(4).Cell(Row:=i, Column:=j).Range.InsertAfter strA
        End If

        j = j + 1   'смещяемся влево
        If j > numOfColumns Then
            j = 1
            i = i + 1 'сдел ряд, первая колона
            If i = numOfRows Then
                For Each strB In Worksheets("Тех Акт (таблиці)").Range("J27:O27")
                    With objDoc.Tables(4).Cell(Row:=i, Column:=j)
                        .Range.ParagraphFormat.Alignment = 1
                        .VerticalAlignment = 1
                        .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
                        .Range.InsertAfter strB
                    End With
                    j = j + 1
                Next strB
                Exit For
            End If
        End If
    Next strA
    
    objDoc.Tables(4).Cell(Row:=numOfRows, Column:=j - 1).Select
    
    With objWord.Selection
        .MoveDown Unit:=5, count:=1
        .Font.Bold = False
        .TypeParagraph
        .TypeText Text:="Оскільки зазначені "
    End With
    For i = 1 To count1
        objWord.Selection.TypeText nameArray2(i)
        If count1 = 1 Then
            Exit For
        ElseIf i < count1 - 1 Then
            objWord.Selection.TypeText Text:=", "
        ElseIf i = count1 - 1 Then
            objWord.Selection.TypeText Text:=" та "
        End If
    Next i
    
    With objWord.Selection
        .TypeText Text:=" не працюють, а деталі та інструменти не виготовляють, розрахунок вартості ремонту даного майна провести неможливо."
        .TypeParagraph
        .ParagraphFormat.Alignment = 1  'Center
        .Font.Bold = True
        .TypeParagraph
        .TypeText Text:="Інформація про очікуваний фінансовий результат списання майна, а також напрями використання коштів, які передбачається одержати в результаті списання"
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .ParagraphFormat.Alignment = 3    'Justify
        .TypeText Text:="Складові частини, які приведуть до отримання в майбутньому економічної вигоди та які можливо використати в господарській діяльності підприємства - відсутні. Запасні частини, придатні до використання - відсутні. Дорогоцінні метали - відсутні."
        .TypeParagraph
    End With
    
    For i = 1 To count1
        If typeBryht(i) = "брухту" Then
            strB = ", буде оприбутковано " & vaga(i) & " кг брухту " & kindBryht(i) & " за ціною " & price(i) & _
            " грн., який потребуватиме оприбуткуванню на балансовий рахунок Товариства." & _
            " Реалізація вторинної сировини відбудеться відповідно вимог чинного законодавства та нормативних документів ТОВ ""Оператор ГТС України"". Очікуваний фінансовий результат, орієнтовно, складе " & _
            price(i) * vaga(i) & " грн."
        Else
            strB = ", буде утворений утиль " & kindBryht(i) & " в кількості " & kilkist(i) & " шт. вагою " & vaga(i) & _
            " кг, який потребуватиме оприбуткуванню на позабалансовий рахунок Товариства з подальшою передачею на утилізацію спеціалізованому підприємству відходів після списання." & ". Операції поводження з відходами відбудуться відповідно вимог чинного законодавства та нормативних документів ТОВ ""Оператор ГТС України""."
        End If
        With objWord.Selection
            .TypeText "У разі надання дозволу на списання " & nameArray2(i)
            .TypeText ", інвентарний номер SAP " & sapNum(i) & strB
            .TypeParagraph
        End With
    Next i
        
    objWord.Selection.TypeText "Детальна інформація по оприбуткуванню наведена в таблиці:"
    objWord.Selection.TypeParagraph
    
    'П'ята таблиця
    numOfRows = count1 + 3
    numOfColumns = 7
    pos = objDoc.Paragraphs.count   'узнаем сколько параграфов уже есть
    pos2 = objDoc.Paragraphs(pos).Range.End 'координаты конечной точки
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    
    With objDoc.Tables(5)
        .Borders.Enable = True
        .Columns(1).SetWidth objWord.Application.CentimetersToPoints(0.99), 0
        .Columns(2).SetWidth objWord.Application.CentimetersToPoints(4.5), 0
        .Columns(3).SetWidth objWord.Application.CentimetersToPoints(3.38), 0
        .Columns(4).SetWidth objWord.Application.CentimetersToPoints(2.12), 0
        .Columns(5).SetWidth objWord.Application.CentimetersToPoints(1.75), 0
        .Columns(6).SetWidth objWord.Application.CentimetersToPoints(1.5), 0
        .Columns(7).SetWidth objWord.Application.CentimetersToPoints(2.74), 0
    End With
    
    'Заполняем таблицу
    i = 1
    j = 1
    
    For Each strA In Worksheets("Тех Акт (таблиці)").Range("AA2:AG26")
        Select Case i
            Case Is = 1
                objDoc.Tables(5).Cell(Row:=i, Column:=j).Range.Font.Bold = True
            Case Is = 2
                objDoc.Tables(5).Cell(Row:=i, Column:=j).Range.Font.Italic = True
                objDoc.Tables(5).Cell(Row:=i, Column:=j).Range.Font.Bold = False
            Case Is > 2
                objDoc.Tables(5).Cell(Row:=i, Column:=j).Range.Font.Bold = False
        End Select
                    
        With objDoc.Tables(5).Cell(Row:=i, Column:=j)
            .Range.ParagraphFormat.Alignment = 1
            .VerticalAlignment = 1
            .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
            .Range.Font.Size = 11
        End With
        
        If j >= 5 And i > 2 Then
            objDoc.Tables(5).Cell(Row:=i, Column:=j).Range.InsertAfter Format(strA, "0.00")
        Else
            objDoc.Tables(5).Cell(Row:=i, Column:=j).Range.InsertAfter strA
        End If

        j = j + 1   'смещяемся влево
        If j > numOfColumns Then
            j = 1
            i = i + 1 'сдел ряд, первая колона
            If i = numOfRows Then
                For Each strB In Worksheets("Тех Акт (таблиці)").Range("AA27:AG27")
                    With objDoc.Tables(5).Cell(Row:=i, Column:=j)
                        .Range.Font.Bold = True
                        .Range.ParagraphFormat.Alignment = 1
                        .VerticalAlignment = 1
                        .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
                        .Range.InsertAfter strB
                    End With
                    j = j + 1
                Next strB
                Exit For
            End If
        End If
    Next strA
    
    objDoc.Tables(5).Cell(Row:=numOfRows, Column:=j - 1).Select
    
    With objWord.Selection
        .MoveDown Unit:=5, count:=1
        .Font.Bold = True
        .Font.Size = 12
        .TypeParagraph
        .ParagraphFormat.Alignment = 1  'Center
        .TypeText "Висновок необхідності списання майна"
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .ParagraphFormat.Alignment = 3    'Justify
        .TypeText "Враховуючи той факт, що майно не використовується у виробничій діяльності, так як знаходиться в неробочому стані та не підлягає подальшій експлуатації, Комісія Миколаївського ЛВУМГ прийняла попереднє рішення про надання згоди на погодження списання вищевказаного власного майна з балансу ТОВ ""Оператор ГТС України"", як такого що не відповідає критеріям активу, а саме:"
        .TypeParagraph
    End With
    
    'Шоста таблиця
    numOfRows = count1 + 3
    numOfColumns = 7
    pos = objDoc.Paragraphs.count   'узнаем сколько параграфов уже есть
    pos2 = objDoc.Paragraphs(pos).Range.End 'координаты конечной точки
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    
    With objDoc.Tables(6)
        .Borders.Enable = True
        .Columns(1).SetWidth objWord.Application.CentimetersToPoints(0.99), 0
        .Columns(2).SetWidth objWord.Application.CentimetersToPoints(4.76), 0
        .Columns(3).SetWidth objWord.Application.CentimetersToPoints(3), 0
        .Columns(4).SetWidth objWord.Application.CentimetersToPoints(2), 0
        .Columns(5).SetWidth objWord.Application.CentimetersToPoints(2), 0
        .Columns(6).SetWidth objWord.Application.CentimetersToPoints(1.83), 0
        .Columns(7).SetWidth objWord.Application.CentimetersToPoints(2.41), 0
    End With
    
    'Заполняем таблицу
    i = 1
    j = 1
    
    For Each strA In Worksheets("Тех Акт (таблиці)").Range("R2:X26")
        Select Case i
            Case Is = 1
                objDoc.Tables(6).Cell(Row:=i, Column:=j).Range.Font.Bold = True
            Case Is = 2
                objDoc.Tables(6).Cell(Row:=i, Column:=j).Range.Font.Italic = True
                objDoc.Tables(6).Cell(Row:=i, Column:=j).Range.Font.Bold = False
            Case Is > 2
                objDoc.Tables(6).Cell(Row:=i, Column:=j).Range.Font.Bold = False
        End Select
                    
        With objDoc.Tables(6).Cell(Row:=i, Column:=j)
            .Range.ParagraphFormat.Alignment = 1
            .VerticalAlignment = 1
            .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
            .Range.Font.Size = 11
        End With
        
        If j >= 4 And i > 2 Then
            objDoc.Tables(6).Cell(Row:=i, Column:=j).Range.InsertAfter Format(strA, "0.00")
        Else
            objDoc.Tables(6).Cell(Row:=i, Column:=j).Range.InsertAfter strA
        End If
        
        j = j + 1   'смещяемся влево
        If j > numOfColumns Then
            j = 1
            i = i + 1 'сдел ряд, первая колона
            If i = numOfRows Then
                For Each strB In Worksheets("Тех Акт (таблиці)").Range("R27:X27")
                    With objDoc.Tables(6).Cell(Row:=i, Column:=j)
                        .Range.Font.Bold = True
                        .Range.ParagraphFormat.Alignment = 1
                        .VerticalAlignment = 1
                        .Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
                        .Range.InsertAfter strB
                    End With
                    j = j + 1
                Next strB
                Exit For
            End If
        End If
    Next strA
    
    objDoc.Tables(6).Cell(Row:=numOfRows, Column:=j - 1).Select
    
    With objWord.Selection
        .MoveDown Unit:=5, count:=1
        .Font.Size = 12
        .TypeParagraph
    End With
    
    'Сьома таблиця
    numOfRows = 9
    numOfColumns = 3
    pos = objDoc.Paragraphs.count   'узнаем сколько параграфов уже есть
    pos2 = objDoc.Paragraphs(pos).Range.End 'координаты конечной точки
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    
    With objDoc.Tables(7)
        .Borders.Enable = False
        .Columns(1).SetWidth objWord.Application.CentimetersToPoints(3.6), 0
        .Columns(2).SetWidth objWord.Application.CentimetersToPoints(7.8), 0
        .Columns(3).SetWidth objWord.Application.CentimetersToPoints(5.6), 0
    End With
    
    'Заполняем таблицу
    i = 1
    j = 1

    For Each strA In Worksheets("Тех Акт (таблиці)").Range("AO5:AQ13")
        If j = 2 Then
            objDoc.Tables(7).Cell(Row:=i, Column:=j).Range.Borders(-3).LineStyle = 1 'wdBottom / wdLineSttleSingle
        End If
        objDoc.Tables(7).Rows(i).SetHeight objWord.Application.CentimetersToPoints(0.8), 0
        objDoc.Tables(7).Rows(i).Cells.VerticalAlignment = 1 'wdAlignVerticalCenter
        objDoc.Tables(7).Cell(Row:=i, Column:=j).Range.ParagraphFormat.SpaceAfter = 0 'удалить интервал после абзаца
        objDoc.Tables(7).Cell(Row:=i, Column:=j).Range.InsertAfter strA
        j = j + 1   'смещяемся влево
        If j > numOfColumns Then
            j = 1
            i = i + 1 'сдел ряд, первая колона
            If i > numOfRows Then
                i = 1
            End If
        End If
    Next strA
    
    'Нумерация страниц
    Set myPgNm = objDoc.Sections(1).Footers(1).PageNumbers.Add(PageNumberAlignment:=1, FirstPage:=True)
    myPgNm.Select
    With objWord.Selection.Range
        .Font.Name = "Times New Roman"
        .Font.Size = 12
    End With
    
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Select
    
'    MsgBox intReportCount & " заметки созданои сохранено в папке " & ThisWorkbook.Path
End Sub
