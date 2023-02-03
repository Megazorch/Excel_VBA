Attribute VB_Name = "MakeReport_2"
Option Explicit

Sub MakeReport_2()
    'Звіт про виконання завдання по відрядженню по Україні"
    'macros by Ivakhiv Roman - megazorch@gmail.com
    Dim fName() As String 'полное имя
    Dim place() As String 'населённый пункт
    Dim short() As String 'скорочене ім'я для підпису
    Dim purpose() As String 'ціль відрядження
    Dim car() As String 'проїзд машиною
    Dim garage() As String 'гаражування
    Dim days As String  ' 17.08.2022 - для коректного відображання слова "днів" в звіті
    Dim Count As Integer
    Dim count2 As Integer
    Dim strA As Variant
    Dim i As Integer
    Dim StartTime As Date, EndTime As Date
    Dim totTime As Variant
    Dim minutes As Integer
    Dim seconds As Double
    Dim result As String
    Dim savePTH As String   '17.08.2022 #22
    Dim objWord As Object   '17.08.2022 #22
    
    StartTime = Timer
    Count = 0
    For Each strA In Range("P.I.B.")
        If strA = Empty Then
            Exit For
        End If
        Count = Count + 1
    Next strA
    
    ReDim fName(Count)  'расширяем так как уже знаем количество
    ReDim place(Count)
    ReDim short(Count)
    ReDim purpose(Count)
    ReDim car(Count)
    ReDim garage(Count)
    
    count2 = 0
    For Each strA In Range("P.I.B.") 'полное имя
        If count2 < Count Then
            count2 = count2 + 1
            fName(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("place") 'населённый пункт
        If count2 < Count Then
            count2 = count2 + 1
            place(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("short_name") 'скорочене ім'я для підпису
        If count2 < Count Then
            count2 = count2 + 1
            short(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("purpose") 'ціль відрядження
        If count2 < Count Then
            count2 = count2 + 1
            purpose(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("transport") 'проїзд машиною
        If count2 < Count Then
            count2 = count2 + 1
            car(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("garage") 'гаражування
        If count2 < Count Then
            count2 = count2 + 1
            garage(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    
    Select Case Range("dob_days")   ' 17.08.2022 - створенно
        Case 1:
            days = " день - з "
        Case 2, 3, 4:
            days = " дні - з "
        Case Is >= 5:
            days = " днів - з "
    End Select
    
    For i = 1 To count2
        Call Report_2(fName(i), place(i), short(i), purpose(i), car(i), garage(i), days) '17.08.2022 - added "days"
    Next i
    
    EndTime = Timer
    totTime = Format(EndTime - StartTime, "0.0")
    If totTime >= 60 Then
        minutes = totTime \ 60
        seconds = Format((totTime / 60 - minutes) * 60, "0")
        result = CStr(minutes) & " хв. " & CStr(seconds) & " сек."
    Else
        result = CStr(totTime) & " сек."
    End If
    
    If Count = 1 Then   '17.08.2022 #22
        savePTH = ActiveWorkbook.pATH & "\Звіт про виконання завдання - " & short(i - 1) & ".docx"
        MsgBox "Звіт успішно стоворено і збереженно в папці: " & vbCrLf & ActiveWorkbook.pATH & "," & vbCrLf & "лише за " _
        & result, vbInformation, "Готово!"
            If MsgBox("Відкрити файл?", vbYesNo, "Готово!") = vbYes Then
                If Dir(savePTH) <> "" Then
                    Set objWord = CreateObject("Word.Application")
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
    Else
        MsgBox Count & " звітів успішно стоворено і збереженно в папці: " & ActiveWorkbook.pATH & "," & vbCrLf _
        & "лише за " & result, vbInformation, "Готово!"
    End If
End Sub

