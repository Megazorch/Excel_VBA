Attribute VB_Name = "MakeReport"
Option Explicit

Sub MakeReport()
    '���������� ������ �� ����������
    'macros by Ivakhiv Roman - megazorch@gmail.com
    Dim fName() As String
    Dim FullN() As String
    Dim place() As String
    Dim short() As String
    Dim separate_calc() As Boolean  '16.09.2022 - #56
    Dim Count As Integer
    Dim count2 As Integer
    Dim strA As Variant
    Dim i As Integer
    Dim StartTime As Date, EndTime As Date
    Dim totTime As Variant
    Dim minutes As Integer
    Dim seconds As Double
    Dim result As String
    Dim savePTH As String           '17.08.2022 - #20
    Dim objWord As Object           '17.08.2022 - #20
    
    StartTime = Timer
    Count = 0
    For Each strA In Range("P.I.B.")
        If strA = Empty Then
            Exit For
        End If
        Count = Count + 1
    Next strA
    
    ReDim fName(Count)  '��������� ��� ��� ��� ����� ����������
    ReDim FullN(Count)  '19.09.2022
    ReDim place(Count)
    ReDim transport(Count)
    ReDim short(Count)
    ReDim separate_calc(Count)  '16.09.2022 - #56
    
    count2 = 0
    For Each strA In Range("full_name") '17.08.2022 - #18: 16.09.2022 - #46
        If count2 < Count Then
            count2 = count2 + 1
            fName(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("P.I.B.") '19.09.2022
        If count2 < Count Then
            count2 = count2 + 1
            FullN(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("place")
        If count2 < Count Then
            count2 = count2 + 1
            place(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("short_name")   '16.09.2022 - #48
        If count2 < Count Then
            count2 = count2 + 1
            short(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("sep_calc")   '16.09.2022 - #56
        If count2 < Count Then
            count2 = count2 + 1
            separate_calc(count2) = strA
        Else
            Exit For
        End If
    Next strA
           
    For i = 1 To count2
        Call Report(fName(i), FullN(i), place(i), short(i), separate_calc(i))
    Next i
    
    EndTime = Timer
    totTime = Format(EndTime - StartTime, "0.0")
    If totTime >= 60 Then
        minutes = totTime \ 60
        seconds = Format((totTime / 60 - minutes) * 60, "0")
        result = CStr(minutes) & " ��. " & CStr(seconds) & " ���."
    Else
        result = CStr(totTime) & " ���."
    End If
    
    
    
    If Count = 1 Then   '17.08.2022 #20
        savePTH = ActiveWorkbook.pATH & "\���������� ������ �� ������. - " & short(i - 1) & ".docx"
        MsgBox "��� ������ ��������� � ���������� � �����: " & vbCrLf & ActiveWorkbook.pATH & "," & vbCrLf & "���� �� " _
        & result, vbInformation, "������!"
            If MsgBox("³������ ����?", vbYesNo, "������!") = vbYes Then
                If Dir(savePTH) <> "" Then
                    Set objWord = CreateObject("Word.Application")
                    objWord.Visible = True
                    objWord.Documents.Open savePTH
                    MsgBox "���� �������. �������� �� ������ Word �� ����� �����.", vbInformation, "���� �������."
                Else
                    MsgBox "���� ������ �_� ?!", vbCritical, "���!"
                End If
            Else
'                objWord.Quit
                Set objWord = Nothing
            End If
    Else
        MsgBox Count & " ���� ������ ��������� � ���������� � �����: " & ActiveWorkbook.pATH & "," & vbCrLf _
        & "���� �� " & result, vbInformation, "������!"
    End If
End Sub
