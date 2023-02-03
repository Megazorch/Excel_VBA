Attribute VB_Name = "MakeReport_2"
Option Explicit

Sub MakeReport_2()
    '��� ��� ��������� �������� �� ���������� �� �����"
    'macros by Ivakhiv Roman - megazorch@gmail.com
    Dim fName() As String '������ ���
    Dim place() As String '��������� �����
    Dim short() As String '��������� ��'� ��� ������
    Dim purpose() As String '���� ����������
    Dim car() As String '����� �������
    Dim garage() As String '�����������
    Dim days As String  ' 17.08.2022 - ��� ���������� ����������� ����� "���" � ���
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
    
    ReDim fName(Count)  '��������� ��� ��� ��� ����� ����������
    ReDim place(Count)
    ReDim short(Count)
    ReDim purpose(Count)
    ReDim car(Count)
    ReDim garage(Count)
    
    count2 = 0
    For Each strA In Range("P.I.B.") '������ ���
        If count2 < Count Then
            count2 = count2 + 1
            fName(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("place") '��������� �����
        If count2 < Count Then
            count2 = count2 + 1
            place(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("short_name") '��������� ��'� ��� ������
        If count2 < Count Then
            count2 = count2 + 1
            short(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("purpose") '���� ����������
        If count2 < Count Then
            count2 = count2 + 1
            purpose(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("transport") '����� �������
        If count2 < Count Then
            count2 = count2 + 1
            car(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    count2 = 0
    
    For Each strA In Range("garage") '�����������
        If count2 < Count Then
            count2 = count2 + 1
            garage(count2) = CStr(strA)
        Else
            Exit For
        End If
    Next strA
    
    Select Case Range("dob_days")   ' 17.08.2022 - ���������
        Case 1:
            days = " ���� - � "
        Case 2, 3, 4:
            days = " �� - � "
        Case Is >= 5:
            days = " ��� - � "
    End Select
    
    For i = 1 To count2
        Call Report_2(fName(i), place(i), short(i), purpose(i), car(i), garage(i), days) '17.08.2022 - added "days"
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
    
    If Count = 1 Then   '17.08.2022 #22
        savePTH = ActiveWorkbook.pATH & "\��� ��� ��������� �������� - " & short(i - 1) & ".docx"
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
                objWord.Quit
                Set objWord = Nothing
            End If
    Else
        MsgBox Count & " ���� ������ ��������� � ���������� � �����: " & ActiveWorkbook.pATH & "," & vbCrLf _
        & "���� �� " & result, vbInformation, "������!"
    End If
End Sub

