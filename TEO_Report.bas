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
    Dim nameArray() As Variant   'B - ������������ ��������� ������
    Dim nameArray2() As Variant  'C - ������� ³������ - ����?
    Dim sapNum() As Variant      'D - SAP �����
    Dim reason2() As Variant     'F - ������� - ��� ��� �����
    Dim reason() As Variant      'G - ������� "��������������� �������" - ����?
    Dim years() As Variant       'J - ��� ����������
    Dim explDate() As Variant    'K - ������������ (���� �������)
    Dim vartist() As Variant     'N - �������
    Dim vaga() As Variant        'O - ����
    Dim typeBryht() As Variant   'P - ��� ������
    Dim kindBryht() As Variant   'Q - ��� ������
    Dim kilkist() As Variant     'R - ʳ������
    Dim price() As Variant       'S - ֳ�� ������� ������
    Dim serialN() As Variant     'T - ������� �����
    Dim nameBryht() As Variant   'U - ����� ���������� (�� �����)
    
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Add
    objWord.Visible = True
    Set objRange = objDoc.Range
    numOfRows = 1
    numOfColumns = 2
    
    '����
    With objDoc.PageSetup
        .LeftMargin = 85.0394
        .RightMargin = 28.34646
        .TopMargin = 56.69291
        .BottomMargin = 56.69291
    End With
        
    '�����
    objDoc.Tables.Add objRange, numOfRows, numOfColumns
    Set objTable = objDoc.Tables(1)
    objTable.Borders.Enable = False
    objTable.Cell(1, 2).Range.Select
    objTable.Cell(1, 2).Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
    objTable.Cell(1, 2).SetWidth 212.5984, 0         'wdAdjustNone = 0
    With objWord.Selection
        .ParagraphFormat.LineSpacing = 18 '�������� �������� 1,5
        .Font.Size = 12
        .Font.Name = "Times New Roman"
        .TypeText Text:="����������"
        .TypeParagraph
        .TypeText Text:="��������� ������������� �����"
        .TypeParagraph
        .TypeText Text:="_______________"
        .Font.Bold = True
        .TypeText Text:="����� ��������"
        .Font.Bold = False
        .TypeParagraph
        .TypeText Text:="""__""_________2021 �."
        .MoveDown Unit:=5, count:=2           '����������
        .ParagraphFormat.SpaceAfter = 0     '������� �������� ����� ������
        .Font.Size = 12
        .Font.Name = "Times New Roman"
        .TypeParagraph
        .ParagraphFormat.Alignment = 1  'Center
        .Font.Bold = True
        .TypeText Text:="���Ͳ��-�����̲��� ������������� ��ֲ�����Ҳ �������� ����� ����������"
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .ParagraphFormat.Alignment = 3    'Justify
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(1.25)
        .TypeText Text:="������� ���� ����� � ������ �������� ���������� �� �������� �����, ��� ����������� �� ������ ����������, �������� ������� ������������� ����� ��� ""�������� ��� ������"" �� 15 ����� 2021 �. �417 � �����:"
        .TypeParagraph
        .TypeParagraph
    End With
        
    '������ ������� � ��������
    numOfRows = 9
    numOfColumns = 3
    pos = objDoc.Paragraphs.count   '������ ������� ���������� ��� ����
    pos2 = objDoc.Paragraphs(pos).Range.End '���������� �������� �����
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    
    With objDoc.Tables(2)
        .Borders.Enable = False
        .Columns(1).SetWidth 99.2126, 0         '3.5cm
        .Columns(2).SetWidth 233.8583, 0        '8.25cm
        .Columns(3).SetWidth 148.8189, 0        '5.25cm
        .Columns(3).Cells.VerticalAlignment = 1 'wdAlignVerticalCenter
    End With
    
    '��������� �������
    i = 1
    j = 1

    For Each strA In Worksheets("��� ��� (�������)").Range("AJ5:AL13")
        objDoc.Tables(2).Cell(Row:=i, Column:=j).Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
        objDoc.Tables(2).Cell(Row:=i, Column:=j).Range.InsertAfter strA
        j = j + 1   '��������� �����
        If j > numOfColumns Then
            j = 1
            i = i + 1 '���� ���, ������ ������
            If i > numOfRows Then
                i = 1
            End If
        End If
    Next strA
    
    objWord.Selection.MoveDown Unit:=5, count:=9
    
    '����������� ������������ �����
    With objWord.Selection
        .TypeParagraph
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(0)
        .TypeText Text:="������� ����������� �������� �� ���������� ��������� ��'����: "
    End With
    
    count1 = 0   '������ ������� ��������� ��� �������
    For Each strA In Worksheets("�� ��������").Range("C2:C24")
        If strA = Empty Then
            Exit For
        End If
        count1 = count1 + 1
    Next strA
    
    count2 = 0
    For Each strA In Worksheets("�� ��������").Range("C2:C24")
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
        .TypeText Text:=" ��� ������������ ��������� ���������� ��������."
        .TypeParagraph
        .TypeParagraph
        .ParagraphFormat.Alignment = 1  'Center
        .Font.Bold = True
        .TypeText Text:="������� �������"
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .ParagraphFormat.Alignment = 3    'Justify
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(1.25)
    End With
    
    '����� ������ �����
    count2 = 0
    nameArray = Iteration(count1, Worksheets("�� ��������").Range("B2:B24")) '������������ ��������� ������
    nameArray2 = Iteration(count1, Worksheets("�� ��������").Range("C2:C24")) '������� ³������ - ����?
    sapNum = Iteration(count1, Worksheets("�� ��������").Range("D2:D24"))    'SAP �����
    reason2 = Iteration(count1, Worksheets("�� ��������").Range("F2:F24"))   '������� - ��� ��� �����
    reason = Iteration(count1, Worksheets("�� ��������").Range("G2:G24"))    '������� "��������������� �������" - ����?
    years = Iteration(count1, Worksheets("�� ��������").Range("J2:J24"))     '��� ����������
    explDate = Iteration(count1, Worksheets("�� ��������").Range("K2:K24"))  '������������ (���� �������)
    vartist = Iteration(count1, Worksheets("�� ��������").Range("N2:N24"))   '�������
    vaga = Iteration(count1, Worksheets("�� ��������").Range("O2:O24"))      '����
    typeBryht = Iteration(count1, Worksheets("�� ��������").Range("P2:P24")) '��� ������
    kindBryht = Iteration(count1, Worksheets("�� ��������").Range("Q2:Q24")) '��� ������
    kilkist = Iteration(count1, Worksheets("�� ��������").Range("R2:R24"))   'ʳ������
    price = Iteration(count1, Worksheets("�� ��������").Range("S2:S24"))     'ֳ�� ������� ������
    serialN = Iteration(count1, Worksheets("�� ��������").Range("T2:T24"))   '������� �����
    nameBryht = Iteration(count1, Worksheets("�� ��������").Range("U2:U24")) '����� ���������� (�� �����)
       
    For i = 1 To count1
        With objWord.Selection
            .TypeText nameArray(i) & ", ����������� ����� SAP " & CStr(sapNum(i)) & _
            " ��� ������������ ��������� ���������� �������� ��� �������� � ������������ " & _
            explDate(i) & ". ����� �������� ����� ������������ ����� " & CStr(years(i)) & _
            " ���� �� ������������ ���� ������������, ����������� � ���������� ����: �������� �� ������� �������� �� ����������� ��� ���������� ������������. ���� ��������� 100%."
            .TypeParagraph
        End With
    Next i
    
    '������ �����
    With objWord.Selection
        .ParagraphFormat.Alignment = 1  'Center
        .Font.Bold = True
        .TypeParagraph
        .TypeText Text:="������� ��������������"
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .ParagraphFormat.Alignment = 3    'Justify
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(1.25)
    End With
    
    For i = 1 To count1
        With objWord.Selection
            .TypeText "ϳ� ��� ���������� ��������� ����� " & nameArray2(i) & ", ����������� ����� " & CStr(sapNum(i)) & _
            ", ��� ������������ ��������� ���������� �������� ���� �����������, �� �� ����������� � ���������� ���� �� �� ������ ���������� ��������������� ������� " & _
            reason(i) & ". � ��'���� � ��������� �������� ������ ��� ������ ���������� �� ������� � ����������, � ����� � ��'���� � ���, �� ���� ���������� �������� �������� �� ������� �������� �������� �������, �������� ������������ ����������, �� ����������� ����� ��������� �����."
            .TypeParagraph
        End With
    Next i
    
    With objWord.Selection
        .TypeText Text:="������� ����������� ������ � ������� 1."
        .TypeParagraph
        .Font.Bold = True
        .TypeParagraph
        .TypeText Text:="������� 1"
        .TypeParagraph
    End With
    
    '������ ������� � ��� ����������������
    numOfRows = count1 + 2
    numOfColumns = 7
    pos = objDoc.Paragraphs.count   '������ ������� ���������� ��� ����
    pos2 = objDoc.Paragraphs(pos).Range.End '���������� �������� �����
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
    
    '��������� �������
    i = 1
    j = 1

    For Each strA In Worksheets("��� ��� (�������)").Range("A2:G26")
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
            .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
            .Range.Font.Size = 11
            .Range.InsertAfter strA
        End With

        j = j + 1   '��������� �����
        If j > numOfColumns Then
            j = 1
            i = i + 1 '���� ���, ������ ������
            If i > numOfRows Then
                Exit For
            End If
        End If
    Next strA
    
    objDoc.Tables(3).Cell(Row:=numOfRows, Column:=j).Select
     
    '����� 3 ������ ���������
    With objWord.Selection
        .MoveDown Unit:=5, count:=1
        .Font.Size = 12
        .TypeParagraph
        .ParagraphFormat.Alignment = 1  'Center
        .Font.Bold = True
        .TypeText Text:="�������� ���������"
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .ParagraphFormat.Alignment = 3    'Justify
        .TypeText Text:="����� ��������� �������������� "
    End With
    
    For i = 1 To count1
        objWord.Selection.TypeText nameArray2(i)
        If count1 = 1 Then
            Exit For
        ElseIf i < count1 - 1 Then
            objWord.Selection.TypeText Text:=", "
        ElseIf i = count1 - 1 Then
            objWord.Selection.TypeText Text:=" �� "
        End If
    Next i
    
    With objWord.Selection
        .TypeText Text:=" �� ������������ �� ��������, �������� �������� ��������� ��� ���������� �������� �� ����������� �������� ���������."
        .TypeParagraph
        .TypeText "��������� ������� �� ���� ������ � ������� 2 (������ �� " & CStr(Worksheets("�� ��������").Range("AA2")) & ")."
        .TypeParagraph
        .ParagraphFormat.Alignment = 0  'Left
        .Font.Bold = True
        .TypeParagraph
        .TypeText Text:="������� 2"
        .TypeParagraph
    End With
    
    '��������� ������� � ������ ������
    numOfRows = count1 + 3
    numOfColumns = 6
    pos = objDoc.Paragraphs.count   '������ ������� ���������� ��� ����
    pos2 = objDoc.Paragraphs(pos).Range.End '���������� �������� �����
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
    
    '��������� �������
    i = 1
    j = 1
    
    For Each strA In Worksheets("��� ��� (�������)").Range("J2:O26")
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
            .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
            .Range.Font.Size = 11
        End With
        
        If j >= 5 And i > 2 Then
            objDoc.Tables(4).Cell(Row:=i, Column:=j).Range.InsertAfter Format(strA, "0.00")
        Else
            objDoc.Tables(4).Cell(Row:=i, Column:=j).Range.InsertAfter strA
        End If

        j = j + 1   '��������� �����
        If j > numOfColumns Then
            j = 1
            i = i + 1 '���� ���, ������ ������
            If i = numOfRows Then
                For Each strB In Worksheets("��� ��� (�������)").Range("J27:O27")
                    With objDoc.Tables(4).Cell(Row:=i, Column:=j)
                        .Range.ParagraphFormat.Alignment = 1
                        .VerticalAlignment = 1
                        .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
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
        .TypeText Text:="������� �������� "
    End With
    For i = 1 To count1
        objWord.Selection.TypeText nameArray2(i)
        If count1 = 1 Then
            Exit For
        ElseIf i < count1 - 1 Then
            objWord.Selection.TypeText Text:=", "
        ElseIf i = count1 - 1 Then
            objWord.Selection.TypeText Text:=" �� "
        End If
    Next i
    
    With objWord.Selection
        .TypeText Text:=" �� ��������, � ����� �� ����������� �� ������������, ���������� ������� ������� ������ ����� �������� ���������."
        .TypeParagraph
        .ParagraphFormat.Alignment = 1  'Center
        .Font.Bold = True
        .TypeParagraph
        .TypeText Text:="���������� ��� ���������� ���������� ��������� �������� �����, � ����� ������� ������������ �����, �� ������������� �������� � ��������� ��������"
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .ParagraphFormat.Alignment = 3    'Justify
        .TypeText Text:="������� �������, �� ��������� �� ��������� � ����������� ��������� ������ �� �� ������� ����������� � ������������ �������� ���������� - ������. ������ �������, ������� �� ������������ - ������. ���������� ������ - ������."
        .TypeParagraph
    End With
    
    For i = 1 To count1
        If typeBryht(i) = "������" Then
            strB = ", ���� ������������� " & vaga(i) & " �� ������ " & kindBryht(i) & " �� ����� " & price(i) & _
            " ���., ���� ������������� �������������� �� ���������� ������� ����������." & _
            " ��������� �������� �������� ���������� �������� ����� ������� ������������� �� ����������� ��������� ��� ""�������� ��� ������"". ���������� ���������� ���������, ��������, ������ " & _
            price(i) * vaga(i) & " ���."
        Else
            strB = ", ���� ��������� ����� " & kindBryht(i) & " � ������� " & kilkist(i) & " ��. ����� " & vaga(i) & _
            " ��, ���� ������������� �������������� �� �������������� ������� ���������� � ��������� ��������� �� ��������� ��������������� ���������� ������ ���� ��������." & ". �������� ���������� � �������� ���������� �������� ����� ������� ������������� �� ����������� ��������� ��� ""�������� ��� ������""."
        End If
        With objWord.Selection
            .TypeText "� ��� ������� ������� �� �������� " & nameArray2(i)
            .TypeText ", ����������� ����� SAP " & sapNum(i) & strB
            .TypeParagraph
        End With
    Next i
        
    objWord.Selection.TypeText "�������� ���������� �� �������������� �������� � �������:"
    objWord.Selection.TypeParagraph
    
    '�'��� �������
    numOfRows = count1 + 3
    numOfColumns = 7
    pos = objDoc.Paragraphs.count   '������ ������� ���������� ��� ����
    pos2 = objDoc.Paragraphs(pos).Range.End '���������� �������� �����
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
    
    '��������� �������
    i = 1
    j = 1
    
    For Each strA In Worksheets("��� ��� (�������)").Range("AA2:AG26")
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
            .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
            .Range.Font.Size = 11
        End With
        
        If j >= 5 And i > 2 Then
            objDoc.Tables(5).Cell(Row:=i, Column:=j).Range.InsertAfter Format(strA, "0.00")
        Else
            objDoc.Tables(5).Cell(Row:=i, Column:=j).Range.InsertAfter strA
        End If

        j = j + 1   '��������� �����
        If j > numOfColumns Then
            j = 1
            i = i + 1 '���� ���, ������ ������
            If i = numOfRows Then
                For Each strB In Worksheets("��� ��� (�������)").Range("AA27:AG27")
                    With objDoc.Tables(5).Cell(Row:=i, Column:=j)
                        .Range.Font.Bold = True
                        .Range.ParagraphFormat.Alignment = 1
                        .VerticalAlignment = 1
                        .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
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
        .TypeText "�������� ����������� �������� �����"
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = False
        .ParagraphFormat.Alignment = 3    'Justify
        .TypeText "���������� ��� ����, �� ����� �� ��������������� � ���������� ��������, ��� �� ����������� � ���������� ���� �� �� ������ ��������� ������������, ����� ������������� ����� �������� �������� ������ ��� ������� ����� �� ���������� �������� ������������� �������� ����� � ������� ��� ""�������� ��� ������"", �� ������ �� �� ������� �������� ������, � ����:"
        .TypeParagraph
    End With
    
    '����� �������
    numOfRows = count1 + 3
    numOfColumns = 7
    pos = objDoc.Paragraphs.count   '������ ������� ���������� ��� ����
    pos2 = objDoc.Paragraphs(pos).Range.End '���������� �������� �����
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
    
    '��������� �������
    i = 1
    j = 1
    
    For Each strA In Worksheets("��� ��� (�������)").Range("R2:X26")
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
            .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
            .Range.Font.Size = 11
        End With
        
        If j >= 4 And i > 2 Then
            objDoc.Tables(6).Cell(Row:=i, Column:=j).Range.InsertAfter Format(strA, "0.00")
        Else
            objDoc.Tables(6).Cell(Row:=i, Column:=j).Range.InsertAfter strA
        End If
        
        j = j + 1   '��������� �����
        If j > numOfColumns Then
            j = 1
            i = i + 1 '���� ���, ������ ������
            If i = numOfRows Then
                For Each strB In Worksheets("��� ��� (�������)").Range("R27:X27")
                    With objDoc.Tables(6).Cell(Row:=i, Column:=j)
                        .Range.Font.Bold = True
                        .Range.ParagraphFormat.Alignment = 1
                        .VerticalAlignment = 1
                        .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
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
    
    '����� �������
    numOfRows = 9
    numOfColumns = 3
    pos = objDoc.Paragraphs.count   '������ ������� ���������� ��� ����
    pos2 = objDoc.Paragraphs(pos).Range.End '���������� �������� �����
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    
    With objDoc.Tables(7)
        .Borders.Enable = False
        .Columns(1).SetWidth objWord.Application.CentimetersToPoints(3.6), 0
        .Columns(2).SetWidth objWord.Application.CentimetersToPoints(7.8), 0
        .Columns(3).SetWidth objWord.Application.CentimetersToPoints(5.6), 0
    End With
    
    '��������� �������
    i = 1
    j = 1

    For Each strA In Worksheets("��� ��� (�������)").Range("AO5:AQ13")
        If j = 2 Then
            objDoc.Tables(7).Cell(Row:=i, Column:=j).Range.Borders(-3).LineStyle = 1 'wdBottom / wdLineSttleSingle
        End If
        objDoc.Tables(7).Rows(i).SetHeight objWord.Application.CentimetersToPoints(0.8), 0
        objDoc.Tables(7).Rows(i).Cells.VerticalAlignment = 1 'wdAlignVerticalCenter
        objDoc.Tables(7).Cell(Row:=i, Column:=j).Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
        objDoc.Tables(7).Cell(Row:=i, Column:=j).Range.InsertAfter strA
        j = j + 1   '��������� �����
        If j > numOfColumns Then
            j = 1
            i = i + 1 '���� ���, ������ ������
            If i > numOfRows Then
                i = 1
            End If
        End If
    Next strA
    
    '��������� �������
    Set myPgNm = objDoc.Sections(1).Footers(1).PageNumbers.Add(PageNumberAlignment:=1, FirstPage:=True)
    myPgNm.Select
    With objWord.Selection.Range
        .Font.Name = "Times New Roman"
        .Font.Size = 12
    End With
    
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Select
    
'    MsgBox intReportCount & " ������� �������� ��������� � ����� " & ThisWorkbook.Path
End Sub
