Attribute VB_Name = "ReportTemplate_2"
Sub Report_2(fName As String, strPlace As String, shrtName As String, purpose As String, car As String, garage As String, days As String)
    'macros by Ivakhiv Roman - megazorch@gmail.com
    Dim rgData As Range
    Dim objWord As Object
    Dim i As Integer
    Dim fileName As String
    Dim objRange As Variant
    Dim objDoc As Variant
    Dim objSelection As Variant
    Dim savePTH As String
    
            
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Add
    objWord.Visible = False
        
    '����
    With objDoc.PageSetup
        .LeftMargin = objWord.Application.CentimetersToPoints(Range("marg_left_10"))   '16.09.2022 - #38
        .RightMargin = objWord.Application.CentimetersToPoints(Range("marg_right_10"))
        .TopMargin = objWord.Application.CentimetersToPoints(Range("marg_top_10"))
        .BottomMargin = objWord.Application.CentimetersToPoints(Range("marg_bottom_10"))
    End With
        
    '�����
    With objWord.Selection
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1)
        .ParagraphFormat.SpaceAfter = 0
        .Font.Size = 10
        .Font.name = "Times New Roman"
        .Font.Italic = True
        .ParagraphFormat.Alignment = 0 'Left
        .ParagraphFormat.LeftIndent = objWord.Application.CentimetersToPoints(10) '������ �� ���� ��������
        .TypeText Text:="������� � 10"
        .TypeParagraph
        .TypeText Text:="�� ��������� ��� ���������� ��������"
        .TypeParagraph
        .TypeText Text:="��� ���������� ��� ""�������� ��� ������"""
        .Font.Italic = False
        .TypeParagraph
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1.5)
        .ParagraphFormat.LeftIndent = objWord.Application.CentimetersToPoints(0)
        .ParagraphFormat.SpaceAfter = 6
        .TypeParagraph
        .TypeParagraph
        .Font.Bold = True '���������
        .Font.Size = 16
        .ParagraphFormat.Alignment = 1 'Center
        .TypeText Text:="��� ��� ��������� �������� �� ���������� �� �����"
        .TypeParagraph
        .Font.Size = 12
        .Font.Bold = False '�������� �����
        .ParagraphFormat.Alignment = 3 'Left
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(1.25)
    End With

    Set objRange = objDoc.Range
        
    numOfRows = 2
    numOfColumns = 1
    pos = objDoc.Paragraphs.Count   '������ ������� ���������� ��� ����
    pos2 = objDoc.Paragraphs(pos).Range.End '���������� �������� �����
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    
    With objDoc.Tables(1)
        .Borders.InsideLineStyle = 1
        .Columns(1).Cells.VerticalAlignment = 1 'wdAlignVerticalCenter
    End With
    
    With objDoc.Tables(1).Cell(Row:=1, Column:=1)
        .Range.ParagraphFormat.Alignment = 1
        .VerticalAlignment = 1
        .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
        .Range.ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1)
        .Range.Font.Size = 14
        .Range.InsertAfter fName
    End With
    
    With objDoc.Tables(1).Cell(Row:=2, Column:=1)
        .Range.ParagraphFormat.Alignment = 1
        .VerticalAlignment = 1
        .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
        .Range.InsertAfter "(ϲ�)"
    End With
    
    objWord.Selection.MoveDown Unit:=5, Count:=3
    
    With objWord.Selection
        .TypeText "��������� � ���������� ��������� �� " & strPlace & "."
        .TypeParagraph
        .TypeText purpose & ", ����� ������ �" & Range("order") & " �� " & Range("order_date") '16.09.2022 - #53
        .TypeParagraph
        .TypeText "����� ���������� " & Range("dob_days") & days & Range("commence") & " �� " & Range("complete") '17.08.2022 - �11; 16.09.2022 - #52
        .TypeParagraph
    End With
    
    If car <> "" Then
        objWord.Selection.TypeText "����� ��������������� - " & car & "."
        objWord.Selection.TypeParagraph
    End If
    If garage <> "" Then
        objWord.Selection.TypeText "̳��� ����������� �������������� � " & garage & "."
        objWord.Selection.TypeParagraph
    End If
    With objWord.Selection
        .TypeParagraph
        .ParagraphFormat.LeftIndent = objWord.Application.CentimetersToPoints(12.5) '������ �� ���� ��������
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1)
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.LineUnitBefore = objWord.Application.LinesToPoints(0)
        .ParagraphFormat.LineUnitAfter = objWord.Application.LinesToPoints(0)
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(0)
        .TypeText "__________________"
        .TypeParagraph
        .Font.Italic = True
        .Font.Size = 9
        .TypeText "     (����� �����������)"
        .TypeParagraph
        .TypeParagraph
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1.5)
        .ParagraphFormat.LeftIndent = objWord.Application.CentimetersToPoints(0)
        .ParagraphFormat.SpaceAfter = 6
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(0)
        .TypeParagraph
        .Font.Size = 16
        .Font.Italic = False
        .Font.Bold = True
        .ParagraphFormat.Alignment = 1 'Center
        .TypeText Text:="�������� �������� ��� ��������� �������� �� ����������"
        .TypeParagraph
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1.15)
    End With
    
    numOfRows = 2
    numOfColumns = 1
    pos = objDoc.Paragraphs.Count   '������ ������� ���������� ��� ����
    pos2 = objDoc.Paragraphs(pos).Range.End '���������� �������� �����
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    
    objDoc.Tables(2).Borders.InsideLineStyle = 1
    objDoc.Tables(2).Borders(-3).LineStyle = 1 'wdBorderBottom = -3
    
    With objWord.Selection
        .MoveDown Unit:=5, Count:=3
        .ParagraphFormat.LeftIndent = objWord.Application.CentimetersToPoints(8)
        .TypeParagraph
        .Font.Bold = False
    End With
        
    numOfRows = 2
    numOfColumns = 3
    pos = objDoc.Paragraphs.Count   '������ ������� ���������� ��� ����
    pos2 = objDoc.Paragraphs(pos).Range.End '���������� �������� �����
    objDoc.Tables.Add objDoc.Range(pos2 - 1, pos2 - 1), numOfRows, numOfColumns
    
    With objDoc.Tables(3)
        .Rows.SetLeftIndent LeftIndent:=175.5, RulerStyle:=2  ' wdAdjustFirstColumn = 2
        .Columns(1).SetWidth objWord.Application.CentimetersToPoints(4.5), 0
        .Columns(1).Cells.VerticalAlignment = 1 'wdAlignVerticalCenter
        .Columns(2).SetWidth objWord.Application.CentimetersToPoints(1.5), 0
        .Columns(2).Cells.VerticalAlignment = 1 'wdAlignVerticalCenter
        .Columns(3).SetWidth objWord.Application.CentimetersToPoints(4.5), 0
        .Columns(3).Cells.VerticalAlignment = 1 'wdAlignVerticalCenter
    End With
    
    With objDoc.Tables(3).Cell(Row:=1, Column:=1)
        .Range.ParagraphFormat.Alignment = 1
        .VerticalAlignment = 3 'Bottom
        .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
        .Range.ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1)
        .Range.Font.Size = 12
        .Range.Borders(-3).LineStyle = 1
    End With
    
    With objDoc.Tables(3).Cell(Row:=1, Column:=3)
        .Range.ParagraphFormat.Alignment = 1
        .VerticalAlignment = 3 'Bottom
        .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
        .Range.ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1)
        .Range.Font.Size = 12
        .Range.Borders(-3).LineStyle = 1
        .Range.InsertAfter Range("head_10")  '16.09.2022 - #39
    End With
    
    With objDoc.Tables(3).Cell(Row:=2, Column:=1)
        .Range.ParagraphFormat.Alignment = 1
        .VerticalAlignment = 0 'Top
        .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
        .Range.ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1)
        .Range.Font.Size = 10
        .Range.InsertAfter "(����� ��������)"
    End With
    
    With objDoc.Tables(3).Cell(Row:=2, Column:=3)
        .Range.ParagraphFormat.Alignment = 1
        .VerticalAlignment = 0 'Top
        .Range.ParagraphFormat.SpaceAfter = 0 '������� �������� ����� ������
        .Range.ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1)
        .Range.Font.Size = 10
        .Range.InsertAfter "(ϲ�)"
    End With
    
      

    
    savePTH = ActiveWorkbook.pATH & "\"
    objDoc.SaveAs savePTH & "��� ��� ��������� �������� - " & shrtName & ".docx"
    objWord.Quit
    Set objWord = Nothing

End Sub

