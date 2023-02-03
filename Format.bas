Attribute VB_Name = "Format"
Option Explicit
'=�(F108<>"��";F108<>"��";F108<>"���";F108<>"���";��(������(F108));F109<>"�")
Sub Format_color()
Attribute Format_color.VB_ProcData.VB_Invoke_Func = " \n14"
' Format_color ������
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
        sReturn = InputBox("������� �������� ������� �������:" & vbCrLf & "(��������: H36:AJ400 ��� h36:aj400," & vbCrLf & "�� ����. �����)", "������� ��������� �� �������")
        If StrPtr(sReturn) = 0 Then
            MsgBox "�� ������ Cancel ��� ESC.", vbOKOnly, "������ ��������"
        ElseIf sReturn = "" Then
            MsgBox "�� ������ OK, ������ �� �������, ���� ��������� ������.", vbOKOnly, "��� ������"
        Else
            MsgBox "�� ������ OK, ������ �������. ��������� ��������� ��� main_table.", vbOKOnly, "������ �������"
            ActiveWorkbook.Names.Add Name:="main_table", RefersTo:=Range(sReturn), Visible:=True
        End If
    End If
    
    For Each i In Range("main_table")
        i.Select
        Select Case i
            Case Is = "��" 'blue
                With Selection.Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent1
                    .TintAndShade = 0.799981688894314
'                    .PatternTintAndShade = 0
                End With
                ActiveCell.Offset(1, 0).Range("A1").Value = "x"
            Case Is = "��"   'green
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
            Case Is = "���"  'green
                With Selection.Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent6
                    .TintAndShade = 0.799981688894314
'                    .PatternTintAndShade = 0
                End With
            Case Is = "²"   'pink
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
            Case Is = "���"  'yellow
                With Selection.Interior
'                    .Pattern = xlSolid
'                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent4
                    .TintAndShade = 0.799981688894314
'                    .PatternTintAndShade = 0
                End With
            Case Is = "��"
                ActiveCell.Offset(1, 0).Range("A1").Select
                With Selection.Font
                    .ThemeColor = xlThemeColorDark1
'                    .TintAndShade = 0
                End With
            Case Is = "�"
                With Selection.Font         '������ ������� ������
                    .ColorIndex = xlAutomatic
'                    .TintAndShade = 0
                End With
            Case "��", "��", "��", "��", "��", "���", "���", "�", "�", "�", "Ĳ", "��", "��", "��", "�", "��", "��", "��", "��", "��", "��", "��", "��", "�", "��", "��", "��", "��", "��", "�", "���", "�", "��", "��", "��", "��", "���", "��", "���", "��"
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
        sReturn = InputBox("������� �������� ������� �������:" & vbCrLf & "(��������: H36:AJ400 ��� h36:aj400," & vbCrLf & "�� ����. �����)", "������� ��������� �� �������")
        If StrPtr(sReturn) = 0 Then
            MsgBox "�� ������ Cancel ��� ESC.", vbOKOnly, "������ ��������"
        ElseIf sReturn = "" Then
            MsgBox "�� ������ OK, ������ �� �������, ���� ��������� ������.", vbOKOnly, "��� ������"
        Else
            MsgBox "�� ������ OK, ������ �������. ��������� ��������� ��� main_table.", vbOKOnly, "������ �������"
            ActiveWorkbook.Names.Add Name:="main_table", RefersTo:=Range(sReturn), Visible:=True
        End If
    End If
    Dim i As Range
    Cells.FormatConditions.Delete
    Application.ScreenUpdating = False
    For Each i In Range("main_table")
        i.Select
        If Not ActiveCell.MergeCells Then
            With Selection.Interior     '������ �������
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With Selection.Font         '������ ������� ������
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
            End With
        End If
    Next i
End Sub

Sub ChangeRefer()
    Dim sReturn As String
    Dim refer As Range
    sReturn = InputBox("������� �������� ������� ����:" & vbCrLf & "(��������: H36:AJ400 ��� h36:aj400," & vbCrLf & "�� ����. �����)", "����� ���������")
    If StrPtr(sReturn) = 0 Then
        MsgBox "�� ������ Cancel ��� ESC.", vbOKOnly, "������ ��������"
    ElseIf sReturn = "" Then
        MsgBox "�� ������ OK, ������ �� �������.", vbOKOnly, "��� ������"
    Else
        MsgBox "�� ������ OK, ������ �������. ������ �� �������� ��������.", vbOKOnly, "������ �������"
        ActiveWorkbook.Names.Add Name:="main_table", RefersTo:=Range(sReturn), Visible:=True
    End If
End Sub
