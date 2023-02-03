Attribute VB_Name = "ReportTemplate"
Option Explicit

Sub Report(fName As String, strName As String, strPlace As String, shrtName As String, sepCalc As Boolean)
    'macros by Ivakhiv Roman - megazorch@gmail.com
    Dim rgData As Range
    Dim objWord As Object
    Dim i As Integer
    Dim fileName As String
    Dim objRange As Variant
    Dim objDoc As Variant
    Dim objSelection As Variant
    Dim dobovi As String, dobSUM As String, doboviTXT As String
    Dim projiv As String, projSM As String, projTXT As String
    Dim proizd As String, proizdSM As String, proizdTXT As String
    Dim totall As String
    Dim vtrCar As String, vtrCarTXT As String
    Dim vtrEls As String, vtrElsTXT As String
    Dim savePTH As String
    Dim dobPlus As Double, dobPlus_days As Integer
    Dim projPlus As Double, projPlus_days As Integer
    Dim proizdPlus As Double, proizdPlus_days As Integer
    Dim otherPlus As Double, carPlus As Double
    Dim checkSUM As Double
    Dim totalSUM As Double
    
    
    If sepCalc = True Then                                                      '18.09.2022 - #78
    
        dobPlus = WorksheetFunction.IfError(WorksheetFunction.VLookup(strName, Range("main_table"), 14, False), 0)          '18.09.2022
        dobPlus_days = WorksheetFunction.IfError(WorksheetFunction.VLookup(strName, Range("main_table"), 15, False), 0)
        projPlus = WorksheetFunction.IfError(WorksheetFunction.VLookup(strName, Range("main_table"), 16, False), 0)
        projPlus_days = WorksheetFunction.IfError(WorksheetFunction.VLookup(strName, Range("main_table"), 17, False), 0)
        proizdPlus = WorksheetFunction.IfError(WorksheetFunction.VLookup(strName, Range("main_table"), 18, False), 0)
        proizdPlus_days = WorksheetFunction.IfError(WorksheetFunction.VLookup(strName, Range("main_table"), 19, False), 0)
        carPlus = WorksheetFunction.IfError(WorksheetFunction.VLookup(strName, Range("main_table"), 20, False), 0)
        otherPlus = WorksheetFunction.IfError(WorksheetFunction.VLookup(strName, Range("main_table"), 21, False), 0)
        checkSUM = WorksheetFunction.IfError(WorksheetFunction.VLookup(strName, Range("main_table"), 22, False), 0)
        
        If dobPlus = 0 And Range("dobovi") = 0 Then
            doboviTXT = "                                                                                                                                               "
        ElseIf dobPlus = 0 And Range("dobovi") > 0 Then
            dobovi = Format(Range("dobovi"), "0.00")
            dobSUM = Format(Range("dobovi") * Range("dob_days"), "0.00")
            doboviTXT = "                           " & dobovi & " x " & Range("dob_days") & " = " & dobSUM & " грн.                                                                           "
        ElseIf dobPlus > 0 And Range("dobovi") = 0 Then
            dobSUM = Format(dobPlus * dobPlus_days, "0.00")
            doboviTXT = "                           " & Format(dobPlus, "0.00") & " x " & dobPlus_days & " = " & dobSUM & " грн.                                                                           "
        Else
            If Range("dobovi") = dobPlus Then
                dobovi = Format(Range("dobovi"), "0.00")
                dobSUM = Format(Range("dobovi") * (Range("dob_days") + dobPlus_days), "0.00")
                doboviTXT = "                           " & dobovi & " x " & (Range("dob_days") + dobPlus_days) & " = " & dobSUM & " грн.                                                                          "
            Else
                dobovi = Format(Range("dobovi"), "0.00")
                dobSUM = Format(Range("dobovi") * Range("dob_days") + dobPlus * dobPlus_days, "0.00")
                doboviTXT = "                           (" & dobovi & " x " & Range("dob_days") & ") + (" & Format(dobPlus, "0.00") & _
                " x " & dobPlus_days & ") = " & dobSUM & " грн.                                                "
            End If
        End If
        
        If projPlus = 0 And Range("projiv") = 0 Then
            projTXT = "                                                                                                                                      "
        ElseIf projPlus = 0 And Range("projiv") > 0 Then
            projiv = Format(Range("projiv"), "0.00")
            projSM = Format(Range("projiv") * Range("proj_days"), "0.00")
            projTXT = "                 " & projiv & " x " & Range("proj_days") & " = " & projSM & " грн.                                                                          "
        ElseIf projPlus > 0 And Range("projiv") = 0 Then
            projSM = Format(projPlus * projPlus_days, "0.00")
            projTXT = "                 " & Format(projPlus, "0.00") & " x " & projPlus_days & " = " & projSM & " грн.                                                                          "
        Else
            If Range("projiv") = projPlus Then
                 projiv = Format(Range("projiv"), "0.00")
                 projSM = Format(Range("projiv") * (Range("proj_days") + projPlus_days), "0.00")
                 projTXT = "                 " & projiv & " x " & (Range("proj_days") + projPlus_days) & " = " & projSM & " грн.                                                                          "
            Else
                projiv = Format(Range("projiv"), "0.00")
                projSM = Format(Range("projiv") * Range("proj_days") + projPlus * projPlus_days, "0.00")
                projTXT = "                 (" & projiv & " x " & Range("proj_days") & ") + (" & Format(projPlus, "0.00") & _
                " x " & projPlus_days & ") = " & projSM & " грн.                                                "
            End If
        End If
        
        If proizdPlus = 0 And Range("proizd") = 0 Then
            proizdTXT = "                                                                                                                                                "
        ElseIf proizdPlus = 0 And Range("proizd") > 0 Then
            proizd = Format(Range("proizd"), "0.00")
            proizdSM = Format(Range("proizd") * Range("proiz_days"), "0.00")
            proizdTXT = "                           " & proizd & " x " & Range("proiz_days") & " = " & proizdSM & " грн.                                                                          "
        ElseIf proizdPlus > 0 And Range("proizd") = 0 Then
            proizdSM = Format(proizdPlus * proizdPlus_days, "0.00")
            proizdTXT = "                           " & Format(proizdPlus, "0.00") & " x " & proizdPlus_days & " = " & proizdSM & " грн.                                                                            "
        Else
            If Range("proizd") = proizdPlus Then
                proizd = Format(Range("proizd"), "0.00")
                proizdSM = Format(Range("proizd") * (Range("proiz_days") + proizdPlus_days), "0.00")
                proizdTXT = "                           " & proizd & " x " & (Range("proiz_days") + proizdPlus_days) & " = " & proizdSM & " грн.                                                                              "
            Else
                proizd = Format(Range("proizd"), "0.00")
                proizdSM = Format(Range("proizd") * Range("proiz_days") + proizdPlus * proizdPlus_days, "0.00")
                proizdTXT = "                           (" & proizd & " x " & Range("proiz_days") & ") + (" & Format(proizdPlus, "0.00") & _
                " x " & proizdPlus_days & ") = " & proizdSM & " грн.                                                     "
            End If
        End If
        
        If carPlus = 0 And Range("for_car") = 0 Then
            vtrCarTXT = "                                                                          "
        ElseIf carPlus = 0 And Range("for_car") > 0 Then
            vtrCar = Format(Range("for_car"), "0.00")
            vtrCarTXT = "                        " & vtrCar & " грн.                               "
        ElseIf carPlus > 0 And Range("for_car") = 0 Then
            vtrCar = Format(carPlus, "0.00")
            vtrCarTXT = "                        " & vtrCar & " грн.                               "
        Else
            vtrCar = Format(Range("for_car") + carPlus, "0.00")
            vtrCarTXT = "                        " & vtrCar & " грн.                               "
        End If
        
        If otherPlus = 0 And Range("other") = 0 Then
            vtrElsTXT = "                                                                       "
        ElseIf otherPlus = 0 And Range("other") > 0 Then
            vtrEls = Format(Range("other"), "0.00")
            vtrElsTXT = "                     " & vtrEls & " грн.                               "
        ElseIf otherPlus > 0 And Range("other") = 0 Then
            vtrEls = Format(otherPlus, "0.00")
            vtrElsTXT = "                     " & vtrEls & " грн.                               "
        Else
            vtrEls = Format(Range("other") + otherPlus, "0.00")
            vtrElsTXT = "                     " & vtrEls & " грн.                               "
        End If
        
        totall = Format(Range("total_sum") + dobPlus * dobPlus_days + projPlus * projPlus_days + proizdPlus * proizdPlus_days + carPlus + otherPlus, "0.00")
        
        
        If Not totall = checkSUM Then
            MsgBox "Загальна сумма: " & totall & ", відрізняеться від контрольної суми: " & checkSUM & " ?!", vbCritical, "Помилка!"
            Exit Sub
        End If
            
    Else
        Select Case Range("dobovi")
            Case 0: doboviTXT = "                                                                                                                                               "
            Case 1 To 99999: dobovi = Format(Range("dobovi"), "0.00")
                dobSUM = Format(Range("dobovi") * Range("dob_days"), "0.00")
                doboviTXT = "                           " & dobovi & " x " & Range("dob_days") & " = " & dobSUM & " грн.                                                                           "
        End Select
        
        Select Case Range("projiv")
            Case 0: projTXT = "                                                                                                                                      "
            Case 1 To 99999: projiv = Format(Range("projiv"), "0.00")
                projSM = Format(Range("proj_days") * Range("projiv"), "0.00")
                projTXT = "                 " & projiv & " x " & Range("proj_days") & " = " & projSM & " грн.                                                                           "
        End Select
    
        Select Case Range("proizd")
            Case 0: proizdTXT = "                                                                                                                                                "
            Case 1 To 99999: proizd = Format(Range("proizd"), "0.00")
                proizdSM = Format(Range("proiz_days") * Range("proizd"), "0.00")
                proizdTXT = "                           " & proizd & " x " & Range("proiz_days") & " = " & proizdSM & " грн.                                                                             "
        End Select
    
        Select Case Range("for_car")
            Case 0: vtrCarTXT = "                                                                          "
            Case 1 To 99999: vtrCar = Format(Range("for_car"), "0.00")
                vtrCarTXT = "                        " & vtrCar & " грн.                               "
        End Select
    
        Select Case Range("other")
            Case 0: vtrElsTXT = "                                                                       "
            Case 1 To 99999: vtrEls = Format(Range("other"), "0.00")
                vtrElsTXT = "                     " & vtrEls & " грн.                               "
        End Select
    
        totall = Format(Range("total_sum"), "0.00")
        
    End If
            
        
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Add
    objWord.Visible = False
        
    'Поля
    With objDoc.PageSetup
        .LeftMargin = objWord.Application.CentimetersToPoints(Range("marg_left_4"))   '16.09.2022 - №36
        .RightMargin = objWord.Application.CentimetersToPoints(Range("marg_right_4"))
        .TopMargin = objWord.Application.CentimetersToPoints(Range("marg_top_4"))
        .BottomMargin = objWord.Application.CentimetersToPoints(Range("marg_bottom_4"))
    End With
        
    'ШАПКА
    With objWord.Selection
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1)
        .ParagraphFormat.SpaceAfter = objWord.Application.LinesToPoints(0)
        .Font.Size = 12
        .Font.name = "Times New Roman"
        .Font.Italic = True
        .ParagraphFormat.Alignment = 0 'Left
        .ParagraphFormat.LeftIndent = objWord.Application.CentimetersToPoints(10) 'отступ от края страницы
        .TypeText Text:="Додаток № 4"
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
        .TypeParagraph
        .Font.Bold = True 'Заголовок
        .Font.Size = 18
        .ParagraphFormat.Alignment = 1 'Center
        .TypeText Text:="Розрахунок витрат на відрядження"
        .TypeParagraph
        .Font.Bold = False 'Основний текст
        .Font.Size = 12
        .ParagraphFormat.Alignment = 0 'Left
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(1.25)
        .TypeText Text:="Згідно з наказом від " & Range("order_date") & " № " & Range("order") _
        & ", виданий "  ' 17.08.2022; 16.09.2022 - #54
        .Font.Underline = True  ' 17.08.2022
        .TypeText fName         ' 17.08.2022
        .Font.Underline = False ' 17.08.2022
        .TypeText " до " & strPlace & "."   ' 17.08.2022
        .TypeParagraph
        .ParagraphFormat.FirstLineIndent = objWord.Application.CentimetersToPoints(0)
        .TypeText "Дати відрядження"        ' 17.08.2022 - changed - "Д..."
        .Font.Underline = True
        .TypeText "        " & Range("date_comm_num") & " - " & Range("date_comp_num") & " р.                                                                         "
        .TypeParagraph
        .Font.Underline = False
        .TypeText "Добові"
        .Font.Underline = True
        .TypeText doboviTXT
        .TypeParagraph
        .Font.Underline = False
        .TypeText "Проживання"
        .Font.Underline = True
        .TypeText projTXT
        .Font.Underline = False
        .TypeParagraph
        .TypeText "Проїзд"
        .Font.Underline = True
        .TypeText proizdTXT
        .Font.Underline = False
        .TypeParagraph
        .TypeText "Витрати на автомобіль (парковка, мийка, інші) "
        .Font.Underline = True
        .TypeText vtrCarTXT
        .Font.Underline = False
        .TypeParagraph
        .TypeText "Інші витрати (плата за участь у виставках, тощо) "
        .Font.Underline = True
        .TypeText vtrElsTXT
        .Font.Underline = False
        .TypeParagraph
        .Font.Bold = True
        .TypeText "ВСЬОГО"
        .Font.Underline = True
        .TypeText "    " & totall & "    грн."
        .Font.Underline = False
        .Font.Bold = False
        .TypeParagraph
        .TypeParagraph
        .TypeText "Розрахунок склав _____________________ " & shrtName
        .ParagraphFormat.SpaceAfter = 0     '17.08.2022 #21
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1) '17.08.2022 #21
        .TypeParagraph
        .Font.Size = 8
        .TypeText "                                                                     (підпис)                                 (ПІБ відряджуваного)"
        .TypeParagraph
        .ParagraphFormat.SpaceAfter = 10
        .TypeParagraph
        .Font.Size = 12
        .TypeText """ПЕРЕВІРЕНО"""
        .TypeParagraph
        .TypeText "___________________                          " & Range("bookkeeper_4") '16.09.2022 - №37
        
        .ParagraphFormat.SpaceAfter = 0     '17.08.2022 #21
        .ParagraphFormat.LineSpacing = objWord.Application.LinesToPoints(1) '17.08.2022 #21
        .TypeParagraph
        .Font.Size = 8
        .TypeText "                  (підпис)                                                                                    (ПІБ бухгалтера)"
    End With
    
    savePTH = ActiveWorkbook.pATH & "\"
    objDoc.SaveAs savePTH & "Розрахунок витрат на відрядж. - " & shrtName & ".docx"
    objWord.Quit
    Set objWord = Nothing

End Sub
