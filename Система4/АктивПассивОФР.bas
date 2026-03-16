Attribute VB_Name = "АктивПассивОФР"
Sub ОбновитьДанные()

    Dim wsСистема As Worksheet
    Dim wsОтчет As Worksheet
    Dim i As Integer
    Dim значение2025 As Variant
    Dim значение2024 As Variant
    Dim разница As Double
    Dim процентОтчетности As Variant
    Dim кодСтатьи As Variant
    Dim значениеG As Variant
    Dim значениеI As Variant
    Dim значениеL As Variant
    Dim значениеO As Variant
    
    Set wsСистема = ThisWorkbook.Sheets("Система4")
    Set wsОтчет = ThisWorkbook.Sheets("Отчетность")
    
    ' --- 1. Актив в A162 ---
    wsСистема.Range("A162").Value = "Актив баланса представлен на " & _
        Format(wsОтчет.Range("M10").Value * 100, "0") & "% внеоборотными активами и на " & _
        Format(wsОтчет.Range("M18").Value * 100, "0") & "% оборотными."

    ' --- 2. Пассив в A226 ---
    wsСистема.Range("A226").Value = "Пассив баланса представлен собственным капиталом на " & _
        Format(wsОтчет.Range("M30").Value * 100, "0") & "% и на " & _
        Format((wsОтчет.Range("M36").Value + wsОтчет.Range("M43").Value) * 100, "0") & "% обязательствами."

    ' --- 3. Заполнение и видимость строк 166-171 (активы) ---
    wsСистема.Rows("166:171").Hidden = False
    For i = 166 To 171
        значение2025 = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F4:L18"), 7, False)
        If IsError(значение2025) Then
            wsСистема.Cells(i, 3).Value = ""
        Else
            wsСистема.Cells(i, 3).Value = значение2025
        End If
        
        ' Заполняем F (динамика vs 2024)
        значение2024 = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F4:L18"), 4, False)
        If IsError(значение2024) Then
            wsСистема.Cells(i, 7).Value = "Ошибка поиска"
        Else
            разница = Abs(значение2025 - значение2024)
            If значение2025 = значение2024 Then
                wsСистема.Cells(i, 7).Value = "Значение осталось прежним по сравнению с 2024"
            ElseIf значение2025 > значение2024 Then
                wsСистема.Cells(i, 7).Value = "Увеличилось на " & Format(разница, "### ### ### ###") & " по сравнению с 2024"
            Else
                wsСистема.Cells(i, 7).Value = "Снизилось на " & Format(разница, "### ### ### ###") & " по сравнению с 2024"
            End If
        End If
        
        процентОтчетности = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F4:M18"), 8, False)
        If IsError(процентОтчетности) Then
            wsСистема.Rows(i).Hidden = True
        Else
            If процентОтчетности < 0.05 Then
                wsСистема.Rows(i).Hidden = True
            Else
                wsСистема.Rows(i).Hidden = False
            End If
        End If
    Next i

    ' --- 4. Заполнение и видимость строк 175–180 (активы) ---
    wsСистема.Rows("175:180").Hidden = False
    For i = 175 To 180
        значение2025 = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F4:L18"), 7, False)
        If IsError(значение2025) Then
            wsСистема.Cells(i, 3).Value = ""
        Else
            wsСистема.Cells(i, 3).Value = значение2025
        End If
        
        ' Заполняем F (динамика vs 2024)
        значение2024 = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F4:L18"), 4, False)
        If IsError(значение2024) Then
            wsСистема.Cells(i, 7).Value = "Ошибка поиска"
        Else
            разница = Abs(значение2025 - значение2024)
            If значение2025 = значение2024 Then
                wsСистема.Cells(i, 7).Value = "Значение осталось прежним по сравнению с 2024"
            ElseIf значение2025 > значение2024 Then
                wsСистема.Cells(i, 7).Value = "Увеличилось на " & Format(разница, "### ### ### ###") & " по сравнению с 2024"
            Else
                wsСистема.Cells(i, 7).Value = "Снизилось на " & Format(разница, "### ### ### ###") & " по сравнению с 2024"
            End If
        End If
        
        процентОтчетности = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F4:M18"), 8, False)
        If IsError(процентОтчетности) Then
            If i = 177 Then
                wsСистема.Rows(i).Hidden = False
            Else
                wsСистема.Rows(i).Hidden = True
            End If
        Else
            If i = 177 Then
                wsСистема.Rows(i).Hidden = False ' Дебиторка — всегда видна
            ElseIf процентОтчетности < 0.05 Then
                wsСистема.Rows(i).Hidden = True
            Else
                wsСистема.Rows(i).Hidden = False
            End If
        End If
    Next i

    ' --- 5. Заполнение и видимость строк 230–235 (собственный капитал) ---
    wsСистема.Rows("230:235").Hidden = False
    For i = 230 To 235
        значение2025 = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F23:L44"), 7, False)
        If IsError(значение2025) Then
            wsСистема.Cells(i, 3).Value = ""
            wsСистема.Rows(i).Hidden = True
        Else
            wsСистема.Cells(i, 3).Value = значение2025
            If значение2025 = 0 Then
                wsСистема.Rows(i).Hidden = True
            Else
                wsСистема.Rows(i).Hidden = False
            End If
        End If
        
        ' Заполняем F (динамика vs 2024)
        значение2024 = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F23:L44"), 4, False)
        If IsError(значение2024) Then значение2024 = 0
        
        If Not IsError(значение2025) Then
            разница = Abs(значение2025 - значение2024)
            If значение2025 = значение2024 Then
                wsСистема.Cells(i, 7).Value = "Значение осталось прежним по сравнению с 2024"
            ElseIf значение2025 > значение2024 Then
                wsСистема.Cells(i, 7).Value = "Увеличилось на " & Format(разница, "### ### ### ###") & " по сравнению с 2024"
            Else
                wsСистема.Cells(i, 7).Value = "Снизилось на " & Format(разница, "### ### ### ###") & " по сравнению с 2024"
            End If
        End If
    Next i

    ' --- 6. Заполнение и видимость строк 239–247 (обязательства) ---
    wsСистема.Rows("239:247").Hidden = False
    For i = 239 To 247
        значение2025 = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F23:L44"), 7, False)
        If IsError(значение2025) Then
            wsСистема.Cells(i, 3).Value = ""
        Else
            wsСистема.Cells(i, 3).Value = значение2025
        End If
        
        процентОтчетности = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F23:M44"), 8, False)
        If IsError(процентОтчетности) Then
            If i = 244 Then
                wsСистема.Rows(i).Hidden = False
            Else
                wsСистема.Rows(i).Hidden = True
            End If
        Else
            If i = 244 Then
                wsСистема.Rows(i).Hidden = False ' Кредиторка — всегда видна
            ElseIf процентОтчетности < 0.05 Then
                wsСистема.Rows(i).Hidden = True
            Else
                wsСистема.Rows(i).Hidden = False
            End If
        End If
        
        ' Заполняем F (динамика vs 2024)
        значение2024 = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F23:L44"), 4, False)
        If IsError(значение2024) Then значение2024 = 0
        
        If Not IsError(значение2025) Then
            разница = Abs(значение2025 - значение2024)
            If значение2025 = значение2024 Then
                wsСистема.Cells(i, 7).Value = "Значение осталось прежним по сравнению с 2024"
            ElseIf значение2025 > значение2024 Then
                wsСистема.Cells(i, 7).Value = "Увеличилось на " & Format(разница, "### ### ### ###") & " по сравнению с 2024"
            Else
                wsСистема.Cells(i, 7).Value = "Снизилось на " & Format(разница, "### ### ### ###") & " по сравнению с 2024"
            End If
        End If
    Next i

    ' --- 7. Заполнение строк 296–306 (ОФР: G, I, L, O > C, E, G, I) ---
    wsСистема.Rows("296:306").Hidden = False
    For i = 296 To 306
        значениеG = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F47:O64"), 2, False)
        значениеI = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F47:O64"), 4, False)
        значениеL = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F47:O64"), 7, False)
        значениеO = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F47:O64"), 10, False)
        
        wsСистема.Cells(i, 3).Value = IIf(IsError(значениеG), "", значениеG)
        wsСистема.Cells(i, 5).Value = IIf(IsError(значениеI), "", значениеI)
        wsСистема.Cells(i, 7).Value = IIf(IsError(значениеL), "", значениеL)
        wsСистема.Cells(i, 9).Value = IIf(IsError(значениеO), "", значениеO)
        
        ' Скрываем строку, если все 4 ячейки пустые или равны 0
        If (IsEmpty(wsСистема.Cells(i, 3).Value) Or wsСистема.Cells(i, 3).Value = 0) And _
           (IsEmpty(wsСистема.Cells(i, 5).Value) Or wsСистема.Cells(i, 5).Value = 0) And _
           (IsEmpty(wsСистема.Cells(i, 7).Value) Or wsСистема.Cells(i, 7).Value = 0) And _
           (IsEmpty(wsСистема.Cells(i, 9).Value) Or wsСистема.Cells(i, 9).Value = 0) Then
            wsСистема.Rows(i).Hidden = True
        Else
            wsСистема.Rows(i).Hidden = False
        End If
    Next i

    ' --- 8. Заполнение строк 309–319 (динамика ОФР: L vs O > C и F) ---
    wsСистема.Rows("309:319").Hidden = False
    For i = 309 To 319
        значениеL = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F47:O64"), 7, False)
        значениеO = Application.VLookup(wsСистема.Cells(i, 11).Value, wsОтчет.Range("F47:O64"), 10, False)
        
        If IsError(значениеL) Then значениеL = 0
        If IsError(значениеO) Then значениеO = 0
        
        wsСистема.Cells(i, 3).Value = значениеL
        
        If значениеL = 0 And значениеO = 0 Then
            wsСистема.Rows(i).Hidden = True
            wsСистема.Cells(i, 7).Value = ""
        Else
            wsСистема.Rows(i).Hidden = False
            разница = значениеL - значениеO
            
            If значениеL = значениеO Then
                wsСистема.Cells(i, 7).Value = "Значение осталось прежним"
            ElseIf значениеL > значениеO Then
                wsСистема.Cells(i, 7).Value = "Увеличилось на " & Format(Abs(разница), "### ### ### ###")
            Else
                wsСистема.Cells(i, 7).Value = "Снизилось на " & Format(Abs(разница), "### ### ### ###")
            End If
        End If
    Next i

    MsgBox "Данные успешно обновлены!", vbInformation

End Sub





