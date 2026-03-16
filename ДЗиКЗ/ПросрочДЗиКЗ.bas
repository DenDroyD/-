Attribute VB_Name = "ПросрочДЗиКЗ"
Option Explicit
' =============================================
' Макрос для поиска просроченной дебиторской/кредиторской задолженности
' Версия: 3.3 (расширенный комментарий с цифрами и процентами)
' - Шапка таблицы на строке 2, данные с 3
' - Исключены контрагенты-"цифры" (без букв)
' - Сумма с разделителем разрядов и двумя знаками после запятой
' - Сортировка по убыванию суммы
' - Ширина, выравнивание, перенос текста по ТЗ
' - ГРАНИЦЫ для всей таблицы, кроме столбца J (полностью белый, без границ)
' - КОММЕНТАРИЙ в столбце K с детальными расчетами (проценты, суммы)
' =============================================

Sub FindOverdueDebt()
    Dim ws62 As Worksheet, ws60 As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, i As Long
    Dim targetRow As Long
    Dim accountNum As String, contractor As String
    Dim startDebit As Double, startCredit As Double
    Dim endDebit As Double, endCredit As Double
    Dim debitTurnover As Double, creditTurnover As Double
    Dim frozenAmount As Double
    Dim lastDataRow As Long
    Dim sortRange As Range
    Dim scenario As String   ' "чистая" или "непогаш"
    
    ' --- Настройка ---
    Set ws62 = ThisWorkbook.Sheets("ОСВ 62")
    Set ws60 = ThisWorkbook.Sheets("ОСВ 60")
    
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("Управление")
    On Error GoTo 0
    
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsTarget.Name = "Управление"
    End If
    
    ' --- Очистка области результатов (столбцы F:K) ---
    wsTarget.Range("F:K").ClearContents
    wsTarget.Range("F:K").ClearFormats
    
    ' --- Установка заголовков на СТРОКЕ 2 и форматирование столбцов ---
    With wsTarget
        .Range("F2").Value = "Дебитор/Кредитор"
        .Range("G2").Value = "Контрагент"
        .Range("H2").Value = "Счет"
        .Range("I2").Value = "Сумма"
        .Range("K2").Value = "Комментарий"
        .Range("F2:K2").Font.Bold = True
        .Range("F2:K2").Interior.color = RGB(220, 230, 241)
        
        .Columns("F").ColumnWidth = 18
        .Columns("F").HorizontalAlignment = xlLeft
        
        .Columns("G").ColumnWidth = 30
        .Columns("G").HorizontalAlignment = xlLeft
        .Columns("G").WrapText = True
        
        .Columns("H").ColumnWidth = 10
        .Columns("H").HorizontalAlignment = xlCenter
        
        .Columns("I").ColumnWidth = 19
        .Columns("I").HorizontalAlignment = xlCenter
        .Columns("I").NumberFormat = "### ### ### ###"
        
        .Columns("J").ColumnWidth = 8.43
        .Columns("J").Interior.ColorIndex = xlNone
        
        .Columns("K").ColumnWidth = 110
        .Columns("K").HorizontalAlignment = xlLeft
        .Columns("K").WrapText = True
        .Columns("K").VerticalAlignment = xlCenter
        
        .Range("F2:K2").HorizontalAlignment = xlCenter
        .Range("F2:K2").VerticalAlignment = xlCenter
    End With
    
    targetRow = 3
    
    ' --- Анализ листа ОСВ 62 ---
    lastRow = ws62.Cells(ws62.Rows.count, "A").End(xlUp).row
    For i = 4 To lastRow
        If Trim(ws62.Cells(i, 1).Value) <> "" And Trim(ws62.Cells(i, 2).Value) <> "" Then
            accountNum = ws62.Cells(i, 1).Value
            contractor = Trim(ws62.Cells(i, 2).Value)
            
            If Not ContainsLetter(contractor) Then GoTo NextRow62
            
            startDebit = val(Replace(ws62.Cells(i, 3).Value, " ", ""))
            startCredit = val(Replace(ws62.Cells(i, 4).Value, " ", ""))
            debitTurnover = val(Replace(ws62.Cells(i, 5).Value, " ", ""))
            creditTurnover = val(Replace(ws62.Cells(i, 6).Value, " ", ""))
            endDebit = val(Replace(ws62.Cells(i, 7).Value, " ", ""))
            endCredit = val(Replace(ws62.Cells(i, 8).Value, " ", ""))
            
            ' --- Дебиторка (Дт) ---
            If startDebit > 0 And endDebit > 0 Then
                If startDebit = endDebit And debitTurnover = 0 And creditTurnover = 0 Then
                    scenario = "чистая"
                    wsTarget.Cells(targetRow, 6).Value = "Дебет"
                    wsTarget.Cells(targetRow, 7).Value = contractor
                    wsTarget.Cells(targetRow, 8).Value = accountNum
                    wsTarget.Cells(targetRow, 9).Value = startDebit
                    ' Столбец J пустой
                    wsTarget.Cells(targetRow, 11).Value = GenerateComment( _
                        debtType:="Дебет", _
                        account:=accountNum, _
                        scenario:=scenario, _
                        startBalance:=startDebit, _
                        repaymentTurnover:=creditTurnover, _
                        frozenAmount:=startDebit, _
                        endBalance:=endDebit)
                    targetRow = targetRow + 1
                End If
            End If
            
            If startDebit > 0 And endDebit >= startDebit Then
                frozenAmount = Application.WorksheetFunction.Max(0, startDebit - creditTurnover)
                If frozenAmount > 0 And endDebit > 0 Then
                    If Not (startDebit = endDebit And debitTurnover = 0 And creditTurnover = 0) Then
                        scenario = "непогаш"
                        wsTarget.Cells(targetRow, 6).Value = "Дебет"
                        wsTarget.Cells(targetRow, 7).Value = contractor
                        wsTarget.Cells(targetRow, 8).Value = accountNum
                        wsTarget.Cells(targetRow, 9).Value = frozenAmount
                        wsTarget.Cells(targetRow, 11).Value = GenerateComment( _
                            debtType:="Дебет", _
                            account:=accountNum, _
                            scenario:=scenario, _
                            startBalance:=startDebit, _
                            repaymentTurnover:=creditTurnover, _
                            frozenAmount:=frozenAmount, _
                            endBalance:=endDebit)
                        targetRow = targetRow + 1
                    End If
                End If
            End If
            
            ' --- Кредиторка (Кт) ---
            If startCredit > 0 And endCredit >= startCredit Then
                frozenAmount = Application.WorksheetFunction.Max(0, startCredit - debitTurnover)
                If frozenAmount > 0 Then
                    scenario = IIf(startCredit = endCredit And debitTurnover = 0 And creditTurnover = 0, "чистая", "непогаш")
                    wsTarget.Cells(targetRow, 6).Value = "Кредит"
                    wsTarget.Cells(targetRow, 7).Value = contractor
                    wsTarget.Cells(targetRow, 8).Value = accountNum
                    wsTarget.Cells(targetRow, 9).Value = frozenAmount
                    wsTarget.Cells(targetRow, 11).Value = GenerateComment( _
                        debtType:="Кредит", _
                        account:=accountNum, _
                        scenario:=scenario, _
                        startBalance:=startCredit, _
                        repaymentTurnover:=debitTurnover, _
                        frozenAmount:=frozenAmount, _
                        endBalance:=endCredit)
                    targetRow = targetRow + 1
                End If
            End If
        End If
NextRow62:
    Next i
    
    ' --- Анализ листа ОСВ 60 ---
    lastRow = ws60.Cells(ws60.Rows.count, "A").End(xlUp).row
    For i = 4 To lastRow
        If Trim(ws60.Cells(i, 1).Value) <> "" And Trim(ws60.Cells(i, 2).Value) <> "" Then
            accountNum = ws60.Cells(i, 1).Value
            contractor = Trim(ws60.Cells(i, 2).Value)
            
            If Not ContainsLetter(contractor) Then GoTo NextRow60
            
            startDebit = val(Replace(ws60.Cells(i, 3).Value, " ", ""))
            startCredit = val(Replace(ws60.Cells(i, 4).Value, " ", ""))
            debitTurnover = val(Replace(ws60.Cells(i, 5).Value, " ", ""))
            creditTurnover = val(Replace(ws60.Cells(i, 6).Value, " ", ""))
            endDebit = val(Replace(ws60.Cells(i, 7).Value, " ", ""))
            endCredit = val(Replace(ws60.Cells(i, 8).Value, " ", ""))
            
            ' --- Кредиторка (Кт) по 60 ---
            If startCredit > 0 And endCredit >= startCredit Then
                frozenAmount = Application.WorksheetFunction.Max(0, startCredit - debitTurnover)
                If frozenAmount > 0 Then
                    scenario = IIf(startCredit = endCredit And debitTurnover = 0 And creditTurnover = 0, "чистая", "непогаш")
                    wsTarget.Cells(targetRow, 6).Value = "Кредит"
                    wsTarget.Cells(targetRow, 7).Value = contractor
                    wsTarget.Cells(targetRow, 8).Value = accountNum
                    wsTarget.Cells(targetRow, 9).Value = frozenAmount
                    wsTarget.Cells(targetRow, 11).Value = GenerateComment( _
                        debtType:="Кредит", _
                        account:=accountNum, _
                        scenario:=scenario, _
                        startBalance:=startCredit, _
                        repaymentTurnover:=debitTurnover, _
                        frozenAmount:=frozenAmount, _
                        endBalance:=endCredit)
                    targetRow = targetRow + 1
                End If
            End If
            
            ' --- Дебиторка (Дт) по 60 (авансы выданные) ---
            If startDebit > 0 And endDebit >= startDebit Then
                frozenAmount = Application.WorksheetFunction.Max(0, startDebit - creditTurnover)
                If frozenAmount > 0 Then
                    scenario = IIf(startDebit = endDebit And debitTurnover = 0 And creditTurnover = 0, "чистая", "непогаш")
                    wsTarget.Cells(targetRow, 6).Value = "Дебет"
                    wsTarget.Cells(targetRow, 7).Value = contractor
                    wsTarget.Cells(targetRow, 8).Value = accountNum
                    wsTarget.Cells(targetRow, 9).Value = frozenAmount
                    wsTarget.Cells(targetRow, 11).Value = GenerateComment( _
                        debtType:="Дебет", _
                        account:=accountNum, _
                        scenario:=scenario, _
                        startBalance:=startDebit, _
                        repaymentTurnover:=creditTurnover, _
                        frozenAmount:=frozenAmount, _
                        endBalance:=endDebit)
                    targetRow = targetRow + 1
                End If
            End If
        End If
NextRow60:
    Next i
    
    ' --- СОРТИРОВКА ПО СУММЕ ---
    lastDataRow = wsTarget.Cells(wsTarget.Rows.count, "I").End(xlUp).row
    If lastDataRow >= 3 Then
        Set sortRange = wsTarget.Range("F3:K" & lastDataRow)
        With wsTarget.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsTarget.Range("I3:I" & lastDataRow), _
                            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .SetRange sortRange
            .Header = xlNo
            .Apply
        End With
    End If
    
    ' --- ГРАНИЦЫ ---
    If lastDataRow >= 2 Then
        With wsTarget.Range("F2:K" & lastDataRow).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(0, 0, 0)
        End With
        With wsTarget.Range("F2:K" & lastDataRow).Borders(xlEdgeLeft): .Weight = xlMedium: End With
        With wsTarget.Range("F2:K" & lastDataRow).Borders(xlEdgeTop): .Weight = xlMedium: End With
        With wsTarget.Range("F2:K" & lastDataRow).Borders(xlEdgeBottom): .Weight = xlMedium: End With
        With wsTarget.Range("F2:K" & lastDataRow).Borders(xlEdgeRight): .Weight = xlMedium: End With
        
        ' Убираем границы и заливку в столбце J
        With wsTarget.Range("J2:J" & lastDataRow)
            .Borders.LineStyle = xlNone
            .Interior.ColorIndex = xlNone
        End With
    End If
    
    wsTarget.Activate
    wsTarget.Range("F2").Select
    MsgBox "Анализ завершен! Найдено записей: " & targetRow - 3, vbInformation, "Готово"
End Sub
' ---------------------------------------------------------
' Функция генерации расширенного комментария с цифрами в миллионах (1 десятичный знак)
' ---------------------------------------------------------
Private Function GenerateComment(debtType As String, account As String, scenario As String, _
                                 startBalance As Double, repaymentTurnover As Double, _
                                 frozenAmount As Double, Optional endBalance As Double = 0) As String
    Dim comment As String
    Dim accountShort As String
    Dim pctRepaid As Double
    
    ' Преобразуем суммы в миллионы с округлением до 1 десятичного знака
    Dim startMln As Double, repayMln As Double, frozenMln As Double
    startMln = Round(startBalance / 1000000, 1)
    repayMln = Round(repaymentTurnover / 1000000, 1)
    frozenMln = Round(frozenAmount / 1000000, 1)
    
    ' Форматируем миллионы: пробелы между разрядами, один десятичный знак
    Dim formattedStartMln As String, formattedRepayMln As String, formattedFrozenMln As String
    formattedStartMln = Format(startMln, "# ##0.0")
    formattedRepayMln = Format(repayMln, "# ##0.0")
    formattedFrozenMln = Format(frozenMln, "# ##0.0")
    
    ' Заменяем точку на запятую в десятичной части (для гарантии)
    formattedStartMln = Replace(formattedStartMln, ".", ",")
    formattedRepayMln = Replace(formattedRepayMln, ".", ",")
    formattedFrozenMln = Replace(formattedFrozenMln, ".", ",")
    
    ' Процент погашения оставляем с одним знаком
    If startBalance > 0 Then
        pctRepaid = (repaymentTurnover / startBalance) * 100
    Else
        pctRepaid = 0
    End If
    
    accountShort = Left(account, 2)
    
    ' --- Блок для 62 счета ---
    If accountShort = "62" Then
        If debtType = "Дебет" Then
            comment = "Покупатель не исполнил обязательства по оплате отгруженных товаров/услуг. "
            If scenario = "чистая" Then
                comment = comment & "Задолженность без движения с начала периода. Сумма долга: " & formattedStartMln & " млн. руб."
            Else
                comment = comment & "Старый долг на начало периода составлял " & formattedStartMln & " млн. руб. "
                comment = comment & "За период поступило оплат на сумму " & formattedRepayMln & " млн. руб. "
                comment = comment & "Это составляет " & Format(pctRepaid, "0.0") & "% от начального долга. "
                comment = comment & "Непогашенный остаток старого долга: " & formattedFrozenMln & " млн. руб."
            End If
        Else ' Кредит по 62 = авансы полученные
            comment = "Не исполнены обязательства перед покупателем по предоплате (аванс полученный). "
            If scenario = "чистая" Then
                comment = comment & "Отгрузки не производились, предоплата зависла. Сумма аванса: " & formattedStartMln & " млн. руб."
            Else
                comment = comment & "Старый аванс на начало периода: " & formattedStartMln & " млн. руб. "
                comment = comment & "За период отгружено товаров/услуг на сумму " & formattedRepayMln & " млн. руб. "
                comment = comment & "Отгружено " & Format(pctRepaid, "0.0") & "% от полученного аванса. "
                comment = comment & "Остаток незакрытого аванса: " & formattedFrozenMln & " млн. руб."
            End If
        End If
    
    ' --- Блок для 60 счета ---
    ElseIf accountShort = "60" Then
        If debtType = "Дебет" Then
            comment = "Поставщик не исполнил обязательства по отгрузке товаров/услуг в счет предоплаты. "
            If scenario = "чистая" Then
                comment = comment & "Аванс без движения, поставок не было. Сумма аванса: " & formattedStartMln & " млн. руб."
            Else
                comment = comment & "Выданный аванс на начало периода: " & formattedStartMln & " млн. руб. "
                comment = comment & "За период поступило товаров/услуг на сумму " & formattedRepayMln & " млн. руб. "
                comment = comment & "Поставлено " & Format(pctRepaid, "0.0") & "% от суммы аванса. "
                comment = comment & "Непогашенный аванс: " & formattedFrozenMln & " млн. руб."
            End If
        Else ' Кредит по 60 = задолженность перед поставщиком
            comment = "Не оплачена задолженность перед поставщиком за поставленные товары/услуги. "
            If scenario = "чистая" Then
                comment = comment & "Долг без движения с начала периода. Сумма долга: " & formattedStartMln & " млн. руб."
            Else
                comment = comment & "Задолженность на начало периода: " & formattedStartMln & " млн. руб. "
                comment = comment & "За период оплачено " & formattedRepayMln & " млн. руб. "
                comment = comment & "Оплачено " & Format(pctRepaid, "0.0") & "% от суммы долга. "
                comment = comment & "Остаток непогашенного долга: " & formattedFrozenMln & " млн. руб."
            End If
        End If
    Else
        ' Прочие счета
        If debtType = "Дебет" Then
            comment = "Контрагент не исполняет обязательства. "
        Else
            comment = "Не исполнены обязательства перед контрагентом. "
        End If
        If scenario = "чистая" Then
            comment = comment & "Отсутствует движение с начала периода. Сумма: " & formattedStartMln & " млн. руб."
        Else
            comment = comment & "Начальное сальдо: " & formattedStartMln & " млн. руб. Погашено: " & formattedRepayMln & " млн. руб. (" & Format(pctRepaid, "0.0") & "%). Остаток: " & formattedFrozenMln & " млн. руб."
        End If
    End If
    
    GenerateComment = comment
End Function

' ---------------------------------------------------------
' Функция проверки наличия букв в строке контрагента
' ---------------------------------------------------------
Private Function ContainsLetter(ByVal txt As String) As Boolean
    Dim i As Integer
    Dim ch As String
    ContainsLetter = False
    For i = 1 To Len(txt)
        ch = Mid(txt, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or _
           (ch >= "А" And ch <= "Я") Or (ch >= "а" And ch <= "я") Or _
           ch = "Ё" Or ch = "ё" Then
            ContainsLetter = True
            Exit Function
        End If
    Next i
End Function


