Attribute VB_Name = "Инфоиззаключки"

Sub ImportDataFromConclusionFile()
    Dim wbThis As Workbook
    Dim wsTarget As Worksheet
    Dim sPath As String
    Dim wbSource As Workbook
    Dim wsSource_System1_2 As Worksheet, wsSource_System3 As Worksheet, wsSource_System4 As Worksheet
    Dim currentFolder As String
    Dim foundFile As String
    Dim searchPattern As String
    Dim fileFormat As String
    Dim hasSystem1_2 As Boolean
    Dim hasSystem3 As Boolean
    Dim hasSystem4 As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Сохраняем текущие настройки Excel
    Dim screenUpdating As Boolean
    Dim calculation As XlCalculation
    Dim enableEvents As Boolean
    
    screenUpdating = Application.screenUpdating
    calculation = Application.calculation
    enableEvents = Application.enableEvents
    
    ' Отключаем обновление экрана и события для ускорения работы
    Application.screenUpdating = False
    Application.calculation = xlCalculationManual
    Application.enableEvents = False
    
    Set wbThis = ThisWorkbook
    
    ' Проверяем существование листа "БКК 3 (2)"
    On Error Resume Next
    Set wsTarget = wbThis.Sheets("Протокол")
    On Error GoTo ErrorHandler
    If wsTarget Is Nothing Then
        MsgBox "Лист 'Протокол' не найден в текущей книге!", vbCritical
        Exit Sub
    End If
    
    ' === ПОИСК ФАЙЛА СО СЛОВОМ "ЗАКЛЮЧЕНИЕ" В НАЗВАНИИ ===
    ' Получаем путь к папке текущего файла
    currentFolder = ThisWorkbook.Path & "\"
    
    ' Ищем файл, содержащий слово "Заключение" в названии
    searchPattern = "*Заключение*"
    
    ' Сначала ищем .xlsm файлы
    foundFile = Dir(currentFolder & searchPattern & ".xlsm")
    
    ' Если не найдено .xlsm файлов, ищем другие форматы Excel
    If foundFile = "" Then
        foundFile = Dir(currentFolder & searchPattern & ".xlsx")
        If foundFile = "" Then
            foundFile = Dir(currentFolder & searchPattern & ".xls")
        End If
    End If
    
    ' Проверка наличия файла
    If foundFile = "" Then
        MsgBox "Файл, содержащий 'Заключение' в названии, не найден в папке: " & currentFolder, vbCritical
        Exit Sub
    End If
    
    ' Полный путь к найденному файлу
    sPath = currentFolder & foundFile
    ' === КОНЕЦ БЛОКА ПОИСКА ФАЙЛА ===
    
    ' Открытие файла источника
    Set wbSource = Workbooks.Open(sPath, ReadOnly:=True)
    
    ' Проверка существования листов в файле-источнике
    On Error Resume Next
    Set wsSource_System1_2 = wbSource.Sheets("Система 1-2")
    Set wsSource_System3 = wbSource.Sheets("Система 3")
    Set wsSource_System4 = wbSource.Sheets("Система4")
    On Error GoTo ErrorHandler
    
    ' Проверяем, какой лист присутствует в файле-источнике
    hasSystem1_2 = Not (wsSource_System1_2 Is Nothing)
    hasSystem3 = Not (wsSource_System3 Is Nothing)
    hasSystem4 = Not (wsSource_System4 Is Nothing)
    
    ' Если ни один из нужных листов не найден
    If Not hasSystem1_2 And Not hasSystem3 And Not hasSystem4 Then
        MsgBox "В файле-источнике не найдено ни листа 'Система 1-2', ни листа 'Система 3', ни листа 'Система4'", vbCritical
        GoTo Cleanup
    End If
    
    ' === ОБРАБОТКА ДАННЫХ В ЗАВИСИМОСТИ ОТ НАЛИЧИЯ ЛИСТОВ ===
    If hasSystem1_2 Then
        ' === ЗАПОЛНЕНИЕ ПО ФОРМУЛАМ ИЗ FormulasList_МКК_1-2.txt ===
        
        ' Ячейка D3
        wsTarget.Range("D5").Value = Date
        
        ' Ячейка A7
    Dim formattedValue1 As String
    formattedValue1 = Format(wsSource_System1_2.Range("B12").Value, "### ### ###")
    wsTarget.Range("A10").Value = "Предоставить лизинговое финансирование по сделке " & wsSource_System1_2.Range("B4").Value & _
                             " ИНН " & wsSource_System1_2.Range("B20").Value & _
                             " на сумму " & formattedValue1 & _
                             " (" & пропись(wsSource_System1_2.Range("B12").Value) & ") " & _
                             "с целью приобретения " & wsSource_System1_2.Range("B33").Value & _
                             " стоимостью " & Format(wsSource_System1_2.Range("B34").Value, "### ### ###") & " рублей"
        
        ' Ячейка A9
        wsTarget.Range("A12").Value = wsSource_System1_2.Range("B131").Value & ", " & wsSource_System1_2.Range("E5").Value
        
        ' Ячейка A11
        wsTarget.Range("A14").Value = "Отказать в лизинговом финансировании по сделке " & wsSource_System1_2.Range("B4").Value & _
                                      " ИНН: " & wsSource_System1_2.Range("B20").Value & " "
        
        ' Ячейка A12
        wsTarget.Range("A15").Value = "Одобрить лизинговое финансирование по сделке " & wsSource_System1_2.Range("B4").Value & _
                                      " ИНН " & wsSource_System1_2.Range("B20").Value & " на параметрах:"
        
        ' Ячейка B13
        wsTarget.Range("B16").Value = wsSource_System1_2.Range("B7").Value
        
        ' Ячейка B14
        If UCase(wsSource_System1_2.Range("B105").Value) = "ДА" Then
            wsTarget.Range("B17").Value = "Повторный"
        Else
            wsTarget.Range("B17").Value = "Новый"
        End If
        
        ' Ячейка B15
        wsTarget.Range("B18").Value = wsSource_System1_2.Range("B33").Value
        
        ' Ячейка B16
        wsTarget.Range("B19").Value = wsSource_System1_2.Range("B35").Value & " ед."
        
        ' Ячейка B17
        Dim totalCost As Double
        totalCost = wsSource_System1_2.Range("B34").Value + wsSource_System1_2.Range("E34").Value
        wsTarget.Range("B20").Value = Format(totalCost, "### ### ###") & " рублей"
        
        ' Ячейка B18
        wsTarget.Range("B21").Value = Format(wsSource_System1_2.Range("B37").Value, "### ### ###") & " рублей"
        
        ' Ячейка B19
        wsTarget.Range("B22").Value = wsSource_System1_2.Range("B36").Value
        
        ' Ячейка B20
        wsTarget.Range("B23").Value = wsSource_System1_2.Range("B38").Value & " мес."
        
        ' Ячейка B21
        wsTarget.Range("B24").Value = wsSource_System1_2.Range("B44").Value
        
        ' Ячейка B22
        wsTarget.Range("B25").Value = wsSource_System1_2.Range("B46").Value
        
        ' Ячейка B23
        wsTarget.Range("B26").Value = wsSource_System1_2.Range("B41").Value
        
        ' Ячейка B24
        wsTarget.Range("B27").Value = wsSource_System1_2.Range("B42").Value
        
        ' Ячейка B25
        wsTarget.Range("B28").Value = wsSource_System1_2.Range("B45").Value
        
        ' Ячейка B29
        wsTarget.Range("B32").Value = wsSource_System1_2.Range("B52").Value & ", ИНН:" & wsSource_System1_2.Range("B53").Value & ", статус: " & wsSource_System1_2.Range("B54").Value
        
        ' Ячейка B30
        wsTarget.Range("B33").Value = "ПЛ " & wsSource_System1_2.Range("E53").Value
        
    ElseIf hasSystem3 Then
        ' === ЗАПОЛНЕНИЕ ПО ФОРМУЛАМ ИЗ FormulasList_БКК_3.txt ===
        
        ' Ячейка D4
        wsTarget.Range("D5").Value = Date
        
        ' Ячейка A8
    Dim formattedValue2 As String
    formattedValue2 = Format(wsSource_System3.Range("C18").Value, "### ### ###")
    wsTarget.Range("A10").Value = "Предоставить лизинговое финансирование по сделке " & wsSource_System3.Range("B2").Value & _
                             " ИНН " & wsSource_System3.Range("C73").Value & _
                             " на сумму " & formattedValue2 & _
                             " (" & пропись(wsSource_System3.Range("C18").Value) & ") " & _
                             "с целью приобретения " & wsSource_System3.Range("C20").Value & _
                             ", " & wsSource_System3.Range("C31").Value & _
                             ", " & wsSource_System3.Range("C42").Value & ", " & wsSource_System3.Range("C53").Value
        
        ' Ячейка A10
        wsTarget.Range("A12").Value = wsSource_System3.Range("G176").Value & ", " & wsSource_System3.Range("G3").Value
        
        ' Ячейка A12
        wsTarget.Range("A14").Value = "Отказать в лизинговом финансировании по сделке " & wsSource_System3.Range("B2").Value & _
                                      " ИНН: " & wsSource_System3.Range("C73").Value & " "
        
        ' Ячейка A13
        wsTarget.Range("A15").Value = "Одобрить лизинговое финансирование по сделке " & wsSource_System3.Range("B2").Value & _
                                      " ИНН " & wsSource_System3.Range("C73").Value & " на параметрах:"
        
        ' Ячейка B14
        wsTarget.Range("B16").Value = wsSource_System3.Range("B5").Value
        
        ' Ячейка B15
        If UCase(wsSource_System3.Range("D161").Value) = "ДА" Then
            wsTarget.Range("B17").Value = "Повторный"
        Else
            wsTarget.Range("B17").Value = "Новый"
        End If
        
        ' Ячейка B16
        wsTarget.Range("B18").Value = wsSource_System3.Range("C20").Value & ", " & wsSource_System3.Range("C31").Value & _
                                      ", " & wsSource_System3.Range("C42").Value & ", " & wsSource_System3.Range("C53").Value
        
        ' Ячейка B17
        wsTarget.Range("B19").Value = wsSource_System3.Range("C19").Value & " ед."
        
        ' Ячейка B18
        wsTarget.Range("B20").Value = Format(wsSource_System3.Range("C17").Value, "### ### ###") & " рублей"
        
        ' Ячейка B19
        wsTarget.Range("B21").Value = Format(wsSource_System3.Range("C18").Value, "### ### ###") & " рублей"
        
        ' Ячейка B20
        wsTarget.Range("B22").Value = wsSource_System3.Range("C21").Value
        
        ' Ячейка B21
        wsTarget.Range("B23").Value = wsSource_System3.Range("C22").Value & " мес."
        
        ' Ячейка B22
        wsTarget.Range("B24").Value = wsSource_System3.Range("C25").Value
        
        ' Ячейка B23
        wsTarget.Range("B25").Value = wsSource_System3.Range("C27").Value
        
        ' Ячейка B24
        wsTarget.Range("B26").Value = wsSource_System3.Range("C65").Value
        
        ' Ячейка B25
        wsTarget.Range("B27").Value = wsSource_System3.Range("C66").Value
        
        ' Ячейка B26
        wsTarget.Range("B28").Value = wsSource_System3.Range("C26").Value
        
        ' Ячейка B30
        wsTarget.Range("B32").Value = wsSource_System3.Range("C91").Value & ", ИНН:" & wsSource_System3.Range("C92").Value & ", статус: " & wsSource_System3.Range("C93").Value
        
        ' Ячейка B31
        wsTarget.Range("B33").Value = "ПЛ " & wsSource_System3.Range("H92").Value
    
    ElseIf hasSystem4 Then
        ' === Заполняем с файла рейтинга ===
        
        ' Ячейка D5
        wsTarget.Range("D5").Value = Date
        
        ' Ячейка A10
    Dim formattedValue3 As String
    formattedValue3 = Format(wsSource_System4.Range("B8").Value, "### ### ###")
    wsTarget.Range("A10").Value = "Предоставить лизинговое финансирование по сделке " & wsSource_System4.Range("A2").Value & _
                             " на сумму " & formattedValue3 & _
                             " (" & пропись(wsSource_System4.Range("B8").Value) & ") " & _
                             "с целью приобретения " & wsSource_System4.Range("B6").Value & _
                             ", стоимостью " & Format(wsSource_System4.Range("B7").Value, "### ### ###") & " руб."
        
        ' Ячейка A12
        wsTarget.Range("A12").Value = wsSource_System4.Range("K2").Value & ", " & wsSource_System4.Range("J2").Value
        
        ' Ячейка A14
        wsTarget.Range("A14").Value = "Отказать в лизинговом финансировании по сделке " & wsSource_System4.Range("A2").Value
        
        ' Ячейка A15
        wsTarget.Range("A15").Value = "Одобрить лизинговое финансирование по сделке " & wsSource_System4.Range("A2").Value & _
                                      " на параметрах:"
                                      
        ' Ячейка B16
        wsTarget.Range("B16").Value = wsSource_System4.Range("B5").Value
        
        'Ячейка B17 --- ВАЖНО: Скрываем окно другого файла ---
        If Not wsSource_System4.Parent Is ThisWorkbook Then
            wsSource_System4.Parent.Windows(1).Visible = False
        End If

        ' --- Активируем наше окно "Протокол" ---
        ThisWorkbook.Activate
        Application.Windows(ThisWorkbook.name).Visible = True
        Application.Windows(ThisWorkbook.name).WindowState = xlMaximized

        ' --- Вызываем форму ---
        Dim frm As frmClientType
        Set frm = New frmClientType

        ' Показываем форму
        frm.Show

        ' Получаем выбранное значение
        If frm.SelectedValue = "Не выбрано" Then
            ' Если пользователь ничего не выбрал (закрыл форму)
            wsTarget.Range("B17").Value = "Не выбран"
        Else
            ' Записываем выбранное значение
            wsTarget.Range("B17").Value = frm.SelectedValue
        End If

        ' Очищаем память
        Unload frm
        Set frm = Nothing

        ' --- Возвращаем видимость окну другого файла (если нужно) ---
        'If Not wsSource_System4.Parent Is ThisWorkbook Then
            'wsSource_System4.Parent.Windows(1).Visible = True
        'End If

        ' Ячейка B18
        wsTarget.Range("B18").Value = wsSource_System4.Range("B6").Value
        
        ' Ячейка B19
        wsTarget.Range("B19").Value = wsSource_System4.Range("B6").Value & " ед."
        
        ' Ячейка B20
        wsTarget.Range("B20").Value = Format(wsSource_System4.Range("B7").Value, "### ### ###") & " руб."
        
        ' Ячейка B21
        wsTarget.Range("B21").Value = Format(wsSource_System4.Range("B8").Value, "### ### ###") & " руб."
        
        ' Ячейка B22
        wsTarget.Range("B22").Value = wsSource_System4.Range("B10").Value
        
        ' Ячейка B23
        wsTarget.Range("B23").Value = wsSource_System4.Range("B11").Value
        
        ' Ячейка B24
        wsTarget.Range("B24").Value = wsSource_System4.Range("B18").Value
        
        ' Ячейка B25
        wsTarget.Range("B25").Value = wsSource_System4.Range("B13").Value
        
        ' Ячейка B26 и B27
        Dim sourceValue As String
        sourceValue = wsSource_System4.Range("B17").Value

        ' Ячейка B26 и B27
        Dim textPart As String
        Dim percentPart As String

        Call SplitTextFromPercent(wsSource_System4.Range("B17").Value, textPart, percentPart)

        wsTarget.Range("B26").Value = textPart     ' "ИП Брат"
        wsTarget.Range("B27").Value = percentPart  ' "4%"
        
        ' Ячейка B28
        wsTarget.Range("B28").Value = wsSource_System4.Range("B14").Value
        
        ' Ячейка B29
        wsTarget.Range("B29").Value = wsSource_System4.Range("B15").Value
        
        ' Ячейка B30
        wsTarget.Range("B30").Value = wsSource_System4.Range("B22").Value
        
        ' Ячейка B31
        wsTarget.Range("B31").Value = wsSource_System4.Range("B16").Value
        
        ' Ячейка B32
        wsTarget.Range("B32").Value = wsSource_System4.Range("B27").Value
        
        ' Ячейка B33
        wsTarget.Range("B33").Value = wsSource_System4.Range("B28").Value
        
    End If
    
    ' Закрытие файла источника
    wbSource.Close SaveChanges:=False
    
      ' Вызываем функцию для заполнения участников собрания
    Call FillMeetingParticipants
    
    ' Восстановление настроек Excel
    Application.screenUpdating = screenUpdating
    Application.calculation = calculation
    Application.enableEvents = enableEvents
    
    MsgBox "Данные успешно загружены!", vbInformation
    Exit Sub

ErrorHandler:
    ' Восстановление настроек Excel
    Application.screenUpdating = screenUpdating
    Application.calculation = calculation
    Application.enableEvents = enableEvents
    
    If Err.number = 9 Then
        MsgBox "Ошибка: " & Err.Description & vbCrLf & _
               "Вероятно, лист не найден в файле-источнике." & vbCrLf & _
               "Проверьте, что в файле " & sPath & " действительно есть необходимые листы", vbCritical
    ElseIf Err.number = 1004 Then
        MsgBox "Ошибка доступа к ячейке. Проверьте, что файл-источник открыт и содержит необходимые данные.", vbCritical
    Else
        MsgBox "Произошла ошибка " & Err.number & ": " & Err.Description, vbCritical
    End If

Cleanup:
    If Not wbSource Is Nothing Then
        On Error Resume Next
        wbSource.Close SaveChanges:=False
    End If
    Application.screenUpdating = screenUpdating
    Application.calculation = calculation
    Application.enableEvents = enableEvents
End Sub

Sub FillMeetingParticipants()
    Dim wsTarget As Worksheet
    Dim targetCell As Range
    Dim resultText As String
    Dim i As Long
    Dim fullName As String
    Dim parts() As String
    Dim lastName As String
    Dim initials As String
    Dim j As Long
    Dim isBKK As Boolean
    Dim isFirstParticipant As Boolean
    
    Set wsTarget = ThisWorkbook.Sheets("Протокол")
    
    ' Определяем, какой комитет выбран
    If UCase(Trim(wsTarget.Range("E1").Value)) = "БКК" Then
        isBKK = True
        Set targetCell = wsTarget.Range("A7")
        ' Показываем A7 и скрываем A8
        wsTarget.Rows(7).Hidden = False
        wsTarget.Rows(8).Hidden = True
    ElseIf UCase(Trim(wsTarget.Range("E1").Value)) = "МКК" Then
        isBKK = False
        Set targetCell = wsTarget.Range("A8")
        ' Показываем A8 и скрываем A7
        wsTarget.Rows(7).Hidden = True
        wsTarget.Rows(8).Hidden = False
    Else
        ' Если не указан тип комитета, выходим
        Exit Sub
    End If
    
    ' Проверяем, открыты ли нужные строки и собираем данные
    resultText = ""
    isFirstParticipant = True
    
    For i = 39 To 45
        If Not wsTarget.Rows(i).Hidden Then
            If wsTarget.Cells(i, "D").Value <> "" Then ' Колонка D
                fullName = Trim(wsTarget.Cells(i, "D").Value)
                
                ' Проверяем, есть ли точка в строке (признак инициалов)
                If InStr(fullName, ".") > 0 Then
                    ' Разделяем на части
                    parts = Split(fullName, " ")
                    
                    If UBound(parts) >= 1 Then
                        ' Ищем фамилию (обычно последняя часть)
                        lastName = parts(UBound(parts))
                        initials = ""
                        
                        ' Собираем инициалы из первых частей
                        For j = 0 To UBound(parts) - 1
                            If Len(Trim(parts(j))) > 0 Then
                                initials = initials & " " & parts(j)
                            End If
                        Next j
                        
                        ' Форматируем: Фамилия И.О.
                        Dim formattedName As String
                        formattedName = lastName & " " & Trim(initials)
                        
                        ' Если это первый участник - добавляем " – председатель"
                        If isFirstParticipant Then
                            formattedName = formattedName & " – председатель"
                            isFirstParticipant = False
                        End If
                        
                        resultText = resultText & formattedName & vbCrLf
                    Else
                        ' Если формат другой, оставляем как есть
                        resultText = resultText & fullName & vbCrLf
                    End If
                Else
                    ' Если формат другой, оставляем как есть
                    resultText = resultText & fullName & vbCrLf
                End If
            End If
        End If
    Next i
    
    ' Убираем последний перенос строки
    If Len(resultText) > 0 Then
        resultText = Left(resultText, Len(resultText) - 1)
    End If
    
    ' Записываем результат в целевую ячейку
    targetCell.Value = resultText
End Sub

' Простая функция для преобразования формата имени
Function FormatNameWithInitials(fullName As String) As String
    Dim parts() As String
    Dim i As Long
    Dim lastName As String
    Dim initials As String
    
    ' Убираем лишние пробелы
    fullName = Trim(fullName)
    
    If fullName = "" Then
        FormatNameWithInitials = ""
        Exit Function
    End If
    
    ' Разделяем на части по пробелам
    parts = Split(fullName, " ")
    
    ' Убираем пустые элементы
    Dim cleanParts() As String
    ReDim cleanParts(0 To 0)
    Dim cleanIndex As Long
    cleanIndex = 0
    
    For i = 0 To UBound(parts)
        If Trim(parts(i)) <> "" Then
            cleanParts(cleanIndex) = Trim(parts(i))
            If i < UBound(parts) Then
                cleanIndex = cleanIndex + 1
                ReDim Preserve cleanParts(0 To cleanIndex)
            End If
        End If
    Next i
    
    ' Если только одна часть - возвращаем как есть
    If UBound(cleanParts) = 0 Then
        FormatNameWithInitials = cleanParts(0)
        Exit Function
    End If
    
    ' Берем фамилию (последний элемент)
    lastName = cleanParts(UBound(cleanParts))
    
    ' Собираем инициалы (все элементы кроме последнего)
    initials = ""
    For i = 0 To UBound(cleanParts) - 1
        If initials <> "" Then initials = initials & " "
        initials = initials & cleanParts(i)
    Next i
    
    ' Формат: Фамилия И.О.
    FormatNameWithInitials = lastName & " " & initials
End Function

' Функция для преобразования числа в текст (упрощенная версия)
Function ConvertNumberToText(number As Double) As String
    ' Это упрощенная версия, в реальной реализации нужно использовать полноценный алгоритм
    ' преобразования чисел в текст на русском языке
    
    ' Для примера просто возвращаем округленное число
    ConvertNumberToText = Format(number, "0")
End Function

Function GetSheetNames(wb As Workbook) As String
    Dim ws As Worksheet
    Dim sheetNames As String
    
    For Each ws In wb.Worksheets
        If sheetNames <> "" Then sheetNames = sheetNames & ", "
        sheetNames = sheetNames & ws.name
    Next ws
    
    GetSheetNames = sheetNames
End Function

Function SplitTextFromPercent(inputString As String, ByRef textPart As String, ByRef percentPart As String)
    Dim tempStr As String
    Dim i As Long
    Dim hasDigit As Boolean
    
    tempStr = Trim(inputString)
    textPart = tempStr
    percentPart = ""
    
    ' Идем с конца строки
    For i = Len(tempStr) To 1 Step -1
        Dim char As String
        char = Mid(tempStr, i, 1)
        
        ' Если нашли цифру, точку, запятую или знак процента
        If IsNumeric(char) Or char = "." Or char = "," Or char = "%" Then
            hasDigit = True
        ElseIf hasDigit Then
            ' Нашли начало числовой части
            textPart = Trim(Left(tempStr, i))
            percentPart = Trim(Mid(tempStr, i + 1))
            Exit Function
        End If
    Next i
    
    ' Если вся строка состоит из чисел/процентов
    If hasDigit Then
        textPart = ""
        percentPart = tempStr
    End If
End Function



