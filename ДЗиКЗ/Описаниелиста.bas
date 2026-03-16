Attribute VB_Name = "Module1"

' ====================================================================
' МОДУЛЬ: Описание листа Excel
' ВЕРСИЯ: 1.0
' АВТОР: AI Assistant
' НАЗНАЧЕНИЕ: Создает подробный текстовый файл с описанием выбранного листа
' ====================================================================

Option Explicit

' ----------------------------------------------------------------
' ОСНОВНАЯ ПРОЦЕДУРА - ТОЧКА ВХОДА
' ----------------------------------------------------------------
Sub СоздатьОписаниеЛиста()
    Dim ws As Worksheet
    Dim selectedSheetName As String
    Dim proceed As Boolean
    
    ' Показать диалоговое окно для выбора листа
    selectedSheetName = ПоказатьДиалогВыбораЛиста()
    
    ' Если пользователь отменил выбор
    If selectedSheetName = "" Then
        MsgBox "Операция отменена пользователем.", vbInformation, "Отмена"
        Exit Sub
    End If
    
    ' Получить ссылку на выбранный лист
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(selectedSheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Лист '" & selectedSheetName & "' не найден!", vbCritical, "Ошибка"
        Exit Sub
    End If
    
    ' Подтверждение действия
    proceed = ПоказатьДиалогПодтверждения(ws.Name)
    If Not proceed Then
        MsgBox "Операция отменена.", vbInformation, "Отмена"
        Exit Sub
    End If
    
    ' Создать описание листа
    Call СоздатьТекстовыйФайлСОписанием(ws)
    
    ' Завершение
    MsgBox "Текстовый файл с описанием листа '" & ws.Name & "' успешно создан!", _
           vbInformation, "Готово"
End Sub

' ----------------------------------------------------------------
' ФУНКЦИЯ: Диалог выбора листа
' ----------------------------------------------------------------
Private Function ПоказатьДиалогВыбораЛиста() As String
    Dim ws As Worksheet
    Dim i As Long
    Dim selectedIndex As Long
    
    ' Создание пользовательской формы (можно заменить на форму UserForm)
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    ' Создаем временный лист для выбора
    Dim tempSheet As Worksheet
    On Error Resume Next
    Set tempSheet = ThisWorkbook.Sheets("TempSelector")
    On Error GoTo 0
    
    If Not tempSheet Is Nothing Then
        Application.DisplayAlerts = False
        tempSheet.Delete
        Application.DisplayAlerts = True
    End If
    
    ' Создаем временный лист
    Set tempSheet = ThisWorkbook.Sheets.Add
    tempSheet.Name = "TempSelector"
    
    ' Заполняем список листов
    tempSheet.Cells(1, 1).Value = "№"
    tempSheet.Cells(1, 2).Value = "Имя листа"
    tempSheet.Cells(1, 3).Value = "Видимость"
    
    i = 2
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "TempSelector" Then
            tempSheet.Cells(i, 1).Value = i - 1
            tempSheet.Cells(i, 2).Value = ws.Name
            tempSheet.Cells(i, 3).Value = IIf(ws.Visible = xlSheetVisible, "Видимый", "Скрытый")
            i = i + 1
        End If
    Next ws
    
    ' Настройка форматирования
    With tempSheet
        .usedRange.Borders.LineStyle = xlContinuous
        .usedRange.Font.Bold = True
        .Columns("A:C").AutoFit
        .Range("A1:C1").Interior.color = RGB(200, 220, 240)
        
        ' Создаем именованный диапазон для списка
        .Names.Add Name:="SheetList", RefersTo:=.Range("B2:B" & i - 1)
    End With
    
    ' Показываем диалоговое окно ввода
    Dim result As Variant
    Dim inputPrompt As String
    Dim inputTitle As String
    
    inputPrompt = "Выберите номер листа для описания:" & vbNewLine & vbNewLine
    For i = 2 To tempSheet.Cells(tempSheet.Rows.count, "B").End(xlUp).row
        inputPrompt = inputPrompt & tempSheet.Cells(i, 1).Value & ". " & _
                      tempSheet.Cells(i, 2).Value & _
                      " (" & tempSheet.Cells(i, 3).Value & ")" & vbNewLine
    Next i
    
    inputPrompt = inputPrompt & vbNewLine & "Введите номер листа:"
    inputTitle = "Выбор листа для анализа"
    
    ' Цикл для повторного запроса при ошибке
    Do
        result = Application.InputBox( _
            Prompt:=inputPrompt, _
            Title:=inputTitle, _
            Default:="1", _
            Type:=1) ' Type 1 = число
            
        If result = False Then ' Пользователь нажал Cancel
            ПоказатьДиалогВыбораЛиста = ""
            Exit Do
        ElseIf IsNumeric(result) Then
            selectedIndex = CLng(result)
            
            ' Проверка корректности номера
            If selectedIndex >= 1 And selectedIndex <= tempSheet.Cells(tempSheet.Rows.count, "B").End(xlUp).row - 1 Then
                ПоказатьДиалогВыбораЛиста = tempSheet.Cells(selectedIndex + 1, 2).Value
                Exit Do
            Else
                MsgBox "Пожалуйста, введите число от 1 до " & _
                       tempSheet.Cells(tempSheet.Rows.count, "B").End(xlUp).row - 1, _
                       vbExclamation, "Неверный номер"
            End If
        Else
            MsgBox "Пожалуйста, введите число!", vbExclamation, "Ошибка ввода"
        End If
    Loop
    
    ' Удаление временного листа
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = True
    
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Function

' ----------------------------------------------------------------
' ФУНКЦИЯ: Диалог подтверждения
' ----------------------------------------------------------------
Private Function ПоказатьДиалогПодтверждения(sheetName As String) As Boolean
    Dim result As VbMsgBoxResult
    
    result = MsgBox( _
        "Вы собираетесь создать подробное описание листа:" & vbNewLine & _
        "'" & sheetName & "'" & vbNewLine & vbNewLine & _
        "Это может занять некоторое время в зависимости от размера листа." & vbNewLine & vbNewLine & _
        "Продолжить?", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Подтверждение создания описания")
    
    ПоказатьДиалогПодтверждения = (result = vbYes)
End Function

' ----------------------------------------------------------------
' ОСНОВНАЯ ПРОЦЕДУРА: Создание текстового файла
' ----------------------------------------------------------------
Private Sub СоздатьТекстовыйФайлСОписанием(ws As Worksheet)
    Dim fso As Object
    Dim txtFile As Object
    Dim filePath As String
    Dim startTime As Double
    Dim usedRange As Range
    Dim cell As Range
    Dim row As Range
    Dim col As Range
    Dim i As Long, j As Long
    Dim lastRow As Long, lastCol As Long
    Dim formulaCount As Long, valueCount As Long, emptyCount As Long
    Dim namedRange As Name
    Dim shape As shape
    Dim chart As ChartObject
    Dim pivot As PivotTable
    
    startTime = Timer
    
    ' Определение пути для сохранения файла
    filePath = ThisWorkbook.Path & "\Описание_листа_" & ws.Name & "_" & _
               Format(Now, "yyyy-mm-dd_HH-mm") & ".txt"
    
    ' Создание файловой системы объекта
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(filePath, True, True) ' Unicode encoding
    
    ' Запись заголовка
    Call ЗаписатьЗаголовокФайла(txtFile, ws)
    
    ' Получение используемого диапазона
    On Error Resume Next
    Set usedRange = ws.usedRange
    On Error GoTo 0
    
    If usedRange Is Nothing Then
        txtFile.WriteLine "Лист не содержит данных."
        txtFile.Close
        Exit Sub
    End If
    
    lastRow = usedRange.Rows.count + usedRange.row - 1
    lastCol = usedRange.Columns.count + usedRange.Column - 1
    
    ' ОБЩАЯ ИНФОРМАЦИЯ О ЛИСТЕ
    Call ЗаписатьОбщуюИнформацию(txtFile, ws, lastRow, lastCol)
    
    ' СОДЕРЖИМОЕ ЯЧЕЕК (подробно)
    Call ЗаписатьСодержимоеЯчеек(txtFile, ws, usedRange, lastRow, lastCol, _
                                  formulaCount, valueCount, emptyCount)
    
    ' ФОРМУЛЫ И РАСЧЕТЫ
    Call ЗаписатьФормулыИРасчеты(txtFile, ws, usedRange, formulaCount)
    
    ' ИМЕНОВАННЫЕ ДИАПАЗОНЫ
    Call ЗаписатьИменованныеДиапазоны(txtFile, ws)
    
    ' ГРАФИЧЕСКИЕ ОБЪЕКТЫ
    Call ЗаписатьГрафическиеОбъекты(txtFile, ws)
    
    ' СВОДНЫЕ ТАБЛИЦЫ
    Call ЗаписатьСводныеТаблицы(txtFile, ws)
    
    ' НАСТРОЙКИ ЛИСТА
    Call ЗаписатьНастройкиЛиста(txtFile, ws)
    
    ' СТАТИСТИКА
    Call ЗаписатьСтатистику(txtFile, ws, lastRow, lastCol, formulaCount, _
                            valueCount, emptyCount, startTime)
    
    ' Закрытие файла
    txtFile.Close
    
    ' Открытие файла в блокноте
    Shell "notepad.exe " & Chr(34) & filePath & Chr(34), vbNormalFocus
    
    ' Очистка
    Set fso = Nothing
    Set txtFile = Nothing
End Sub

' ----------------------------------------------------------------
' ПОДПРОЦЕДУРА: Запись заголовка файла
' ----------------------------------------------------------------
Private Sub ЗаписатьЗаголовокФайла(txtFile As Object, ws As Worksheet)
    With txtFile
        .WriteLine "=" & String(100, "=")
        .WriteLine "ПОДРОБНОЕ ОПИСАНИЕ ЛИСТА EXCEL"
        .WriteLine "=" & String(100, "=")
        .WriteLine "Файл: " & ThisWorkbook.Name
        .WriteLine "Лист: " & ws.Name
        .WriteLine "Дата создания: " & Format(Now, "dd.mm.yyyy HH:mm:ss")
        .WriteLine "Автор: " & Application.UserName
        .WriteLine String(100, "-")
        .WriteLine ""
    End With
End Sub

' ----------------------------------------------------------------
' ПОДПРОЦЕДУРА: Общая информация о листе
' ----------------------------------------------------------------
Private Sub ЗаписатьОбщуюИнформацию(txtFile As Object, ws As Worksheet, _
                                     lastRow As Long, lastCol As Long)
    With txtFile
        .WriteLine "1. ОБЩАЯ ИНФОРМАЦИЯ О ЛИСТЕ"
        .WriteLine String(80, "-")
        .WriteLine "Имя листа: " & ws.Name
        .WriteLine "Индекс листа: " & ws.Index
        .WriteLine "Видимость: " & ПолучитьВидимостьЛиста(ws)
        .WriteLine "Защита: " & IIf(ws.ProtectContents, "Защищен", "Не защищен")
        .WriteLine "Цвет ярлычка: " & ПолучитьЦветЯрлычка(ws)
        .WriteLine ""
        .WriteLine "Размеры используемого диапазона:"
        .WriteLine "  • Строк: " & lastRow & " (до строки " & lastRow & ")"
        .WriteLine "  • Столбцов: " & lastCol & " (до столбца " & ПолучитьИмяСтолбца(lastCol) & ")"
        .WriteLine "  • Ячеек: " & lastRow * lastCol
        .WriteLine "  • Диапазон: A1:" & ПолучитьИмяСтолбца(lastCol) & lastRow
        .WriteLine ""
    End With
End Sub

' ----------------------------------------------------------------
' ПОДПРОЦЕДУРА: Содержимое ячеек
' ----------------------------------------------------------------
Private Sub ЗаписатьСодержимоеЯчеек(txtFile As Object, ws As Worksheet, _
                                    usedRange As Range, lastRow As Long, _
                                    lastCol As Long, ByRef formulaCount As Long, _
                                    ByRef valueCount As Long, ByRef emptyCount As Long)
    Dim cell As Range
    Dim row As Range
    Dim i As Long, j As Long
    Dim hasData As Boolean
    
    With txtFile
        .WriteLine "2. ПОДРОБНОЕ СОДЕРЖАНИЕ ЯЧЕЕК"
        .WriteLine String(80, "-")
        .WriteLine "Формат: [Строка,Столбец] Адрес | Тип | Значение | Формат | Примечания"
        .WriteLine String(80, "-")
        
        ' Проход по всем строкам и столбцам используемого диапазона
        For i = 1 To lastRow
            hasData = False
            Dim rowContent As String
            rowContent = ""
            
            For j = 1 To lastCol
                Set cell = ws.Cells(i, j)
                
                ' Проверяем, есть ли что-то в ячейке
                If Not IsEmpty(cell) Or cell.HasFormula Or _
                   Not IsEmpty(cell.comment) Or cell.FormatConditions.count > 0 Then
                    
                    hasData = True
                    rowContent = rowContent & ЗаписатьПодробностиЯчейки(cell, i, j)
                    
                    ' Статистика
                    If cell.HasFormula Then
                        formulaCount = formulaCount + 1
                    ElseIf Not IsEmpty(cell) Then
                        valueCount = valueCount + 1
                    Else
                        emptyCount = emptyCount + 1
                    End If
                Else
                    emptyCount = emptyCount + 1
                End If
            Next j
            
            ' Записываем строку, если в ней есть данные
            If hasData Then
                .WriteLine "=== СТРОКА " & i & " ==="
                .WriteLine rowContent
                .WriteLine ""
            End If
        Next i
    End With
End Sub

' ----------------------------------------------------------------
' ФУНКЦИЯ: Подробности ячейки
' ----------------------------------------------------------------
Private Function ЗаписатьПодробностиЯчейки(cell As Range, rowNum As Long, colNum As Long) As String
    Dim details As String
    Dim cellAddr As String
    Dim cellType As String
    Dim cellValue As String
    Dim cellFormat As String
    Dim cellNotes As String
    
    ' Адрес ячейки
    cellAddr = "[" & rowNum & "," & colNum & "] " & cell.Address(False, False)
    
    ' Тип ячейки
    If cell.HasFormula Then
        cellType = "ФОРМУЛА"
    ElseIf IsDate(cell.Value) Then
        cellType = "ДАТА"
    ElseIf IsNumeric(cell.Value) Then
        cellType = "ЧИСЛО"
    ElseIf IsEmpty(cell) Then
        cellType = "ПУСТО"
    Else
        cellType = "ТЕКСТ"
    End If
    
    ' Значение ячейки
    If cell.HasFormula Then
        cellValue = "Формула: " & cell.Formula
        If cell.Value <> "" Then
            cellValue = cellValue & " = " & CStr(cell.Value)
        End If
    ElseIf IsError(cell.Value) Then
        cellValue = "ОШИБКА: " & CStr(cell.Value)
    ElseIf IsDate(cell.Value) Then
        cellValue = Format(cell.Value, "dd.mm.yyyy")
        If TimeValue(cell.Value) <> 0 Then
            cellValue = cellValue & " " & Format(cell.Value, "HH:mm:ss")
        End If
    Else
        cellValue = CStr(cell.Value)
    End If
    
    ' Формат ячейки
    cellFormat = cell.NumberFormat
    If cellFormat = "General" Then cellFormat = "Общий"
    
    ' Примечания и особенности
    cellNotes = ""
    
    ' Примечания
    If Not cell.comment Is Nothing Then
        cellNotes = cellNotes & "Примечание: " & cell.comment.text & "; "
    End If
    
    ' Условное форматирование
    If cell.FormatConditions.count > 0 Then
        cellNotes = cellNotes & "Условное форматирование (" & _
                   cell.FormatConditions.count & " правил); "
    End If
    
    ' Проверка данных
    If Not cell.Validation Is Nothing Then
        cellNotes = cellNotes & "Проверка данных; "
    End If
    
    ' Гиперссылка
    If Not cell.Hyperlinks Is Nothing Then
        If cell.Hyperlinks.count > 0 Then
            cellNotes = cellNotes & "Гиперссылка: " & cell.Hyperlinks(1).Address & "; "
        End If
    End If
    
    ' Стиль
    If cell.Style <> "Normal" Then
        cellNotes = cellNotes & "Стиль: " & cell.Style & "; "
    End If
    
    ' Форматирование
    If cell.Font.Bold Then cellNotes = cellNotes & "Жирный; "
    If cell.Font.Italic Then cellNotes = cellNotes & "Курсив; "
    If cell.Font.Underline Then cellNotes = cellNotes & "Подчеркивание; "
    If cell.Font.color <> -16777216 Then cellNotes = cellNotes & "Цвет текста; "
    If cell.Interior.color <> 16777215 Then cellNotes = cellNotes & "Заливка; "
    If cell.Borders.count > 0 Then
        Dim hasBorder As Boolean
        hasBorder = False
        Dim border As border
        For Each border In cell.Borders
            If border.LineStyle <> xlLineStyleNone Then
                hasBorder = True
                Exit For
            End If
        Next
        If hasBorder Then cellNotes = cellNotes & "Границы; "
    End If
    
    ' Собираем все вместе
    details = cellAddr & " | " & cellType & " | " & _
              Left(cellValue, 100) & " | " & _
              Left(cellFormat, 20) & " | " & _
              Left(cellNotes, 50)
    
    ЗаписатьПодробностиЯчейки = details & vbCrLf
End Function

' ----------------------------------------------------------------
' ПОДПРОЦЕДУРА: Формулы и расчеты
' ----------------------------------------------------------------
Private Sub ЗаписатьФормулыИРасчеты(txtFile As Object, ws As Worksheet, _
                                     usedRange As Range, formulaCount As Long)
    Dim cell As Range
    Dim formulaCells As Range
    Dim uniqueFunctions As Object
    Dim func As Variant
    
    With txtFile
        .WriteLine "3. ФОРМУЛЫ И РАСЧЕТЫ"
        .WriteLine String(80, "-")
        
        If formulaCount = 0 Then
            .WriteLine "На листе нет формул."
            .WriteLine ""
            Exit Sub
        End If
        
        .WriteLine "Всего формул на листе: " & formulaCount
        .WriteLine ""
        
        ' Сбор уникальных функций
        Set uniqueFunctions = CreateObject("Scripting.Dictionary")
        
        On Error Resume Next
        Set formulaCells = usedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        
        If Not formulaCells Is Nothing Then
            .WriteLine "СПИСОК ФОРМУЛ:"
            .WriteLine String(60, "-")
            
            For Each cell In formulaCells
                .WriteLine "Ячейка " & cell.Address & ":"
                .WriteLine "  Формула: " & cell.Formula
                
                If Not IsEmpty(cell.Value) Then
                    .WriteLine "  Результат: " & CStr(cell.Value)
                End If
                
                ' Анализ зависимостей
                Dim precedents As Range
                On Error Resume Next
                Set precedents = cell.DirectPrecedents
                On Error GoTo 0
                
                If Not precedents Is Nothing Then
                    .WriteLine "  Зависит от: " & precedents.Address
                End If
                
                ' Определение функций в формуле
                Dim formulaText As String
                formulaText = UCase(cell.Formula)
                
                Dim funcList As Variant
                funcList = Array("SUM", "AVERAGE", "COUNT", "IF", "VLOOKUP", "HLOOKUP", _
                                "INDEX", "MATCH", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", _
                                "MAX", "MIN", "ROUND", "DATE", "TEXT", "LEFT", "RIGHT", _
                                "MID", "FIND", "SEARCH", "LEN", "TRIM", "CONCATENATE", _
                                "&", "+", "-", "*", "/", "^", "(", ")", "$")
                
                For Each func In funcList
                    If InStr(formulaText, func) > 0 Then
                        If Not uniqueFunctions.Exists(func) Then
                            uniqueFunctions.Add func, 1
                        Else
                            uniqueFunctions(func) = uniqueFunctions(func) + 1
                        End If
                    End If
                Next func
                
                .WriteLine ""
            Next cell
        End If
        
        ' Статистика по функциям
        If uniqueFunctions.count > 0 Then
            .WriteLine "СТАТИСТИКА ИСПОЛЬЗУЕМЫХ ФУНКЦИЙ:"
            .WriteLine String(60, "-")
            
            For Each func In uniqueFunctions.Keys
                .WriteLine "  " & func & ": " & uniqueFunctions(func) & " раз"
            Next func
        End If
        
        .WriteLine ""
    End With
    
    Set uniqueFunctions = Nothing
End Sub

' ----------------------------------------------------------------
' ПОДПРОЦЕДУРА: Именованные диапазоны
' ----------------------------------------------------------------
Private Sub ЗаписатьИменованныеДиапазоны(txtFile As Object, ws As Worksheet)
    Dim namedRange As Name
    Dim count As Long
    
    With txtFile
        .WriteLine "4. ИМЕНОВАННЫЕ ДИАПАЗОНЫ"
        .WriteLine String(80, "-")
        
        count = 0
        For Each namedRange In ws.Names
            If Not namedRange.RefersTo Is Nothing Then
                .WriteLine "Имя: " & namedRange.Name
                .WriteLine "  Диапазон: " & namedRange.RefersTo
                .WriteLine "  Видимость: " & IIf(namedRange.Visible, "Видимый", "Скрытый")
                .WriteLine "  Комментарий: " & IIf(namedRange.comment = "", "(нет)", namedRange.comment)
                .WriteLine ""
                count = count + 1
            End If
        Next namedRange
        
        If count = 0 Then
            .WriteLine "На листе нет именованных диапазонов."
        Else
            .WriteLine "Всего именованных диапазонов: " & count
        End If
        
        .WriteLine ""
    End With
End Sub

' ----------------------------------------------------------------
' ПОДПРОЦЕДУРА: Графические объекты
' ----------------------------------------------------------------
Private Sub ЗаписатьГрафическиеОбъекты(txtFile As Object, ws As Worksheet)
    Dim shape As shape
    Dim chart As ChartObject
    Dim shapeCount As Long, chartCount As Long
    
    With txtFile
        .WriteLine "5. ГРАФИЧЕСКИЕ ОБЪЕКТЫ"
        .WriteLine String(80, "-")
        
        shapeCount = 0
        chartCount = 0
        
        If ws.Shapes.count > 0 Then
            For Each shape In ws.Shapes
                shapeCount = shapeCount + 1
                
                .WriteLine "Объект #" & shapeCount & ": " & shape.Name
                .WriteLine "  Тип: " & ПолучитьТипОбъекта(shape.Type)
                .WriteLine "  Размер: " & Round(shape.Width, 2) & " x " & _
                           Round(shape.Height, 2) & " точек"
                .WriteLine "  Положение: (" & Round(shape.Top, 2) & ", " & _
                           Round(shape.Left, 2) & ")"
                
                If shape.Type = msoChart Then
                    chartCount = chartCount + 1
                    On Error Resume Next
                    .WriteLine "  Тип диаграммы: " & shape.chart.ChartType
                    .WriteLine "  Название: " & shape.chart.ChartTitle.text
                    On Error GoTo 0
                End If
                
                .WriteLine ""
            Next shape
            
            .WriteLine "ИТОГО объектов: " & shapeCount
            .WriteLine "  • Диаграмм: " & chartCount
            .WriteLine "  • Прочих объектов: " & (shapeCount - chartCount)
        Else
            .WriteLine "На листе нет графических объектов."
        End If
        
        .WriteLine ""
    End With
End Sub

' ----------------------------------------------------------------
' ПОДПРОЦЕДУРА: Сводные таблицы
' ----------------------------------------------------------------
Private Sub ЗаписатьСводныеТаблицы(txtFile As Object, ws As Worksheet)
    Dim pivot As PivotTable
    Dim pivotCount As Long
    
    With txtFile
        .WriteLine "6. СВОДНЫЕ ТАБЛИЦЫ"
        .WriteLine String(80, "-")
        
        pivotCount = 0
        
        If ws.PivotTables.count > 0 Then
            For Each pivot In ws.PivotTables
                pivotCount = pivotCount + 1
                
                .WriteLine "Сводная таблица #" & pivotCount & ": " & pivot.Name
                .WriteLine "  Диапазон: " & pivot.TableRange1.Address
                .WriteLine "  Источник данных: " & pivot.SourceData
                .WriteLine "  Поля строк: " & pivot.RowFields.count
                .WriteLine "  Поля столбцов: " & pivot.ColumnFields.count
                .WriteLine "  Поля значений: " & pivot.DataFields.count
                .WriteLine "  Фильтры: " & pivot.PageFields.count
                .WriteLine ""
            Next pivot
            
            .WriteLine "Всего сводных таблиц: " & pivotCount
        Else
            .WriteLine "На листе нет сводных таблиц."
        End If
        
        .WriteLine ""
    End With
End Sub

' ----------------------------------------------------------------
' ПОДПРОЦЕДУРА: Настройки листа
' ----------------------------------------------------------------
Private Sub ЗаписатьНастройкиЛиста(txtFile As Object, ws As Worksheet)
    With txtFile
        .WriteLine "7. НАСТРОЙКИ ЛИСТА"
        .WriteLine String(80, "-")
        
        ' Настройки отображения
        .WriteLine "НАСТРОЙКИ ОТОБРАЖЕНИЯ: "
        .WriteLine " • Сетка: " & IIf(ws.DisplayGridlines, "Включена", "Выключена") ' ИСПРАВЛЕНО!
        .WriteLine " • Заголовки строк/столбцов: " & IIf(ws.DisplayHeadings, "Включены", "Выключены")
        .WriteLine " • Нулевые значения: " & IIf(ws.DisplayZeros, "Отображаются", "Скрыты")
        .WriteLine " • Формулы вместо значений: " & IIf(ws.DisplayFormulas, "Показываются", "Скрыты")
        
        ' Защита
        .WriteLine "ЗАЩИТА ЛИСТА: "
        .WriteLine " • Защищён: " & IIf(ws.ProtectContents, "Да", "Нет")
        
        ' Безопасный способ получения количества разрешенных диапазонов
        Dim rangeCount As Long
        rangeCount = 0
        
        If ws.ProtectContents Then
            On Error Resume Next
            rangeCount = ws.Protection.AllowEditRanges.count
            On Error GoTo 0
        End If
        
        If ws.ProtectContents And rangeCount > 0 Then
            .WriteLine " • Разрешённые диапазоны для редактирования: " & rangeCount & " диапазон(ов)"
        ElseIf ws.ProtectContents Then
            .WriteLine " • Разрешённые диапазоны для редактирования: нет (лист полностью защищён)"
        Else
            .WriteLine " • Разрешённые диапазоны для редактирования: не применимо (лист не защищён)"
        End If
    End With
End Sub

' ----------------------------------------------------------------
' ПОДПРОЦЕДУРА: Статистика
' ----------------------------------------------------------------
Private Sub ЗаписатьСтатистику(txtFile As Object, ws As Worksheet, _
                               lastRow As Long, lastCol As Long, _
                               formulaCount As Long, valueCount As Long, _
                               emptyCount As Long, startTime As Double)
    Dim totalCells As Long
    Dim processingTime As Double
    
    totalCells = lastRow * lastCol
    processingTime = Timer - startTime
    
    With txtFile
        .WriteLine "8. СТАТИСТИКА И СВОДКА"
        .WriteLine String(80, "-")
        
        .WriteLine "ОБЩАЯ СТАТИСТИКА:"
        .WriteLine "  • Всего ячеек в используемом диапазоне: " & totalCells
        .WriteLine "  • Ячеек с формулами: " & formulaCount & " (" & _
                   Format(formulaCount / totalCells * 100, "0.0") & "%)"
        .WriteLine "  • Ячеек со значениями: " & valueCount & " (" & _
                   Format(valueCount / totalCells * 100, "0.0") & "%)"
        .WriteLine "  • Пустых ячеек: " & emptyCount & " (" & _
                   Format(emptyCount / totalCells * 100, "0.0") & "%)"
        .WriteLine ""
        
        .WriteLine "ДОПОЛНИТЕЛЬНАЯ ИНФОРМАЦИЯ:"
        .WriteLine "  • Именованные диапазоны: " & ws.Names.count
        .WriteLine "  • Графические объекты: " & ws.Shapes.count
        .WriteLine "  • Сводные таблицы: " & ws.PivotTables.count
        .WriteLine "  • Окно разбито на области: " & IIf(ws.Panes.count > 1, "Да", "Нет")
        .WriteLine ""
        
        .WriteLine "ПРОИЗВОДИТЕЛЬНОСТЬ:"
        .WriteLine "  • Время обработки: " & Format(processingTime, "0.00") & " секунд"
        .WriteLine "  • Скорость обработки: " & _
                   Format(totalCells / processingTime, "0") & " ячеек/сек"
        .WriteLine ""
        
        .WriteLine String(100, "=")
        .WriteLine "ОПИСАНИЕ ЗАВЕРШЕНО"
        .WriteLine "Файл сохранен: " & ThisWorkbook.Path & "\Описание_листа_" & _
                   ws.Name & "_" & Format(Now, "yyyy-mm-dd_HH-mm") & ".txt"
        .WriteLine String(100, "=")
    End With
End Sub

' ----------------------------------------------------------------
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' ----------------------------------------------------------------

Private Function ПолучитьВидимостьЛиста(ws As Worksheet) As String
    Select Case ws.Visible
        Case xlSheetVisible: ПолучитьВидимостьЛиста = "Видимый"
        Case xlSheetHidden: ПолучитьВидимостьЛиста = "Скрытый"
        Case xlSheetVeryHidden: ПолучитьВидимостьЛиста = "Очень скрытый"
        Case Else: ПолучитьВидимостьЛиста = "Неизвестно"
    End Select
End Function

Private Function ПолучитьЦветЯрлычка(ws As Worksheet) As String
    If ws.Tab.ColorIndex = xlColorIndexNone Then
        ПолучитьЦветЯрлычка = "Нет"
    Else
        Dim color As Long
        color = ws.Tab.color
        
        ' Преобразование RGB
        Dim red As Integer, green As Integer, blue As Integer
        red = color Mod 256
        green = (color \ 256) Mod 256
        blue = (color \ 65536) Mod 256
        
        ПолучитьЦветЯрлычка = "RGB(" & red & ", " & green & ", " & blue & ")"
    End If
End Function

Private Function ПолучитьИмяСтолбца(colNum As Long) As String
    Dim dividend As Long
    Dim columnName As String
    Dim modulo As Long
    
    dividend = colNum
    columnName = ""
    
    While dividend > 0
        modulo = (dividend - 1) Mod 26
        columnName = Chr(65 + modulo) & columnName
        dividend = (dividend - modulo) \ 26
    Wend
    
    ПолучитьИмяСтолбца = columnName
End Function

Private Function ПолучитьТипОбъекта(shapeType As MsoShapeType) As String
    Select Case shapeType
        Case msoAutoShape: ПолучитьТипОбъекта = "Автофигура"
        Case msoCallout: ПолучитьТипОбъекта = "Выноска"
        Case msoChart: ПолучитьТипОбъекта = "Диаграмма"
        Case msoComment: ПолучитьТипОбъекта = "Комментарий"
        Case msoFreeform: ПолучитьТипОбъекта = "Произвольная форма"
        Case msoGroup: ПолучитьТипОбъекта = "Группа"
        Case msoEmbeddedOLEObject: ПолучитьТипОбъекта = "OLE-объект"
        Case msoFormControl: ПолучитьТипОбъекта = "Элемент управления"
        Case msoLine: ПолучитьТипОбъекта = "Линия"
        Case msoLinkedOLEObject: ПолучитьТипОбъекта = "Связанный OLE-объект"
        Case msoLinkedPicture: ПолучитьТипОбъекта = "Связанное изображение"
        Case msoOLEControlObject: ПолучитьТипОбъекта = "Элемент управления OLE"
        Case msoPicture: ПолучитьТипОбъекта = "Изображение"
        Case msoPlaceholder: ПолучитьТипОбъекта = "Заполнитель"
        Case msoMedia: ПолучитьТипОбъекта = "Медиа"
        Case msoTextBox: ПолучитьТипОбъекта = "Текстовое поле"
        Case msoTextEffect: ПолучитьТипОбъекта = "Текстовый эффект"
        Case msoScriptAnchor: ПолучитьТипОбъекта = "Якорь скрипта"
        Case Else: ПолучитьТипОбъекта = "Неизвестный тип (" & shapeType & ")"
    End Select
End Function

' ----------------------------------------------------------------
' ПРОЦЕДУРА ДЛЯ ЗАПУСКА ИЗ МЕНЮ (дополнительно)
' ----------------------------------------------------------------
Sub ДобавитьВМеню()
    Dim cmdBar As CommandBar
    Dim cmdButton As CommandBarButton
    
    ' Удалить старую кнопку, если существует
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls("Анализ листа").Delete
    On Error GoTo 0
    
    ' Добавить новую кнопку в меню
    Set cmdBar = Application.CommandBars("Worksheet Menu Bar")
    Set cmdButton = cmdBar.Controls.Add(Type:=msoControlButton, Before:=11)
    
    With cmdButton
        .Caption = "Анализ листа"
        .FaceId = 225 ' Иконка документа
        .OnAction = "СоздатьОписаниеЛиста"
        .TooltipText = "Создать подробное текстовое описание текущего листа"
    End With
End Sub

' ----------------------------------------------------------------
' ПРОЦЕДУРА ДЛЯ УДАЛЕНИЯ ИЗ МЕНЮ
' ----------------------------------------------------------------
Sub УдалитьИзМеню()
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls("Анализ листа").Delete
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
' ИНИЦИАЛИЗАЦИЯ (добавить в модуль ThisWorkbook)
' ----------------------------------------------------------------
' Private Sub Workbook_Open()
'     ' Автоматически добавить кнопку в меню при открытии книги
'     Call ДобавитьВМеню
' End Sub
'
' Private Sub Workbook_BeforeClose(Cancel As Boolean)
'     ' Удалить кнопку из меню при закрытии книги
'     Call УдалитьИзМеню
' End Sub

