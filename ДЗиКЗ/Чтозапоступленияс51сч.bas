Attribute VB_Name = "Чтозапоступленияс51сч"
Sub ProcessAllSheets()
    Dim startTime As Double
    startTime = Timer ' Запоминаем время начала (в секундах с начала дня)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    
    On Error GoTo ErrorHandler

    ' Открываем диалог выбора файла
    Dim targetFilePath As String
    targetFilePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx; *.xls; *.xlsm), *.xlsx; *.xls; *.xlsm", _
        Title:="Выберите файл с данными по описаниям", _
        MultiSelect:=False)
    
    If targetFilePath = "False" Then
        MsgBox "Выбор файла отменён. Обработка прервана.", vbExclamation
        GoTo CleanUp
    End If

    ' Подсчитываем общее количество контрагентов для прогресс-бара
    Dim totalAgents As Long
    totalAgents = CountAgentsInSheets()

    If totalAgents = 0 Then
        MsgBox "Нет контрагентов для обработки.", vbInformation
        GoTo CleanUp
    End If

    Dim currentAgent As Long
    currentAgent = 0

    ' Обрабатываем все три листа
    ProcessSingleSheet ThisWorkbook.Sheets("Основные КА"), "A3:A7,A13:A17", 3, targetFilePath, currentAgent, totalAgents
    ProcessSingleSheet ThisWorkbook.Sheets("Дебиторы"), "A4:A8, A12:A16, A20:A24, A28:A32", 7, targetFilePath, currentAgent, totalAgents
    ProcessSingleSheet ThisWorkbook.Sheets("Кредиторы"), "A4:A8, A12:A16, A20:A24, A28:A32", 7, targetFilePath, currentAgent, totalAgents

    ' Вычисляем время выполнения
    Dim endTime As Double
    Dim elapsedSeconds As Long
    Dim minutes As Long, seconds As Long
    
    endTime = Timer
    elapsedSeconds = CLng(endTime - startTime)
    
    ' Обработка перехода через полночь (редко, но возможно)
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400
    
    minutes = elapsedSeconds \ 60
    seconds = elapsedSeconds Mod 60
    
    Dim timeMessage As String
    If minutes > 0 Then
        timeMessage = minutes & " мин " & seconds & " сек"
    Else
        timeMessage = seconds & " сек"
    End If

    ' Уведомление об успешном завершении + время
    MsgBox "Загрузка информации завершена." & vbCrLf & _
           "Время выполнения: " & timeMessage, vbInformation

CleanUp:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Произошла ошибка: " & Err.Description, vbCritical
    Resume CleanUp
End Sub


' Подсчёт общего количества непустых ячеек в указанных диапазонах
Function CountAgentsInSheets() As Long
    Dim count As Long
    count = 0

    count = count + CountNonEmptyCells(ThisWorkbook.Sheets("Основные КА").Range("A3:A7,A13:A17"))
    count = count + CountNonEmptyCells(ThisWorkbook.Sheets("Дебиторы").Range("A4:A8, A12:A16, A20:A24, A28:A32"))
    count = count + CountNonEmptyCells(ThisWorkbook.Sheets("Кредиторы").Range("A4:A8, A12:A16, A20:A24, A28:A32"))

    CountAgentsInSheets = count
End Function


' Считает непустые и неошибочные ячейки в объединённом диапазоне
Function CountNonEmptyCells(rng As Range) As Long
    Dim cell As Range
    Dim cnt As Long
    cnt = 0
    For Each cell In rng
        If cell.Value <> "" And Not IsError(cell.Value) Then
            cnt = cnt + 1
        End If
    Next cell
    CountNonEmptyCells = cnt
End Function


' Обновлённая версия с прогресс-баром
Sub ProcessSingleSheet(ws As Worksheet, agentAddress As String, resultOffset As Integer, targetFilePath As String, ByRef currentAgent As Long, totalAgents As Long)
    Dim agentRange As Range
    Dim cell As Range
    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    Dim isOpened As Boolean
    
    Set agentRange = ws.Range(agentAddress)
    
    ' Открытие целевого файла
isOpened = False
On Error Resume Next
Set wbTarget = GetWorkbookByPath(targetFilePath)
If wbTarget Is Nothing Then
    Set wbTarget = Workbooks.Open(targetFilePath, ReadOnly:=True, AddToMru:=False)
    isOpened = True
End If
On Error GoTo 0

If wbTarget Is Nothing Then
    MsgBox "Не удалось открыть файл: " & targetFilePath, vbCritical
    Exit Sub
End If

' Поиск нужного листа: сначала "Лист_1", затем "Коп сюда"
On Error Resume Next
Set wsTarget = wbTarget.Sheets("Лист_1")
On Error GoTo 0

If wsTarget Is Nothing Then
    On Error Resume Next
    Set wsTarget = wbTarget.Sheets("Коп сюда")
    On Error GoTo 0
End If

If wsTarget Is Nothing Then
    MsgBox "Лист 'Лист_1' или 'Коп сюда' не найден в файле: " & targetFilePath, vbExclamation
    If isOpened Then wbTarget.Close False
    Exit Sub
End If
    
    ' Основной цикл по контрагентам
    For Each cell In agentRange
        If cell.Value <> "" And Not IsError(cell.Value) Then
            currentAgent = currentAgent + 1
            UpdateProgressBar currentAgent, totalAgents ' Обновляем прогресс

            Dim targetAgent As String
            targetAgent = Trim(CStr(cell.Value))
            
            Dim resultCell As Range
            Set resultCell = cell.Offset(0, resultOffset)
            
            Dim result As String
            result = AnalyzeAgent(targetAgent, wsTarget)
            resultCell.Value = result
        End If
    Next cell
    
    If isOpened Then
        wbTarget.Close False
    End If
End Sub


' Обновляет статусную строку с процентом выполнения
Sub UpdateProgressBar(currentStep As Long, totalSteps As Long)
    If totalSteps <= 0 Then Exit Sub
    Dim percent As Long
    percent = Int((currentStep / totalSteps) * 100)
    Application.StatusBar = "Обработка контрагентов: " & currentStep & " из " & totalSteps & " (" & percent & "%)"
    DoEvents ' Позволяет Excel обновить интерфейс
End Sub


' Остальные функции без изменений (можно копировать как есть)
Function AnalyzeAgent(targetAgent As String, wsTarget As Worksheet) As String
    Dim lastRow As Long
    Dim descriptions As New Collection
    Dim i As Long
    
    lastRow = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).row
    If lastRow < 2 Then
        AnalyzeAgent = "Нет данных"
        Exit Function
    End If
    
    Dim dataRange As Range
    Set dataRange = wsTarget.Range("B2:D" & lastRow)
    Dim dataArray As Variant
    dataArray = dataRange.Value
    
    For i = 1 To UBound(dataArray, 1)
        If InStr(1, CStr(dataArray(i, 2)), targetAgent, vbTextCompare) > 0 Or _
           InStr(1, CStr(dataArray(i, 3)), targetAgent, vbTextCompare) > 0 Then
            descriptions.Add CStr(dataArray(i, 1))
        End If
    Next i
    
    If descriptions.count = 0 Then
        AnalyzeAgent = "Нет данных"
        Exit Function
    End If
    
    Dim phraseDict As Object
    Set phraseDict = CreateObject("Scripting.Dictionary")
    
    Dim text As Variant
    For Each text In descriptions
        ProcessText CStr(text), phraseDict
    Next text
    
    AnalyzeAgent = GetTopThreePhrasesWithFreq(phraseDict)
End Function


Function GetWorkbookByPath(fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.FullName = fullPath Then
            Set GetWorkbookByPath = wb
            Exit Function
        End If
    Next wb
    Set GetWorkbookByPath = Nothing
End Function


Sub ProcessText(text As String, dict As Object)
    Static stopWords As Object
    If stopWords Is Nothing Then
        Set stopWords = CreateObject("Scripting.Dictionary")
        Dim stopWordsArray As Variant
        stopWordsArray = Array("поступление", "расчетный", "расчетного", "счета", "счет", "списание", _
                              "оплата", "частичная", "окончательный", "расчет", "вх", "д", "тч", _
                              "ндс", "от", "за", "на", "по", "в", "с", "г", "тыс", "руб", "вхд", _
                              "дог", "дол", "акт", "улд", "упд", "бп", "ч", "00", "06", "20", "25", _
                              "26", "30", "31", "40", "81", "82", "19", "194", "190", "186", "183", "181")
        
        Dim word As Variant
        For Each word In stopWordsArray
            stopWords.Add word, True
        Next word
    End If
    
    Dim cleanedText As String
    Dim i As Long
    For i = 1 To Len(text)
        Dim char As String
        char = Mid(text, i, 1)
        If (char >= "А" And char <= "я") Or (char >= "A" And char <= "z") Then
            cleanedText = cleanedText & LCase(char)
        Else
            cleanedText = cleanedText & " "
        End If
    Next i
    
    Dim words() As String
    words = Split(Application.WorksheetFunction.Trim(cleanedText), " ")
    
    Dim filteredWords() As String
    ReDim filteredWords(UBound(words))
    Dim wordCount As Long
    wordCount = 0
    
    Dim wordVal As Variant
    For Each wordVal In words
        wordVal = Trim(wordVal)
        If Len(wordVal) >= 3 And Not stopWords.Exists(wordVal) And Not IsNumeric(wordVal) Then
            filteredWords(wordCount) = wordVal
            wordCount = wordCount + 1
        End If
    Next wordVal
    
    If wordCount >= 3 Then
        ReDim Preserve filteredWords(wordCount - 1)
        
        Dim start As Long
        For start = 0 To wordCount - 3
            Dim phrase As String
            phrase = filteredWords(start) & " " & filteredWords(start + 1) & " " & filteredWords(start + 2)
            
            If Not ContainsNumbers(phrase) Then
                If dict.Exists(phrase) Then
                    dict(phrase) = dict(phrase) + 1
                Else
                    dict.Add phrase, 1
                End If
            End If
        Next start
    End If
End Sub


Function ContainsNumbers(phrase As String) As Boolean
    Dim i As Long
    For i = 1 To Len(phrase)
        If Mid(phrase, i, 1) Like "[0-9]" Then
            ContainsNumbers = True
            Exit Function
        End If
    Next i
    ContainsNumbers = False
End Function


Function GetTopThreePhrasesWithFreq(dict As Object) As String
    If dict.count = 0 Then
        GetTopThreePhrasesWithFreq = "Нет фраз"
        Exit Function
    End If
    
    Dim phrases() As String
    Dim frequencies() As Long
    ReDim phrases(dict.count - 1)
    ReDim frequencies(dict.count - 1)
    
    Dim i As Long, idx As Long
    Dim key As Variant
    idx = 0
    For Each key In dict.Keys
        phrases(idx) = key
        frequencies(idx) = dict(key)
        idx = idx + 1
    Next key
    
    For i = 0 To UBound(phrases) - 1
        For idx = i + 1 To UBound(phrases)
            If frequencies(idx) > frequencies(i) Then
                Swap phrases(i), phrases(idx)
                Swap frequencies(i), frequencies(idx)
            End If
        Next idx
    Next i
    
    Dim result As String
    Dim count As Long
    count = Application.WorksheetFunction.Min(3, dict.count)
    
    For i = 0 To count - 1
        If i > 0 Then result = result & ", "
        result = result & phrases(i) & " (" & frequencies(i) & ")"
    Next i
    
    GetTopThreePhrasesWithFreq = result
End Function


Sub Swap(ByRef a As Variant, ByRef b As Variant)
    Dim temp As Variant
    temp = a
    a = b
    b = temp
End Sub






