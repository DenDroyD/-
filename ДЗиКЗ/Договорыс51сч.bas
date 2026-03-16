Attribute VB_Name = "Договорыс51сч"
Sub ProcessContracts()
    Dim startTime As Double
    startTime = Timer ' <<< Запоминаем время начала

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = True
    
    On Error GoTo ErrorHandler

    ' Открываем диалог выбора файла
    Dim targetFilePath As String
    targetFilePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx; *.xls; *.xlsm), *.xlsx; *.xls; *.xlsm", _
        Title:="Выберите файл с данными по договорам", _
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
    ProcessContractsSingleSheet ThisWorkbook.Sheets("Дебиторы"), "A4:A8, A12:A16, A20:A24, A28:A32", 5, targetFilePath, currentAgent, totalAgents
    ProcessContractsSingleSheet ThisWorkbook.Sheets("Кредиторы"), "A4:B8, A12:A16, A20:A24, A28:A32", 5, targetFilePath, currentAgent, totalAgents

    ' === ВЫЧИСЛЕНИЕ ВРЕМЕНИ ВЫПОЛНЕНИЯ ===
    Dim endTime As Double
    Dim elapsedSeconds As Long
    Dim minutes As Long, seconds As Long
    
    endTime = Timer
    elapsedSeconds = CLng(endTime - startTime)
    
    ' Обработка перехода через полночь (на всякий случай)
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400
    
    minutes = elapsedSeconds \ 60
    seconds = elapsedSeconds Mod 60
    
    Dim timeMessage As String
    If minutes > 0 Then
        timeMessage = minutes & " мин " & seconds & " сек"
    Else
        timeMessage = seconds & " сек"
    End If

    ' Уведомление с временем выполнения
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
Sub ProcessContractsSingleSheet(ws As Worksheet, agentAddress As String, resultOffset As Integer, targetFilePath As String, ByRef currentAgent As Long, totalAgents As Long)
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
            result = AnalyzeAgentForContracts(targetAgent, wsTarget)
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
Function AnalyzeAgentForContracts(targetAgent As String, wsTarget As Worksheet) As String
    Dim lastRow As Long
    Dim descriptions As New Collection
    Dim i As Long
    
    lastRow = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).row
    If lastRow < 2 Then
        AnalyzeAgentForContracts = "Нет данных"
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
        AnalyzeAgentForContracts = "Нет данных"
        Exit Function
    End If
    
    Dim contractDict As Object
    Set contractDict = CreateObject("Scripting.Dictionary")
    
    Dim text As Variant
    For Each text In descriptions
        ExtractContractInfo CStr(text), contractDict
    Next text
    
    AnalyzeAgentForContracts = GetTopThreeContractsWithFreq(contractDict)
End Function


Sub ExtractContractInfo(text As String, dict As Object)
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.Pattern = "(договор|контракт)[а-я]*\s*№?\s*[0-9\/\-а-яА-Яa-zA-Z]+\s*от?\s*[0-9]{1,2}\.[0-9]{1,2}\.[0-9]{2,4}"
    
    Dim matches As Object
    Set matches = regEx.Execute(text)
    
    Dim match As Object
    For Each match In matches
        Dim contractPhrase As String
        contractPhrase = match.Value
        contractPhrase = NormalizeContractPhrase(contractPhrase)
        
        If dict.Exists(contractPhrase) Then
            dict(contractPhrase) = dict(contractPhrase) + 1
        Else
            dict.Add contractPhrase, 1
        End If
    Next match
End Sub


Function NormalizeContractPhrase(phrase As String) As String
    Dim result As String
    result = phrase
    
    If InStr(1, result, "контракт", vbTextCompare) > 0 Then
        result = Replace(result, "контракт", "договор", 1, -1, vbTextCompare)
    End If
    
    result = Application.WorksheetFunction.Trim(result)
    result = Replace(result, "дог ", "договор ", 1, -1, vbTextCompare)
    result = Replace(result, " от ", " от ", 1, -1, vbTextCompare)
    
    NormalizeContractPhrase = result
End Function


Function GetTopThreeContractsWithFreq(dict As Object) As String
    If dict.count = 0 Then
        GetTopThreeContractsWithFreq = "Нет договоров"
        Exit Function
    End If
    
    Dim contracts() As String
    Dim frequencies() As Long
    ReDim contracts(1 To dict.count)
    ReDim frequencies(1 To dict.count)
    
    Dim i As Long, j As Long
    Dim key As Variant
    i = 1
    For Each key In dict.Keys
        contracts(i) = key
        frequencies(i) = dict(key)
        i = i + 1
    Next key
    
    For i = 1 To dict.count - 1
        For j = i + 1 To dict.count
            If frequencies(j) > frequencies(i) Then
                Dim tempFreq As Long
                tempFreq = frequencies(i)
                frequencies(i) = frequencies(j)
                frequencies(j) = tempFreq
                
                Dim tempContract As String
                tempContract = contracts(i)
                contracts(i) = contracts(j)
                contracts(j) = tempContract
            End If
        Next j
    Next i
    
    Dim result As String
    Dim count As Long
    count = Application.WorksheetFunction.Min(3, dict.count)
    
    For i = 1 To count
        If i > 1 Then result = result & ", "
        result = result & contracts(i) & " (" & frequencies(i) & ")"
    Next i
    
    GetTopThreeContractsWithFreq = result
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

