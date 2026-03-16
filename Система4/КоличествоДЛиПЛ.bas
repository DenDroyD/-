Attribute VB_Name = "КоличествоДЛиПЛ"
Sub AnalyzeLeasingDataCopy()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim lastRowSource As Long
    Dim i As Long, lastRowTarget As Long
    Dim cellValue As Variant
    Dim periodStart As Date, periodEnd As Date, prevPeriodEnd As Date
    Dim countTS As Long, countDL As Long, countActiveTS As Long
    Dim firstRow As Boolean
    Dim targetCell As Range
    Dim activeContracts As Long, archivedContracts As Long
    
    Set targetWorkbook = ThisWorkbook
    Set targetSheet = targetWorkbook.ActiveSheet
    
    ' Запрашиваем файл
    filePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx; *.xlsm; *.xls), *.xlsx; *.xlsm; *.xls", _
        Title:="Выберите файл СПАРК с данными по лизингу")
    
    If filePath = "False" Then Exit Sub
    
    ' Открываем файл
    Set sourceWorkbook = Workbooks.Open(filePath, ReadOnly:=True)
    
    ' Находим лист с данными
    On Error Resume Next
    Set sourceSheet = sourceWorkbook.Sheets("report")
    If sourceSheet Is Nothing Then Set sourceSheet = sourceWorkbook.Sheets(1)
    On Error GoTo ErrorHandler
    
    If sourceSheet Is Nothing Then
        MsgBox "Не удалось найти лист с данными!", vbExclamation
        Exit Sub
    End If
    
    ' ---- Подсчёт действующих и архивных договоров на текущую дату ----
    activeContracts = 0
    archivedContracts = 0
    
    With sourceSheet
        lastRowSource = .Cells(.Rows.Count, "F").End(xlUp).row
        Dim j As Long
        For j = 2 To lastRowSource
            If IsDate(.Cells(j, 5).Value) And IsDate(.Cells(j, 6).Value) Then
                Dim startDate As Date
                Dim endDate As Date
                startDate = .Cells(j, 5).Value
                endDate = .Cells(j, 6).Value
                
                If endDate >= Date And startDate <= Date Then
                    activeContracts = activeContracts + 1
                ElseIf endDate < Date Then
                    archivedContracts = archivedContracts + 1
                End If
            End If
        Next j
    End With
    
    ' ---- Анализ по периодам ----
    firstRow = True
    prevPeriodEnd = 0
    
    ' ?? ИЗМЕНЕНО: теперь обрабатываются только строки 113–117
    For i = 113 To 117
        cellValue = targetSheet.Cells(i, 1).Value
        
        ' Пропускаем пустые ячейки
        If IsEmpty(cellValue) Then GoTo NextRow
        
        ' Определяем дату окончания интервала (periodEnd)
        If IsNumeric(cellValue) And cellValue >= 1900 And cellValue <= 2100 Then
            ' Это год
            periodEnd = DateSerial(cellValue, 12, 31)
        ElseIf IsDate(cellValue) Then
            ' Это дата
            Dim d As Date
            d = CDate(cellValue)
            If Day(d) = 1 Then
                ' Считаем это месяцем (первый день месяца) – интервал до последнего дня месяца
                periodEnd = DateSerial(Year(d), Month(d) + 1, 0)
            Else
                ' Конкретный день
                periodEnd = d
            End If
        Else
            ' Не удалось распознать – пропускаем
            GoTo NextRow
        End If
        
        ' Определяем начало интервала
        If firstRow Then
            periodStart = Date ' сегодня
            firstRow = False
        Else
            periodStart = prevPeriodEnd + 1 ' следующий день после предыдущего окончания
        End If
        
        ' Если начало интервала позже конца (например, из-за неправильного порядка дат) – пропускаем
        If periodStart > periodEnd Then GoTo NextRow
        
        ' Подсчёт для данного интервала
        countTS = 0
        countDL = 0
        countActiveTS = 0
        
        With sourceSheet
            For j = 2 To lastRowSource
                If IsDate(.Cells(j, 5).Value) And IsDate(.Cells(j, 6).Value) Then
                    startDate = .Cells(j, 5).Value
                    endDate = .Cells(j, 6).Value
                    
                    ' 1. Договоры, завершающиеся в интервале (и ещё не завершённые)
                    If endDate >= periodStart And endDate <= periodEnd And endDate > Date Then
                        countDL = countDL + 1
                        If IsNumeric(.Cells(j, 7).Value) Then
                            countTS = countTS + .Cells(j, 7).Value
                        End If
                    End If
                    
                    ' 2. Активные ТС в интервале (договор действует хотя бы один день интервала и ещё не завершён)
                    If startDate <= periodEnd And endDate >= periodStart And endDate >= Date Then
                        If IsNumeric(.Cells(j, 7).Value) Then
                            countActiveTS = countActiveTS + .Cells(j, 7).Value
                        End If
                    End If
                End If
            Next j
        End With
        
        ' ? Запись результатов (старые данные автоматически перезаписываются)
        targetSheet.Cells(i, 2).Value = countTS       ' B – ТС завершающиеся
        targetSheet.Cells(i, 4).Value = countDL       ' D – ДЛ завершающиеся
        targetSheet.Cells(i, 6).Value = countActiveTS ' F – активные ТС в интервале
        
        ' Запоминаем конец текущего интервала для следующей итерации
        prevPeriodEnd = periodEnd
        
NextRow:
    Next i
    
    ' ---- Запись итоговой фразы в A118 ----
    targetSheet.Cells(118, 1).Value = "На " & Format(Date, "dd.mm.yyyy") & " клиент имеется " & activeContracts & " действующих и " & archivedContracts & " архивных договоров лизинга"
    
    ' Закрываем исходный файл
    sourceWorkbook.Close SaveChanges:=False
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Данные проанализированы успешно!" & vbNewLine & _
           "Столбец B – ТС завершающиеся" & vbNewLine & _
           "Столбец D – ДЛ завершающиеся" & vbNewLine & _
           "Столбец F – активные ТС в интервалах (на " & Format(Date, "dd.mm.yyyy") & ")", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка: " & Err.Description, vbCritical
    If Not sourceWorkbook Is Nothing Then
        On Error Resume Next
        sourceWorkbook.Close SaveChanges:=False
    End If
    Resume CleanUp
End Sub

