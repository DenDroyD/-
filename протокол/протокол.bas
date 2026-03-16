Attribute VB_Name = "протокол"
Sub СоздатьКопиюПервогоЛистаБезМакросов()
    Dim newFileName As String
    Dim savePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim mkckValue As String
    Dim dealText As String
    Dim i As Long

    Application.screenUpdating = False
    Application.calculation = xlCalculationManual
    Application.enableEvents = False

    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Sheets(1) ' Первый лист

    ' Получаем значение из E1 (ожидается "МКК" или "БКК")
    mkckValue = Trim(ws.Range("E1").Value)
    If mkckValue = "" Then mkckValue = "БКК" ' резервное значение

    ' Читаем текст из A10 (объединённая ячейка A10:D10)
    Dim fullText As String
    fullText = Trim(ws.Range("A10").Value)
    If fullText = "" Then
        MsgBox "Ячейка A10 пуста. Невозможно сформировать имя файла.", vbExclamation
        GoTo Finalize
    End If

    ' Извлекаем фрагмент от "сделка"/"сделке" до ИНН + 10-12 цифр
    dealText = ExtractDealFragment(fullText)
    If dealText = "" Then
        MsgBox "Не удалось извлечь информацию о сделке из ячейки A10.", vbExclamation
        GoTo Finalize
    End If

    ' Корректируем "сделка" ? "сделке"
    If InStr(1, dealText, "сделка", vbTextCompare) > 0 Then
        dealText = Replace(dealText, "сделка", "сделке", 1, 1, vbTextCompare)
    End If

    ' Формируем имя файла
    If Left(LCase(dealText), 2) = "по" Then
        newFileName = "Протокол " & mkckValue & " " & dealText & ".xlsx"
    Else
        newFileName = "Протокол " & mkckValue & " по " & dealText & ".xlsx"
    End If

    ' Очищаем имя от недопустимых символов
    newFileName = CleanFileName(newFileName)

    ' Путь сохранения - та же папка
    savePath = ThisWorkbook.Path & "\" & newFileName

    ' Копируем только первый лист в новую книгу
    ws.Copy
    Set wb = ActiveWorkbook

    ' Удаляем фигуры (все, кроме диаграмм)
    For i = wb.Sheets(1).Shapes.count To 1 Step -1
        Set shp = wb.Sheets(1).Shapes(i)
        If shp.Type <> msoChart Then
            shp.Delete
        End If
    Next i

    ' === УДАЛЯЕМ СТОЛБЦЫ E, F и G ===
    With wb.Sheets(1)
        .Columns("E:G").Delete Shift:=xlToLeft
    End With

    ' Сохраняем как XLSX (без макросов)
    Application.DisplayAlerts = False
    wb.SaveAs fileName:=savePath, fileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True

    MsgBox "Файл успешно сохранён как:" & vbCrLf & newFileName, vbInformation, "Готово"

Finalize:
    Application.screenUpdating = True
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    MsgBox "Ошибка: " & Err.Description, vbCritical, "Произошла ошибка"
    Resume Finalize
End Sub


' === Извлечение фрагмента: от "сделка"/"сделке" до ИНН + 10-12 цифр ===
Function ExtractDealFragment(text As String) As String
    Dim startIdx As Long, endIdx As Long
    Dim pos As Long, innStart As Long
    Dim digits As String
    Dim count As Integer

    ' Находим начало слова "сделка" или "сделке"
    pos = InStr(1, LCase(text), "сделка")
    If pos = 0 Then
        pos = InStr(1, LCase(text), "сделке")
    End If

    If pos = 0 Then
        ExtractDealFragment = ""
        Exit Function
    End If

    ' Ищем "ИНН"
    innStart = InStr(pos, LCase(text), "инн")
    If innStart = 0 Then
        ExtractDealFragment = ""
        Exit Function
    End If

    ' Пропускаем "ИНН" и пробелы
    innStart = innStart + 3
    Do While innStart <= Len(text) And Not IsNumeric(Mid(text, innStart, 1))
        innStart = innStart + 1
    Loop

    If innStart > Len(text) Then
        ExtractDealFragment = ""
        Exit Function
    End If

    ' Собираем цифры после ИНН (10-12 цифр)
    digits = ""
    count = 0
    Do While innStart <= Len(text) And IsNumeric(Mid(text, innStart, 1)) And count < 12
        digits = digits & Mid(text, innStart, 1)
        innStart = innStart + 1
        count = count + 1
    Loop

    If Len(digits) < 10 Then
        ExtractDealFragment = ""
        Exit Function
    End If

    ' Конец фрагмента - последняя цифра ИНН
    endIdx = innStart - 1

    ' Извлекаем подстроку от начала "сделка" до конца ИНН
    ExtractDealFragment = Mid(text, pos, endIdx - pos + 1)
    ExtractDealFragment = Trim(ExtractDealFragment)
    ExtractDealFragment = Replace(ExtractDealFragment, "  ", " ") ' лишние пробелы
End Function


' === Очистка имени файла от недопустимых символов ===
Function CleanFileName(name As String) As String
    Dim invalidChars As String
    Dim i As Integer
    invalidChars = "\/:*?""<>|[]{}=;,+`~!@#$%^&"

    For i = 1 To Len(invalidChars)
        name = Replace(name, Mid(invalidChars, i, 1), "_")
    Next i

    name = Application.WorksheetFunction.Trim(name)
    name = Replace(name, " ", "_")
    name = Replace(name, "__", "_")
    name = Replace(name, "__", "_") ' повтор для безопасности

    CleanFileName = name
End Function

