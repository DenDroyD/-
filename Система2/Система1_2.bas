Attribute VB_Name = "Система1_2"
Sub ImportDataFromExternalFile_System1_2()
    Dim wbThis As Workbook, wbSource As Workbook
    Dim wsTarget As Worksheet
    Dim sPath As String
    Dim wsSource_Scor As Worksheet, wsSource_Bukh As Worksheet, wsSource_Egrul As Worksheet, wsSource_Org As Worksheet
    Dim sourceValue As Variant
    Dim strC11 As String, pos As Integer
    Dim egrulResult As String, egrulResult2 As String
    Dim sumValue As Double
    Dim sumForD164 As Double, sumD145_147 As Double
    
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
    
    ' Проверяем существование листа "Система 1-2"
    On Error Resume Next
    Set wsTarget = wbThis.Sheets("Система 1-2")
    On Error GoTo ErrorHandler
    If wsTarget Is Nothing Then
        MsgBox "Лист 'Система 1-2' не найден в текущей книге!", vbCritical
        Exit Sub
    End If
    
    ' Получаем путь к папке текущего файла
    Dim currentFolder As String
    currentFolder = ThisWorkbook.Path & "\"

    ' Ищем файл, содержащий слово "Скоринг" в названии
    Dim searchPattern As String
    searchPattern = "*Скоринг*"

    Dim foundFile As String
    foundFile = Dir(currentFolder & searchPattern & ".xlsm") ' Сначала ищем .xlsm файлы

    ' Если не найдено .xlsm файлов, ищем другие форматы Excel
    If foundFile = "" Then
        foundFile = Dir(currentFolder & searchPattern & ".xlsx")
        If foundFile = "" Then
            foundFile = Dir(currentFolder & searchPattern & ".xls")
        End If
    End If

    ' Проверка наличия файла
    If foundFile = "" Then
        MsgBox "Файл, содержащий 'Скоринг' в названии, не найден в папке: " & currentFolder, vbCritical
        Exit Sub
    End If

    ' Полный путь к найденному файлу
    sPath = currentFolder & foundFile
    
    ' Открытие файла источника
    Set wbSource = Workbooks.Open(sPath, ReadOnly:=True)
    
    ' Проверка существования листов в файле-источнике
    On Error Resume Next
    Set wsSource_Scor = wbSource.Sheets("Скоринг")
    Set wsSource_Bukh = wbSource.Sheets("Бух.отч.")
    Set wsSource_Egrul = wbSource.Sheets("EGRUL")
    Set wsSource_Org = wbSource.Sheets("Organization Info")
    On Error GoTo ErrorHandler
    
    ' Дополнительная проверка существования листов
    If wsSource_Scor Is Nothing Then
        MsgBox "Лист 'Скоринг' не найден в файле: " & sPath & vbCrLf & _
               "Доступные листы: " & GetSheetNames(wbSource), vbCritical
        GoTo Cleanup
    End If
    
    If wsSource_Bukh Is Nothing Then
        MsgBox "Лист 'Бух.отч.' не найден в файле: " & sPath & vbCrLf & _
               "Доступные листы: " & GetSheetNames(wbSource), vbCritical
        GoTo Cleanup
    End If
    
    If wsSource_Egrul Is Nothing Then
        MsgBox "Лист 'EGRUL' не найден в файле: " & sPath & vbCrLf & _
               "Доступные листы: " & GetSheetNames(wbSource), vbCritical
        GoTo Cleanup
    End If
    
    If wsSource_Org Is Nothing Then
        MsgBox "Лист 'Organization Info' не найден в файле: " & sPath & vbCrLf & _
               "Доступные листы: " & GetSheetNames(wbSource), vbCritical
        GoTo Cleanup
    End If
    
    ' === ОБРАБОТКА ФОРМУЛ ИЗ FormulasList_Система_1-2.txt ===
    
    ' Ячейка C5
    wsTarget.Range("C5").Value = wsSource_Scor.Range("C7").Value
    
    ' Ячейка E5
    wsTarget.Range("E5").Value = wsSource_Scor.Range("C6").Value
    
    ' Ячейка D7
    wsTarget.Range("D7").Value = wsSource_Scor.Range("K2").Value
    
    ' Ячейка B8
    wsTarget.Range("B8").Value = wsSource_Scor.Range("C3").Value
    
    ' Ячейка B9
    wsTarget.Range("B9").Value = wsSource_Scor.Range("M2").Value
    
    ' Ячейка B10
    If wsSource_Scor.Range("C53").Value = 0 Then
        wsTarget.Range("B10").Value = ""
    Else
        wsTarget.Range("B10").Value = wsSource_Scor.Range("C53").Value
    End If
    
    ' Ячейка B11
    If wsSource_Scor.Range("C52").Value = 0 Then
        wsTarget.Range("B11").Value = ""
    Else
        wsTarget.Range("B11").Value = wsSource_Scor.Range("C52").Value
    End If
    
    ' Ячейка B12, оставил формулу
    'wsTarget.Range("B12").Value = wsSource_Scor.Range("U14").Value - wsTarget.Range("C47").Value
    
    ' Ячейка B13, теперь тянет общую
    wsTarget.Range("B13").Value = wsSource_Scor.Range("U14").Value
    
    ' Ячейка A15
    ' Эта ячейка зависит от других ячеек, поэтому обработаем позже
    
    ' Ячейка B18
    strC11 = wsSource_Scor.Range("C11").Value & " """
    pos = InStr(strC11, " """)
    If pos > 0 Then
        wsTarget.Range("B18").Value = Left(strC11, pos - 1)
    Else
        wsTarget.Range("B18").Value = strC11
    End If
    
    ' Ячейка B19
    strC11 = wsSource_Scor.Range("C11").Value & " """
    pos = InStr(strC11, " """)
    If pos > 0 Then
        wsTarget.Range("B19").Value = Mid(strC11, pos + 1)
    Else
        wsTarget.Range("B19").Value = ""
    End If
    
    ' Ячейка B20
    wsTarget.Range("B20").Value = wsSource_Scor.Range("C10").Value
    
    ' Ячейка B21
    wsTarget.Range("B21").Value = wsSource_Scor.Range("C13").Value
    
    ' Ячейка B23
    egrulResult = ""
    If Not IsEmpty(wsSource_Egrul.Range("C2").Value) And wsSource_Egrul.Range("C2").Value <> 0 Then
        egrulResult = egrulResult & Application.Proper(Trim(wsSource_Egrul.Range("A2").Value)) & " " & Trim(wsSource_Egrul.Range("C2").Value) & "%" & vbNewLine
    End If
    If Not IsEmpty(wsSource_Egrul.Range("C3").Value) And wsSource_Egrul.Range("C3").Value <> 0 Then
        egrulResult = egrulResult & Application.Proper(Trim(wsSource_Egrul.Range("A3").Value)) & " " & Trim(wsSource_Egrul.Range("C3").Value) & "%" & vbNewLine
    End If
    If Not IsEmpty(wsSource_Egrul.Range("C4").Value) And wsSource_Egrul.Range("C4").Value <> 0 Then
        egrulResult = egrulResult & Application.Proper(Trim(wsSource_Egrul.Range("A4").Value)) & " " & Trim(wsSource_Egrul.Range("C4").Value) & "%" & vbNewLine
    End If
    If Not IsEmpty(wsSource_Egrul.Range("C5").Value) And wsSource_Egrul.Range("C5").Value <> 0 Then
        egrulResult = egrulResult & Application.Proper(Trim(wsSource_Egrul.Range("A5").Value)) & " " & Trim(wsSource_Egrul.Range("C5").Value) & "%" & vbNewLine
    End If
    If Not IsEmpty(wsSource_Egrul.Range("C6").Value) And wsSource_Egrul.Range("C6").Value <> 0 Then
        egrulResult = egrulResult & Application.Proper(Trim(wsSource_Egrul.Range("A6").Value)) & " " & Trim(wsSource_Egrul.Range("C6").Value) & "%" & vbNewLine
    End If
    
    ' Удаление последнего символа новой строки, если он есть
    If Len(egrulResult) > 0 Then
        If Right(egrulResult, Len(vbNewLine)) = vbNewLine Then
            egrulResult = Left(egrulResult, Len(egrulResult) - Len(vbNewLine))
        End If
    End If
    
    wsTarget.Range("B23").Value = egrulResult
    
    ' Ячейка B24
    egrulResult2 = ""
    If wsSource_Egrul.Range("B2").Value <> "" Then
        egrulResult2 = Application.Proper(Trim(wsSource_Egrul.Range("A2").Value))
    End If
    If wsSource_Egrul.Range("B3").Value <> "" Then
        If egrulResult2 <> "" Then egrulResult2 = egrulResult2 & ", "
        egrulResult2 = egrulResult2 & Application.Proper(Trim(wsSource_Egrul.Range("A3").Value))
    End If
    If wsSource_Egrul.Range("B4").Value <> "" Then
        If egrulResult2 <> "" Then egrulResult2 = egrulResult2 & ", "
        egrulResult2 = egrulResult2 & Application.Proper(Trim(wsSource_Egrul.Range("A4").Value))
    End If
    If wsSource_Egrul.Range("B5").Value <> "" Then
        If egrulResult2 <> "" Then egrulResult2 = egrulResult2 & ", "
        egrulResult2 = egrulResult2 & Application.Proper(Trim(wsSource_Egrul.Range("A5").Value))
    End If
    If wsSource_Egrul.Range("B6").Value <> "" Then
        If egrulResult2 <> "" Then egrulResult2 = egrulResult2 & ", "
        egrulResult2 = egrulResult2 & Application.Proper(Trim(wsSource_Egrul.Range("A6").Value))
    End If
    wsTarget.Range("B24").Value = egrulResult2
    
    ' === ОБРАБОТКА ЯЧЕЙКИ B25 (АНАЛОГ C81) В ФОНОВОМ РЕЖИМЕ ===
    ' Ячейка B25
    Dim valueB25 As Variant
    Dim appOKVED As Object
    Dim wbOKVED As Workbook
    Dim okvedPath As String
    Dim orgInfoB2Value As Variant
    Dim okvedError As Boolean
    
    okvedError = False
    okvedPath = "S:\Transcend_disk_4\Credit Check\Для работы\Шаблон заключения\Авто\ОКВЭД.xlsx"
    
    ' Проверяем существование файла ОКВЭД
    If Dir(okvedPath) = "" Then
        wsTarget.Range("B25").Value = "Файл ОКВЭД не найден"
        okvedError = True
    Else
        orgInfoB2Value = wsSource_Org.Range("B2").Value
        
        ' Создаем отдельное приложение Excel в фоновом режиме
        On Error Resume Next
        Set appOKVED = CreateObject("Excel.Application")
        appOKVED.Visible = False
        appOKVED.screenUpdating = False
        appOKVED.enableEvents = False
        appOKVED.DisplayAlerts = False
        
        ' Открываем файл ОКВЭД
        Set wbOKVED = appOKVED.Workbooks.Open(okvedPath, ReadOnly:=True)
        If wbOKVED Is Nothing Then
            okvedError = True
        Else
            ' Выполняем VLOOKUP
            On Error Resume Next
            valueB25 = Application.VLookup(orgInfoB2Value, wbOKVED.Sheets("ОКВЭД 2").Range("B4:C2841"), 2, False)
            If Err.Number <> 0 Then
                valueB25 = CVErr(xlErrNA)
            End If
            On Error GoTo 0
            
            ' Закрываем файл ОКВЭД
            wbOKVED.Close False
            appOKVED.Quit
        End If
    End If
    
    ' Обрабатываем результат
    If Not okvedError And Not IsError(valueB25) Then
        ' Сохраняем значение как текст, чтобы сохранить точку
        wsTarget.Range("B25").NumberFormat = "@" ' Текстовый формат
        wsTarget.Range("B25").Value = CStr(valueB25)
    Else
        wsTarget.Range("B25").Value = "Не найдено"
    End If
    
    ' Очищаем объекты
    If Not wbOKVED Is Nothing Then
        On Error Resume Next
        wbOKVED.Close False
    End If
    If Not appOKVED Is Nothing Then
        On Error Resume Next
        appOKVED.Quit
    End If
    Set wbOKVED = Nothing
    Set appOKVED = Nothing
    
    ' Ячейка E25
    ' Эта ячейка зависит от данных из файла ОКВЭД, но мы не будем обрабатывать ее напрямую
    ' Пользователь может использовать формулы в Excel для вычисления этого значения
    
    ' Ячейка B28
    wsTarget.Range("B28").Value = wsSource_Org.Range("B4").Value
    
    ' Ячейка B33 (Первый ПЛ)
    wsTarget.Range("B33").Value = wsSource_Scor.Range("E6").Value & " " & wsSource_Scor.Range("G6").Value & " " & _
                                  wsSource_Scor.Range("H6").Value
    
    ' Ячейка E33 (Второй ПЛ)
    wsTarget.Range("E33").Value = wsSource_Scor.Range("E7").Value & " " & wsSource_Scor.Range("G7").Value & " " & _
                                  wsSource_Scor.Range("H7").Value
    
    ' Ячейка B34
    wsTarget.Range("B34").Value = wsSource_Scor.Range("K6").Value
    
    ' Ячейка E34
    wsTarget.Range("E34").Value = wsSource_Scor.Range("K7").Value
    
    ' Ячейка B35
    wsTarget.Range("B35").Value = wsSource_Scor.Range("J6").Value
    
    ' Ячейка E35
    wsTarget.Range("E35").Value = wsSource_Scor.Range("J7").Value
    
    ' Ячейка B36
    wsTarget.Range("B36").Value = wsSource_Scor.Range("M6").Value
    
    ' Ячейка E36
    wsTarget.Range("E36").Value = wsSource_Scor.Range("M7").Value
    
    ' Ячейка B37
    wsTarget.Range("B37").Value = Application.WorksheetFunction.Ceiling_Math(wsSource_Scor.Range("U6").Value, 100000, 1)
    
    ' Ячейка E37
    wsTarget.Range("E37").Value = Application.WorksheetFunction.Ceiling_Math(wsSource_Scor.Range("U7").Value, 100000, 1)
    
    ' Ячейка B38
    wsTarget.Range("B38").Value = wsSource_Scor.Range("N6").Value
    
    ' Ячейка E38
    wsTarget.Range("E38").Value = wsSource_Scor.Range("N7").Value
    
    ' Ячейка B39
    wsTarget.Range("B39").Value = wsSource_Scor.Range("O6").Value
    
    ' Ячейка E39
    wsTarget.Range("E39").Value = wsSource_Scor.Range("O7").Value
    
    ' Ячейка B40
    wsTarget.Range("B40").Value = wsSource_Scor.Range("P6").Value
    
    ' Ячейка E40
    wsTarget.Range("E40").Value = wsSource_Scor.Range("P7").Value
    
    ' Ячейка B41
    If wsSource_Scor.Range("C17").Value = "Брокер" Then
        wsTarget.Range("B41").Value = wsSource_Scor.Range("C23").Value & " ИНН:" & wsSource_Scor.Range("C22").Value
    ElseIf wsSource_Scor.Range("C17").Value = "Поставщик (агент ЮЛ)" Or wsSource_Scor.Range("C17").Value = "Поставщик (агент ФЛ)" Then
        wsTarget.Range("B41").Value = wsSource_Scor.Range("C19").Value & " ИНН:" & wsSource_Scor.Range("C18").Value
    ElseIf wsSource_Scor.Range("C17").Value = "Маркетплейс" Then
        wsTarget.Range("B41").Value = wsSource_Scor.Range("C25").Value & " ИНН:" & wsSource_Scor.Range("C24").Value
    Else
        wsTarget.Range("B41").Value = wsSource_Scor.Range("C17").Value
    End If
    
    ' Ячейка E41
    If wsSource_Scor.Range("C17").Value = "Брокер" Then
        wsTarget.Range("E41").Value = wsSource_Scor.Range("C23").Value & " ИНН:" & wsSource_Scor.Range("C22").Value
    ElseIf wsSource_Scor.Range("C17").Value = "Поставщик (агент ЮЛ)" Or wsSource_Scor.Range("C17").Value = "Поставщик (агент ФЛ)" Then
        wsTarget.Range("E41").Value = wsSource_Scor.Range("C19").Value & " ИНН:" & wsSource_Scor.Range("C18").Value
    ElseIf wsSource_Scor.Range("C17").Value = "Маркетплейс" Then
        wsTarget.Range("E41").Value = wsSource_Scor.Range("C25").Value & " ИНН:" & wsSource_Scor.Range("C24").Value
    Else
        wsTarget.Range("E41").Value = wsSource_Scor.Range("C17").Value
    End If
    
    ' Ячейка B42
    wsTarget.Range("B42").Value = wsSource_Scor.Range("C26").Value
    
    ' Ячейка E42
    wsTarget.Range("E42").Value = wsSource_Scor.Range("C26").Value
    
    ' Ячейка B44
    wsTarget.Range("B45").Value = wsSource_Scor.Range("Q6").Value
    
    ' Ячейка E44
    wsTarget.Range("E45").Value = wsSource_Scor.Range("Q7").Value
    
    ' Ячейка B45
    wsTarget.Range("B46").Value = wsSource_Scor.Range("R6").Value
    
    ' Ячейка E45
    wsTarget.Range("E46").Value = wsSource_Scor.Range("R7").Value
    
    ' Ячейка B51
    wsTarget.Range("B52").Value = wsSource_Scor.Range("V6").Value
    
    ' Ячейка B52
    wsTarget.Range("B53").Value = wsSource_Scor.Range("W6").Value
    
    ' Ячейка E52
    wsTarget.Range("E53").Value = wsSource_Scor.Range("Y6").Value
    
    ' === Финансовые показатели из листа "Бух.отч." ===
    
    ' Ячейка B81
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("C:C"), _
        Application.WorksheetFunction.Match(2110, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("B81").Value = sourceValue
    
    ' Ячейка C81
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
        Application.WorksheetFunction.Match(2110, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("C81").Value = sourceValue
    
    ' Ячейка D81
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("E:E"), _
        Application.WorksheetFunction.Match(2110, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("D81").Value = sourceValue
    
    ' Ячейка E81, есть формула в заключении
    'If IsNumeric(wsTarget.Range("B80").Value) And wsTarget.Range("B80").Value <> 0 Then
      '  wsTarget.Range("E80").Value = (wsTarget.Range("C80").Value - wsTarget.Range("B80").Value) / wsTarget.Range("B80").Value
   ' Else
    '    wsTarget.Range("E80").Value = "Нет данных"
   ' End If
    
     'Ячейка B81
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("C:C"), _
        Application.WorksheetFunction.Match(2400, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("B82").Value = sourceValue
    
    ' Ячейка C81
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
        Application.WorksheetFunction.Match(2400, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("C82").Value = sourceValue
    
    ' Ячейка D81
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("E:E"), _
        Application.WorksheetFunction.Match(2400, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("D82").Value = sourceValue
    
    ' Ячейка E81, есть формула в заключении
    'If IsNumeric(wsTarget.Range("B81").Value) And wsTarget.Range("B81").Value <> 0 Then
    '    wsTarget.Range("E81").Value = (wsTarget.Range("C81").Value - wsTarget.Range("B81").Value) / wsTarget.Range("B81").Value
    'Else
     '   wsTarget.Range("E81").Value = "Нет данных"
    'End If
    
    ' Ячейка B82
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("C:C"), _
        Application.WorksheetFunction.Match(1300, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("B83").Value = sourceValue
    
    ' Ячейка C82
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
        Application.WorksheetFunction.Match(1300, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("C83").Value = sourceValue
    
    ' Ячейка D82
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("E:E"), _
        Application.WorksheetFunction.Match(1300, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("D83").Value = sourceValue
    
    ' Ячейка E82, есть формула в заключении
    'If IsNumeric(wsTarget.Range("B82").Value) And wsTarget.Range("B82").Value <> 0 Then
    '    wsTarget.Range("E82").Value = (wsTarget.Range("C82").Value - wsTarget.Range("B82").Value) / wsTarget.Range("B82").Value
   ' Else
    '    wsTarget.Range("E82").Value = "Нет данных"
   ' End If
    
    ' Ячейка B83
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("C:C"), _
        Application.WorksheetFunction.Match(1150, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("B84").Value = sourceValue
    
    ' Ячейка C83
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
        Application.WorksheetFunction.Match(1150, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("C84").Value = sourceValue
    
    ' Ячейка E83, есть формула в заключении
    'If IsNumeric(wsTarget.Range("B83").Value) And wsTarget.Range("B83").Value <> 0 Then
    '    wsTarget.Range("E83").Value = (wsTarget.Range("C83").Value - wsTarget.Range("B83").Value) / wsTarget.Range("B83").Value
    'Else
     '   wsTarget.Range("E83").Value = "Нет данных"
   ' End If
    
    ' Ячейка B84
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("C:C"), _
        Application.WorksheetFunction.Match(1230, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("B85").Value = sourceValue
    
    ' Ячейка C84
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
        Application.WorksheetFunction.Match(1230, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("C85").Value = sourceValue
    
    ' Ячейка E84, есть формула в заключении
    'If IsNumeric(wsTarget.Range("B84").Value) And wsTarget.Range("B84").Value <> 0 Then
     '   wsTarget.Range("E84").Value = (wsTarget.Range("C84").Value - wsTarget.Range("B84").Value) / wsTarget.Range("B84").Value
    'Else
     '   wsTarget.Range("E84").Value = "Нет данных"
   ' End If
    
    ' Ячейка B85
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("C:C"), _
        Application.WorksheetFunction.Match(1520, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("B86").Value = sourceValue
    
    ' Ячейка C85
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
        Application.WorksheetFunction.Match(1520, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("C86").Value = sourceValue
    
    ' Ячейка E85
    'If IsNumeric(wsTarget.Range("B85").Value) And wsTarget.Range("B85").Value <> 0 Then
     '   wsTarget.Range("E85").Value = (wsTarget.Range("C85").Value - wsTarget.Range("B85").Value) / wsTarget.Range("B85").Value
    'Else
     '   wsTarget.Range("E85").Value = "Нет данных"
    'End If
    
    ' === Для Чек - листа ===
    
    ' Ячейка F83
      On Error Resume Next
      sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
          Application.WorksheetFunction.Match(2200, wsSource_Bukh.Range("B:B"), 0))
      If Err.Number <> 0 Then sourceValue = ""
      On Error GoTo ErrorHandler
      wsTarget.Range("F83").Value = sourceValue
    
    ' Ячейка F84
      On Error Resume Next
      sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
          Application.WorksheetFunction.Match(1150, wsSource_Bukh.Range("B:B"), 0))
      If Err.Number <> 0 Then sourceValue = ""
      On Error GoTo ErrorHandler
      wsTarget.Range("F84").Value = sourceValue

    ' Ячейка F85
      On Error Resume Next
      sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
          Application.WorksheetFunction.Match(1160, wsSource_Bukh.Range("B:B"), 0))
      If Err.Number <> 0 Then sourceValue = ""
      On Error GoTo ErrorHandler
      wsTarget.Range("F85").Value = sourceValue
    
    ' === Дополнительные ячейки ===
    
    ' Ячейка B104
    If IsEmpty(wsSource_Scor.Range("C50").Value) Or wsSource_Scor.Range("C50").Value = "" Or wsSource_Scor.Range("C50").Value = "нет информации" Then
        wsTarget.Range("B105").Value = "Нет"
    Else
        wsTarget.Range("B105").Value = "Да"
    End If
    
    ' Ячейка E104
    On Error Resume Next
    Dim strC39 As String
    strC39 = wsSource_Scor.Range("C39").Value
    Dim posProsrochki As Integer
    posProsrochki = InStr(strC39, " просрочки")
    If posProsrochki > 0 Then
        wsTarget.Range("E105").Value = Left(strC39, posProsrochki - 1)
    Else
        wsTarget.Range("E105").Value = " "
    End If
    On Error GoTo ErrorHandler
    
    ' Ячейка B105
    On Error Resume Next
    Dim strC49 As String
    strC49 = wsSource_Scor.Range("C49").Value
    Dim posSpace1 As Integer, posSpace2 As Integer
    posSpace1 = InStr(strC49, " ")
    If posSpace1 > 0 Then
        posSpace2 = InStr(posSpace1 + 1, strC49, " ")
        If posSpace2 > 0 Then
            wsTarget.Range("B106").Value = Left(strC49, posSpace2 - 1)
        Else
            wsTarget.Range("B106").Value = " "
        End If
    Else
        wsTarget.Range("B105").Value = " "
    End If
    On Error GoTo ErrorHandler
    
    ' Ячейка B106
    If IsEmpty(wsSource_Scor.Range("C50").Value) Or wsSource_Scor.Range("C50").Value = 0 Then
        wsTarget.Range("B107").Value = ""
    Else
        wsTarget.Range("B107").Value = wsSource_Scor.Range("C50").Value
    End If
    
    ' Ячейка E106
    If IsEmpty(wsSource_Scor.Range("C40").Value) Or wsSource_Scor.Range("C40").Value = 0 Then
        wsTarget.Range("E107").Value = ""
    Else
        wsTarget.Range("E107").Value = wsSource_Scor.Range("C40").Value
    End If
    
    ' Ячейка B107, смена логики чтобы все работало
    wsTarget.Range("F108").Value = wsSource_Scor.Range("C48").Value
    'sumForD164 = 0
    'If IsNumeric(wsSource_Scor.Range("C48").Value) Then sumForD164 = sumForD164 + wsSource_Scor.Range("C48").Value
    'If IsNumeric(wsTarget.Range("B43").Value) And IsNumeric(wsTarget.Range("B35").Value) Then sumForD164 = sumForD164 + wsTarget.Range("B43").Value * wsTarget.Range("B35").Value
    'If IsNumeric(wsTarget.Range("E43").Value) And IsNumeric(wsTarget.Range("E35").Value) Then sumForD164 = sumForD164 + wsTarget.Range("E43").Value * wsTarget.Range("E35").Value
    'wsTarget.Range("B107").Value = sumForD164 / 1000
    
    ' Ячейка E107
    If IsNumeric(wsSource_Scor.Range("C41").Value) Then
        wsTarget.Range("E108").Value = wsSource_Scor.Range("C41").Value / 1000
    Else
        wsTarget.Range("E108").Value = 0
    End If
    
    ' Ячейка B108, формула в самой ячейке
    'sumD145_147 = 0
    'If IsNumeric(wsTarget.Range("F79").Value) Then sumD145_147 = sumD145_147 + wsTarget.Range("F79").Value
    'If IsNumeric(wsTarget.Range("F80").Value) Then sumD145_147 = sumD145_147 + wsTarget.Range("F80").Value
    'If sumD145_147 <> 0 Then
    '    wsTarget.Range("B108").Value = wsTarget.Range("B107").Value / sumD145_147
    'Else
     '   wsTarget.Range("B108").Value = ""
   ' End If
    
    ' Ячейка E108, формула в самой ячейке
    'If sumD145_147 <> 0 Then
    '    wsTarget.Range("E108").Value = wsTarget.Range("E107").Value / sumD145_147
    'Else
     '   wsTarget.Range("E108").Value = ""
    'End If
    
    ' Ячейка C110, формула в самой ячейке
    'wsTarget.Range("C110").Value = wsTarget.Range("B107").Value + wsTarget.Range("E107").Value
    
    ' Ячейка C111, формула в самой ячейке
    'wsTarget.Range("C111").Value = wsTarget.Range("B108").Value + wsTarget.Range("E108").Value
    
    ' Ячейка B130
    wsTarget.Range("B131").Value = wsSource_Scor.Range("C5").Value
    
    ' Ячейка E130
    wsTarget.Range("E131").Value = Date
    
    ' Закрытие файла источника
    wbSource.Close SaveChanges:=False
    
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
    
    If Err.Number = 9 Then
        MsgBox "Ошибка: " & Err.Description & vbCrLf & _
               "Вероятно, лист не найден в файле-источнике." & vbCrLf & _
               "Проверьте, что в файле " & sPath & " действительно есть необходимые листы", vbCritical
    ElseIf Err.Number = 1004 Then
        MsgBox "Ошибка доступа к ячейке. Проверьте, что файл-источник открыт и содержит необходимые данные.", vbCritical
    Else
        MsgBox "Произошла ошибка " & Err.Number & ": " & Err.Description, vbCritical
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

Function GetSheetNames(wb As Workbook) As String
    Dim ws As Worksheet
    Dim sheetNames As String
    
    For Each ws In wb.Worksheets
        If sheetNames <> "" Then sheetNames = sheetNames & ", "
        sheetNames = sheetNames & ws.name
    Next ws
    
    GetSheetNames = sheetNames
End Function

