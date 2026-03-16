Attribute VB_Name = "Система3"
Sub ImportDataFromExternalFile_System3()
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
    
    ' Проверяем существование листа "Система 3(Ф)"
    On Error Resume Next
    Set wsTarget = wbThis.Sheets("Система 3")
    On Error GoTo ErrorHandler
    If wsTarget Is Nothing Then
        MsgBox "Лист 'Система 3(Ф)' не найден в текущей книге!", vbCritical
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
    
    ' === ОБРАБОТКА ФОРМУЛ ИЗ formulas_export.txt ===
    
    ' Ячейка $G$2
    wsTarget.Range("G2").Value = wsSource_Scor.Range("C7").Value
    
    ' Ячейка $G$3
    wsTarget.Range("G3").Value = wsSource_Scor.Range("C6").Value
    
    ' Ячейка $E$5
    wsTarget.Range("E5").Value = wsSource_Scor.Range("K2").Value
    
    ' Ячейка $B$5
    wsTarget.Range("B5").Value = wsSource_Scor.Range("C4").Value
    
    ' Ячейка $B$6
    wsTarget.Range("B6").Value = wsSource_Scor.Range("C3").Value
    
    ' Ячейка $B$7
    wsTarget.Range("B7").Value = wsSource_Scor.Range("M2").Value
    
    ' Ячейка $B$8
    If wsSource_Scor.Range("C53").Value = 0 Then
        wsTarget.Range("B8").Value = ""
    Else
        wsTarget.Range("B8").Value = wsSource_Scor.Range("C53").Value
    End If
    
    ' Ячейка $B$9
    If wsSource_Scor.Range("C52").Value = 0 Then
        wsTarget.Range("B9").Value = ""
    Else
        wsTarget.Range("B9").Value = wsSource_Scor.Range("C52").Value
    End If
    
    ' Ячейка $B$10 (зависит от C18, которую мы еще заполним)
    ' Сначала заполним C18, потом B10
    ' Ячейка $C$18
    sumValue = 0
    On Error Resume Next
    sumValue = sumValue + wsSource_Scor.Range("U6").Value
    sumValue = sumValue + wsSource_Scor.Range("U7").Value
    sumValue = sumValue + wsSource_Scor.Range("U8").Value
    sumValue = sumValue + wsSource_Scor.Range("U9").Value
    sumValue = sumValue + wsSource_Scor.Range("U10").Value
    sumValue = sumValue + wsSource_Scor.Range("U11").Value
    sumValue = sumValue + wsSource_Scor.Range("U12").Value
    sumValue = sumValue + wsSource_Scor.Range("U13").Value
    On Error GoTo ErrorHandler
    
    wsTarget.Range("C18").Value = Application.WorksheetFunction.Ceiling_Math(sumValue, 100000, 1)
    
    ' Теперь можем заполнить B10
    wsTarget.Range("B10").Value = wsTarget.Range("C18").Value
    
    ' Ячейка $B$11
    wsTarget.Range("B11").Value = wsSource_Scor.Range("U14").Value
    
    ' Ячейка $C$17
    wsTarget.Range("C17").Value = wsSource_Scor.Range("S14").Value
    
    ' Ячейка $C$19
    wsTarget.Range("C19").Value = wsSource_Scor.Range("J14").Value
    
    ' Ячейка $C$20 Первый ПЛ
    wsTarget.Range("C20").Value = wsSource_Scor.Range("E6").Value & " " & wsSource_Scor.Range("G6").Value & " " & _
                              wsSource_Scor.Range("H6").Value & ", стоимостью " & _
                              Format(wsSource_Scor.Range("K6").Value, "### ### ###") & " рублей"
    
    ' Ячейка $C$21
    wsTarget.Range("C21").Value = wsSource_Scor.Range("M6").Value
    
    ' Ячейка $C$22
    wsTarget.Range("C22").Value = wsSource_Scor.Range("N6").Value
    
    ' Ячейка $C$23
    wsTarget.Range("C23").Value = wsSource_Scor.Range("P6").Value
    
    ' Ячейка $C$24
    wsTarget.Range("C24").Value = wsSource_Scor.Range("O6").Value
    
    ' Ячейка $C$26
    wsTarget.Range("C26").Value = wsSource_Scor.Range("Q6").Value
    
    ' Ячейка $C$27
    wsTarget.Range("C27").Value = wsSource_Scor.Range("R6").Value
    
    ' Ячейка $C$31 Второй ПЛ
    wsTarget.Range("C31").Value = wsSource_Scor.Range("E7").Value & " " & wsSource_Scor.Range("G7").Value & " " & _
                                  wsSource_Scor.Range("H7").Value & ", стоимостью " & _
                                  Format(wsSource_Scor.Range("K7").Value, "### ### ###") & " рублей"
    
    ' Ячейка $C$32
    wsTarget.Range("C32").Value = wsSource_Scor.Range("M7").Value
    
    ' Ячейка $C$33
    wsTarget.Range("C33").Value = wsSource_Scor.Range("N7").Value
    
    ' Ячейка $C$34
    wsTarget.Range("C34").Value = wsSource_Scor.Range("P7").Value
    
    ' Ячейка $C$35
    wsTarget.Range("C35").Value = wsSource_Scor.Range("O7").Value
    
    ' Ячейка $C$37
    wsTarget.Range("C37").Value = wsSource_Scor.Range("Q7").Value
    
    ' Ячейка $C$38
    wsTarget.Range("C38").Value = wsSource_Scor.Range("R7").Value
    
    ' Ячейка $C$42 Третий ПЛ
    wsTarget.Range("C42").Value = wsSource_Scor.Range("E8").Value & " " & wsSource_Scor.Range("G8").Value & " " & _
                                  wsSource_Scor.Range("H8").Value & ", стоимостью " & _
                                  Format(wsSource_Scor.Range("K8").Value, "### ### ###") & " рублей"
    
    ' Ячейка $C$43
    wsTarget.Range("C43").Value = wsSource_Scor.Range("M8").Value
    
    ' Ячейка $C$44
    wsTarget.Range("C44").Value = wsSource_Scor.Range("N8").Value
    
    ' Ячейка $C$45
    wsTarget.Range("C45").Value = wsSource_Scor.Range("P8").Value
    
    ' Ячейка $C$46
    wsTarget.Range("C46").Value = wsSource_Scor.Range("O8").Value
    
    ' Ячейка $C$48
    wsTarget.Range("C48").Value = wsSource_Scor.Range("Q8").Value
    
    ' Ячейка $C$49
    wsTarget.Range("C49").Value = wsSource_Scor.Range("R8").Value
    
    ' Ячейка $C$53 Четвертый ПЛ
    wsTarget.Range("C53").Value = wsSource_Scor.Range("E9").Value & " " & wsSource_Scor.Range("G9").Value & " " & _
                                  wsSource_Scor.Range("H9").Value & ", стоимостью " & _
                                  Format(wsSource_Scor.Range("K9").Value, "### ### ###") & " рублей"
    
    ' Ячейка $C$54
    wsTarget.Range("C54").Value = wsSource_Scor.Range("M9").Value
    
    ' Ячейка $C$55
    wsTarget.Range("C55").Value = wsSource_Scor.Range("N9").Value
    
    ' Ячейка $C$56
    wsTarget.Range("C56").Value = wsSource_Scor.Range("P9").Value
    
    ' Ячейка $C$57
    wsTarget.Range("C57").Value = wsSource_Scor.Range("O9").Value
    
    ' Ячейка $C$59
    wsTarget.Range("C59").Value = wsSource_Scor.Range("Q9").Value
    
    ' Ячейка $C$64
    wsTarget.Range("C64").Value = wsSource_Scor.Range("C17").Value
    
    ' Ячейка $C$65
    If wsSource_Scor.Range("C17").Value = "Брокер" Then
        wsTarget.Range("C65").Value = wsSource_Scor.Range("C23").Value & " ИНН:" & wsSource_Scor.Range("C22").Value
    ElseIf wsSource_Scor.Range("C17").Value = "Поставщик (агент ЮЛ)" Or wsSource_Scor.Range("C17").Value = "Поставщик (агент ФЛ)" Then
        wsTarget.Range("C65").Value = wsSource_Scor.Range("C19").Value & " ИНН:" & wsSource_Scor.Range("C18").Value
    ElseIf wsSource_Scor.Range("C17").Value = "Маркетплейс" Then
        wsTarget.Range("C65").Value = wsSource_Scor.Range("C25").Value & " ИНН:" & wsSource_Scor.Range("C24").Value
    Else
        wsTarget.Range("C65").Value = wsSource_Scor.Range("C17").Value
    End If
    
    ' Ячейка $C$66
    wsTarget.Range("C66").Value = wsSource_Scor.Range("C26").Value
    
    ' Ячейка $C$71
    strC11 = wsSource_Scor.Range("C11").Value & " """
    pos = InStr(strC11, " """)
    If pos > 0 Then
        wsTarget.Range("C71").Value = Mid(strC11, pos + 1)
    Else
        wsTarget.Range("C71").Value = ""
    End If
    
    ' Ячейка $C$72
    strC11 = wsSource_Scor.Range("C11").Value & " """
    pos = InStr(strC11, " """)
    If pos > 0 Then
        wsTarget.Range("C72").Value = Left(strC11, pos - 1)
    Else
        wsTarget.Range("C72").Value = strC11
    End If
    
    ' Ячейка $C$73
    wsTarget.Range("C73").Value = wsSource_Scor.Range("C10").Value
    
    ' Ячейка $C$74
    wsTarget.Range("C74").Value = wsSource_Scor.Range("C13").Value
    
    ' Ячейка $C$77 убрал, формула в самом заключении
    
    ' Ячейка $C$79
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

wsTarget.Range("C79").Value = egrulResult
    
    ' Ячейка $C$80
    egrulResult2 = ""
    If wsSource_Egrul.Range("B2").Value <> "" Then
        egrulResult2 = Application.Proper(wsSource_Egrul.Range("A2").Value)
    End If
    If wsSource_Egrul.Range("B3").Value <> "" Then
        If egrulResult2 <> "" Then egrulResult2 = egrulResult2 & ", "
        egrulResult2 = egrulResult2 & Application.Proper(wsSource_Egrul.Range("A3").Value)
    End If
    If wsSource_Egrul.Range("B4").Value <> "" Then
        If egrulResult2 <> "" Then egrulResult2 = egrulResult2 & ", "
        egrulResult2 = egrulResult2 & Application.Proper(wsSource_Egrul.Range("A4").Value)
    End If
    If wsSource_Egrul.Range("B5").Value <> "" Then
        If egrulResult2 <> "" Then egrulResult2 = egrulResult2 & ", "
        egrulResult2 = egrulResult2 & Application.Proper(wsSource_Egrul.Range("A5").Value)
    End If
    If wsSource_Egrul.Range("B6").Value <> "" Then
        If egrulResult2 <> "" Then egrulResult2 = egrulResult2 & ", "
        egrulResult2 = egrulResult2 & Application.Proper(wsSource_Egrul.Range("A6").Value)
    End If
    wsTarget.Range("C80").Value = egrulResult2
    
    ' === ОБРАБОТКА ЯЧЕЙКИ $C$81 В ФОНОВОМ РЕЖИМЕ ===
    ' Ячейка "$C$81"
    Dim valueC81 As Variant
    Dim appOKVED As Object
    Dim wbOKVED As Workbook
    Dim okvedPath As String
    Dim orgInfoB2Value As Variant
    Dim okvedError As Boolean
    
    okvedError = False
    okvedPath = "S:\Transcend_disk_4\Credit Check\Для работы\Шаблон заключения\Авто\ОКВЭД.xlsx"
    
    ' Проверяем существование файла ОКВЭД
    If Dir(okvedPath) = "" Then
        wsTarget.Range("C81").Value = "Файл ОКВЭД не найден"
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
            valueC81 = Application.VLookup(orgInfoB2Value, wbOKVED.Sheets("ОКВЭД 2").Range("B4:C2841"), 2, False)
            If Err.Number <> 0 Then
                valueC81 = CVErr(xlErrNA)
            End If
            On Error GoTo 0
            
            ' Закрываем файл ОКВЭД
            wbOKVED.Close False
            appOKVED.Quit
        End If
    End If
    
    ' Обрабатываем результат
    If Not okvedError And Not IsError(valueC81) Then
        ' Сохраняем значение как текст, чтобы сохранить точку
        wsTarget.Range("C81").NumberFormat = "@" ' Текстовый формат
        wsTarget.Range("C81").Value = CStr(valueC81)
    Else
        wsTarget.Range("C81").Value = "Не найдено"
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
    
    ' Ячейка $C$84
    wsTarget.Range("C84").Value = wsSource_Org.Range("B4").Value
    
    ' Ячейка $C$85 убрал, формула в самом заключении
    
    ' Ячейка $C$91
    wsTarget.Range("C91").Value = wsSource_Scor.Range("V6").Value
    
    ' Ячейка $C$92
    wsTarget.Range("C92").Value = wsSource_Scor.Range("W6").Value
    
    ' Ячейка $H$92
    wsTarget.Range("H92").Value = wsSource_Scor.Range("Y6").Value
    
    ' Ячейка $G$176
    wsTarget.Range("G176").Value = wsSource_Scor.Range("C5").Value
    
    ' Ячейка $G$177
    wsTarget.Range("G177").Value = Date
    
    ' === Финансовые показатели из листа "Бух.отч." ===
    
    ' Ячейка $B$144
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("C:C"), _
        Application.WorksheetFunction.Match(2110, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("B144").Value = sourceValue
    
    ' Ячейка $E$144
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("C:C"), _
        Application.WorksheetFunction.Match(2400, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("E144").Value = sourceValue
    
    ' Ячейка $G$144
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("C:C"), _
        Application.WorksheetFunction.Match(1300, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("G144").Value = sourceValue
    
    ' Ячейка $H$144
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("C:C"), _
        Application.WorksheetFunction.Match(1600, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("H144").Value = sourceValue
    
    ' Ячейка $B$145
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
        Application.WorksheetFunction.Match(2110, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("B145").Value = sourceValue
    
    ' Ячейка $E$145
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
        Application.WorksheetFunction.Match(2400, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("E145").Value = sourceValue
    
    ' Ячейка $G$145
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
        Application.WorksheetFunction.Match(1300, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("G145").Value = sourceValue
    
    ' Ячейка $H$145
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
        Application.WorksheetFunction.Match(1600, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("H145").Value = sourceValue
    
    ' Ячейка $B$147
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("E:E"), _
        Application.WorksheetFunction.Match(2110, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("B147").Value = sourceValue
    
    ' Ячейка $E$147
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("E:E"), _
        Application.WorksheetFunction.Match(2400, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("E147").Value = sourceValue
    
    ' Ячейка $G$147
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("E:E"), _
        Application.WorksheetFunction.Match(1300, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("G147").Value = sourceValue
    
    ' Ячейка $H$147
    On Error Resume Next
    sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("E:E"), _
        Application.WorksheetFunction.Match(1600, wsSource_Bukh.Range("B:B"), 0))
    If Err.Number <> 0 Then sourceValue = ""
    On Error GoTo ErrorHandler
    wsTarget.Range("H147").Value = sourceValue
    
       ' === Для Чек - листа ===
    
    ' Ячейка I144
      On Error Resume Next
      sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
          Application.WorksheetFunction.Match(2200, wsSource_Bukh.Range("B:B"), 0))
      If Err.Number <> 0 Then sourceValue = ""
      On Error GoTo ErrorHandler
      wsTarget.Range("I144").Value = sourceValue
    
    ' Ячейка I145
      On Error Resume Next
      sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
          Application.WorksheetFunction.Match(1150, wsSource_Bukh.Range("B:B"), 0))
      If Err.Number <> 0 Then sourceValue = ""
      On Error GoTo ErrorHandler
      wsTarget.Range("I145").Value = sourceValue

    ' Ячейка I146
      On Error Resume Next
      sourceValue = Application.WorksheetFunction.Index(wsSource_Bukh.Range("D:D"), _
          Application.WorksheetFunction.Match(1160, wsSource_Bukh.Range("B:B"), 0))
      If Err.Number <> 0 Then sourceValue = ""
      On Error GoTo ErrorHandler
      wsTarget.Range("I146").Value = sourceValue
    
    
    ' === Дополнительные вычисления для финансовых показателей, переведены в формулы===
    
    ' Ячейка $D$144
    'If IsNumeric(wsTarget.Range("B144").Value) And wsTarget.Range("B144").Value <> 0 Then
       ' wsTarget.Range("D144").Value = wsTarget.Range("B144").Value / 12
   ' Else
      '  wsTarget.Range("D144").Value = ""
  '  End If
    
    ' Ячейка $F$144
   ' If IsNumeric(wsTarget.Range("E144").Value) And wsTarget.Range("E144").Value <> 0 Then
   '     wsTarget.Range("F144").Value = wsTarget.Range("E144").Value / 12
   ' Else
    '    wsTarget.Range("F144").Value = ""
   ' End If
    
    ' Ячейка $D$145
   ' If IsNumeric(wsTarget.Range("B145").Value) And wsTarget.Range("B145").Value <> 0 Then
   '     wsTarget.Range("D145").Value = wsTarget.Range("B145").Value / 12
   ' Else
   '     wsTarget.Range("D145").Value = ""
   ' End If
  '
    ' Ячейка $F$145
   ' If IsNumeric(wsTarget.Range("E145").Value) And wsTarget.Range("E145").Value <> 0 Then
   '     wsTarget.Range("F145").Value = wsTarget.Range("E145").Value / 12
   ' Else
   '     wsTarget.Range("F145").Value = ""
    'End If
    
    ' Ячейка $D$147
    'If IsNumeric(wsTarget.Range("B147").Value) And wsTarget.Range("B147").Value <> 0 Then
   '     wsTarget.Range("D147").Value = wsTarget.Range("B147").Value / 12
   ' Else
   '     wsTarget.Range("D147").Value = ""
  '  End If
    
    ' Ячейка $F$147
   ' If IsNumeric(wsTarget.Range("E147").Value) And wsTarget.Range("E147").Value <> 0 Then
   '     wsTarget.Range("F147").Value = wsTarget.Range("E147").Value / 12
   ' Else
   '     wsTarget.Range("F147").Value = ""
 '   End If
    
    ' Ячейка $B$148
   ' If IsNumeric(wsTarget.Range("B146").Value) And wsTarget.Range("B146").Value <> 0 Then
   '     wsTarget.Range("B148").Value = (wsTarget.Range("B147").Value - wsTarget.Range("B146").Value) / wsTarget.Range("B146").Value
   ' Else
   '     wsTarget.Range("B148").Value = "Нет данных"
   ' End If
    
    ' Ячейка $D$148
  '  If IsNumeric(wsTarget.Range("D146").Value) And wsTarget.Range("D146").Value <> 0 Then
   '     wsTarget.Range("D148").Value = (wsTarget.Range("D147").Value - wsTarget.Range("D146").Value) / wsTarget.Range("D146").Value
   ' Else
   '     wsTarget.Range("D148").Value = "Нет данных"
   ' End If
    
    ' Ячейка $E$148
   ' If IsNumeric(wsTarget.Range("E146").Value) And wsTarget.Range("E146").Value <> 0 Then
   '     wsTarget.Range("E148").Value = (wsTarget.Range("E147").Value - wsTarget.Range("E146").Value) / wsTarget.Range("E146").Value
   ' Else
  '      wsTarget.Range("E148").Value = "Нет данных"
   ' End If
    
    ' Ячейка $F$148
   ' If IsNumeric(wsTarget.Range("F146").Value) And wsTarget.Range("F146").Value <> 0 Then
   '     wsTarget.Range("F148").Value = (wsTarget.Range("F147").Value - wsTarget.Range("F146").Value) / wsTarget.Range("F146").Value
   ' Else
   '     wsTarget.Range("F148").Value = "Нет данных"
   ' End If
    
    ' Ячейка $G$148
    'If IsNumeric(wsTarget.Range("G146").Value) And wsTarget.Range("G146").Value <> 0 Then
   '     wsTarget.Range("G148").Value = (wsTarget.Range("G147").Value - wsTarget.Range("G146").Value) / wsTarget.Range("G146").Value
   ' Else
    '    wsTarget.Range("G148").Value = "Нет данных"
   ' End If
    
    ' Ячейка $H$148
   ' If IsNumeric(wsTarget.Range("H146").Value) And wsTarget.Range("H146").Value <> 0 Then
    '    wsTarget.Range("H148").Value = (wsTarget.Range("H147").Value - wsTarget.Range("H146").Value) / wsTarget.Range("H146").Value
   ' Else
     '   wsTarget.Range("H148").Value = "Нет данных"
    'End If
    
    ' === Дополнительные ячейки ===
    
    ' Ячейка $D$161
    If IsEmpty(wsSource_Scor.Range("C50").Value) Or wsSource_Scor.Range("C50").Value = "" Or wsSource_Scor.Range("C50").Value = "нет информации" Then
        wsTarget.Range("D161").Value = "Нет"
    Else
        wsTarget.Range("D161").Value = "Да"
    End If
    
    ' Ячейка $H$161
    On Error Resume Next
    Dim strC39 As String
    strC39 = wsSource_Scor.Range("C39").Value
    Dim posProsrochki As Integer
    posProsrochki = InStr(strC39, " просрочки")
    If posProsrochki > 0 Then
        wsTarget.Range("H161").Value = Left(strC39, posProsrochki - 1)
    Else
        wsTarget.Range("H161").Value = " "
    End If
    On Error GoTo ErrorHandler
    
    ' Ячейка $D$162
    On Error Resume Next
    Dim strC49 As String
    strC49 = wsSource_Scor.Range("C49").Value
    Dim posSpace1 As Integer, posSpace2 As Integer
    posSpace1 = InStr(strC49, " ")
    If posSpace1 > 0 Then
        posSpace2 = InStr(posSpace1 + 1, strC49, " ")
        If posSpace2 > 0 Then
            wsTarget.Range("D162").Value = Left(strC49, posSpace2 - 1)
        Else
            wsTarget.Range("D162").Value = " "
        End If
    Else
        wsTarget.Range("D162").Value = " "
    End If
    On Error GoTo ErrorHandler
    
    ' Ячейка $D$163
    If IsEmpty(wsSource_Scor.Range("C50").Value) Or wsSource_Scor.Range("C50").Value = 0 Then
        wsTarget.Range("D163").Value = ""
    Else
        wsTarget.Range("D163").Value = wsSource_Scor.Range("C50").Value
    End If
    
    ' Ячейка $H$163
    If IsEmpty(wsSource_Scor.Range("C40").Value) Or wsSource_Scor.Range("C40").Value = 0 Then
        wsTarget.Range("H163").Value = ""
    Else
        wsTarget.Range("H163").Value = wsSource_Scor.Range("C40").Value
    End If
    
    ' Ячейка $D$164, результат перенесен в $I$164, в самой ячейке $D$164 будет формула в заключении
     wsTarget.Range("I164").Value = wsSource_Scor.Range("C48").Value
    
    ' Ячейка $H$164
    If IsNumeric(wsSource_Scor.Range("C41").Value) Then
        wsTarget.Range("H164").Value = wsSource_Scor.Range("C41").Value / 1000
    Else
        wsTarget.Range("H164").Value = 0
    End If
    
    ' Ячейка $D$165 , выведена в формулу в ячейке
    'sumD145_147 = 0
    'If IsNumeric(wsTarget.Range("D145").Value) Then sumD145_147 = sumD145_147 + wsTarget.Range("D145").Value
    'If IsNumeric(wsTarget.Range("D147").Value) Then sumD145_147 = sumD145_147 + wsTarget.Range("D147").Value
    'If sumD145_147 <> 0 Then
        'wsTarget.Range("D165").Value = wsTarget.Range("D164").Value / sumD145_147
    'Else
        'wsTarget.Range("D165").Value = ""
    'End If
    
    ' Ячейка $H$165 , выведена в формулу в ячейке
    'If sumD145_147 <> 0 Then
      '  wsTarget.Range("H165").Value = wsTarget.Range("H164").Value / sumD145_147
   ' Else
      '  wsTarget.Range("H165").Value = ""
   ' End If
    
    ' Ячейка $D$167, выведена в формулу в ячейке
    'wsTarget.Range("D167").Value = wsTarget.Range("D164").Value + wsTarget.Range("H164").Value
    
    ' Ячейка $D$168, выведена в формулу в ячейке
    'If IsNumeric(wsTarget.Range("D165").Value) And IsNumeric(wsTarget.Range("H165").Value) Then
        'wsTarget.Range("D168").Value = wsTarget.Range("D165").Value + wsTarget.Range("H165").Value
    'Else
        'wsTarget.Range("D168").Value = ""
   ' End If
    
        ' === УПРАВЛЕНИЕ ВИДИМОСТЬЮ СТРОК ===
    ' Сначала скрываем все потенциально скрытые строки
    wsTarget.Rows("31:41").Hidden = True
    wsTarget.Rows("42:52").Hidden = True
    wsTarget.Rows("53:63").Hidden = True
    
    ' Раскрываем строки в зависимости от заполнения ячеек
    With wsTarget
        ' Если ячейка C32 заполнена, показываем строки 31-38
        If Not IsEmpty(.Range("C32").Value) And Trim(.Range("C32").Value) <> "" Then
            .Rows("31:41").Hidden = False
        End If
        
        ' Если ячейка C43 заполнена, показываем строки 42-49
        If Not IsEmpty(.Range("C43").Value) And Trim(.Range("C43").Value) <> "" Then
            .Rows("42:52").Hidden = False
        End If
        
        ' Если ячейка C53 заполнена, показываем строки 53-64
        If Not IsEmpty(.Range("C54").Value) And Trim(.Range("C54").Value) <> "" Then
            .Rows("53:63").Hidden = False
        End If
    End With
    
    ' Закрытие файла источника
    wbSource.Close SaveChanges:=False
    
    ' Восстановление настроек Excel
    Application.screenUpdating = screenUpdating
    Application.calculation = calculation
    Application.enableEvents = enableEvents
    
    MsgBox "Данные успешно загружены!", vbInformation
    Exit Sub
    
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

