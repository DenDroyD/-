Attribute VB_Name = "ЧекоАфф"
Sub GetAffCompanyFromChecko()
    Dim apiKey As String, url As String
    Dim httpReq As Object, jsonText As String
    Dim targetSheet As Worksheet
    Dim financialData As Object, companyData As Object, DataObj As Object
    Dim report2024 As Object
    Dim i As Long
    Dim startRow As Long, endRow As Long
    Dim additionalInn As String
    Dim outputColH As String, outputColI As String, outputColJ As String
    Dim fileDialog As fileDialog
    Dim selectedFile As String
    Dim fileNum As Integer
    Dim fileContent As String
    Dim requestCount As Long ' Счётчик запросов

    requestCount = 0
    
    ' Настройки
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Sheets("Система4")
    On Error GoTo ErrorHandler
    
    If targetSheet Is Nothing Then
        MsgBox "Лист 'Система4' не найден!", vbCritical
        Exit Sub
    End If

    ' === ЗАГРУЗКА API-КЛЮЧА ИЗ TXT-ФАЙЛА ===
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Выберите файл с API-ключом (текстовый файл .txt)"
        .Filters.Clear
        .Filters.Add "Текстовые файлы", "*.txt"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        Else
            MsgBox "Файл не выбран. Операция отменена.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Читаем содержимое файла
    fileNum = FreeFile
    On Error Resume Next
    Open selectedFile For Input As #fileNum
    If Err.Number <> 0 Then
        MsgBox "Ошибка при открытии файла: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    Line Input #fileNum, fileContent
    Close #fileNum
    
    apiKey = Trim(fileContent)
    
    If Len(apiKey) = 0 Then
        MsgBox "Файл пуст или не содержит ключа. Операция отменена.", vbCritical
        Exit Sub
    End If

    ' === ДОПОЛНИТЕЛЬНЫЕ ЗАПРОСЫ ПО ЮЛ ИЗ B52:B61 ===
    startRow = 52
    endRow = 61
    outputColH = "H"
    outputColI = "I"
    outputColJ = "J"
    
    For i = startRow To endRow
        additionalInn = Trim(targetSheet.Cells(i, "B").Value)
        
        ' Очищаем ячейки, если ИНН пустой
        If Len(additionalInn) = 0 Then
            ClearAdditionalRow targetSheet, i
            GoTo NextIteration
        End If
        
        ' Проверка длины ИНН
        If Len(additionalInn) <> 10 And Len(additionalInn) <> 12 Then
            MsgBox "Некорректный ИНН в строке " & i & " (B" & i & "). Пропущен.", vbExclamation
            GoTo NextIteration
        End If
        
        ' ======================
        ' 1. ЗАПРОС ФИНАНСОВЫХ ДАННЫХ (/v2/finances)
        ' ======================
        url = "https://api.checko.ru/v2/finances?key=" & apiKey & "&inn=" & additionalInn
        
        Set httpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        If httpReq Is Nothing Then Set httpReq = CreateObject("MSXML2.XMLHTTP")
        If httpReq Is Nothing Then
            MsgBox "Не удалось создать объект HTTP.", vbCritical
            ' Устанавливаем нули и продолжаем
            SetZerosForFinancialData targetSheet, i
        Else
            httpReq.Open "GET", url, False
            httpReq.setRequestHeader "User-Agent", "Excel Financial Data Fetcher"
            httpReq.setRequestHeader "Accept", "application/json"
            httpReq.setTimeouts 30000, 30000, 30000, 30000
            httpReq.send
            requestCount = requestCount + 1
            
            If httpReq.Status = 200 Then
                jsonText = httpReq.responseText
                
                On Error Resume Next
                Set financialData = JsonConverter.ParseJson(jsonText)
                If Err.Number <> 0 Then
                    Debug.Print "Ошибка парсинга JSON финансов для ИНН в B" & i
                    Err.Clear
                    SetZerosForFinancialData targetSheet, i
                Else
                    On Error GoTo ErrorHandler
                    
                    If Not financialData Is Nothing Then
                        If financialData.Exists("data") Then
                            Set DataObj = financialData("data")
                            
                            ' Проверяем наличие данных за 2024, но не прерываем выполнение
                            If DataObj.Exists("2024") Then
                                Set report2024 = DataObj("2024")
                                
                                ' Извлекаем нужные показатели и ДЕЛИМ НА 1000 для перевода в тысячи рублей
                                Dim revenue As Double, netProfit As Double, assets As Double
                                revenue = GetJsonValue(report2024, "2110") / 1000
                                netProfit = GetJsonValue(report2024, "2400") / 1000
                                assets = GetJsonValue(report2024, "1600") / 1000
                                
                                ' Записываем в нужные ячейки
                                With targetSheet
                                    .Cells(i, outputColH).Value = revenue   ' H - выручка
                                    .Cells(i, outputColI).Value = netProfit ' I - чистая прибыль
                                    .Cells(i, outputColJ).Value = assets    ' J - активы
                                End With
                            Else
                                ' Если данных за 2024 нет, устанавливаем нули
                                SetZerosForFinancialData targetSheet, i
                            End If
                        Else
                            SetZerosForFinancialData targetSheet, i
                        End If
                    Else
                        SetZerosForFinancialData targetSheet, i
                    End If
                End If
            Else
                Debug.Print "Ошибка финансов для ИНН в строке " & i & ": " & httpReq.Status & " - " & httpReq.statusText
                ' Устанавливаем нули при ошибке запроса
                SetZerosForFinancialData targetSheet, i
            End If
        End If
        
        ' Форматируем финансовые ячейки - ЧИСЛОВОЙ ФОРМАТ С РАЗДЕЛИТЕЛЯМИ
        With targetSheet.Range(outputColH & i & ":" & outputColJ & i)
            .NumberFormat = "# ### ###"
            .HorizontalAlignment = xlCenter
        End With

        ' ======================
        ' 2. ЗАПРОС РЕГИСТРАЦИОННЫХ ДАННЫХ (/v2/company)
        ' ======================
        url = "https://api.checko.ru/v2/company?key=" & apiKey & "&inn=" & additionalInn
        
        Set httpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        If httpReq Is Nothing Then Set httpReq = CreateObject("MSXML2.XMLHTTP")
        If httpReq Is Nothing Then
            MsgBox "Не удалось создать объект HTTP.", vbCritical
            GoTo NextIteration
        End If
        
        httpReq.Open "GET", url, False
        httpReq.setRequestHeader "User-Agent", "Excel Company Data Fetcher"
        httpReq.setRequestHeader "Accept", "application/json"
        httpReq.setTimeouts 30000, 30000, 30000, 30000
        httpReq.send
        requestCount = requestCount + 1
        
        If httpReq.Status <> 200 Then
            Debug.Print "Ошибка компании для ИНН в строке " & i & ": " & httpReq.Status & " - " & httpReq.statusText
            GoTo NextIteration
        End If
        
        jsonText = httpReq.responseText
        
        On Error Resume Next
        Set companyData = JsonConverter.ParseJson(jsonText)
        If Err.Number <> 0 Then
            Debug.Print "Ошибка парсинга JSON компании для ИНН в B" & i
            Err.Clear
            GoTo NextIteration
        End If
        On Error GoTo ErrorHandler
        
        If Not companyData.Exists("data") Then GoTo NextIteration
        Set DataObj = companyData("data")
        
        ' === 1. Название компании + вид деятельности > столбец A ===
        Dim companyInfo As String
        companyInfo = ""
        
        If DataObj.Exists("НаимСокр") Then
            companyInfo = DataObj("НаимСокр")
            companyInfo = Replace(companyInfo, "\", "")
        End If
        
        If DataObj.Exists("ОКВЭД") Then
            Dim okvedObj As Object
            Set okvedObj = DataObj("ОКВЭД")
            If okvedObj.Exists("Наим") Then
                If Len(companyInfo) > 0 Then
                    companyInfo = companyInfo & " (" & okvedObj("Наим") & ")"
                Else
                    companyInfo = "(" & okvedObj("Наим") & ")"
                End If
            End If
        End If
        
        targetSheet.Cells(i, "A").Value = companyInfo

        ' === 2. Дата регистрации > столбец C ===
        If DataObj.Exists("ДатаРег") Then
            Dim formattedDate As String
            Dim dateParts() As String
            dateParts = Split(DataObj("ДатаРег"), "-")
            If UBound(dateParts) = 2 Then
                formattedDate = dateParts(2) & "." & dateParts(1) & "." & dateParts(0)
                targetSheet.Cells(i, "C").Value = formattedDate
            End If
        End If

        ' === 3. ФИО руководителя > столбец D ===
        Dim directorFIO As String
        directorFIO = ""
        If DataObj.Exists("Руковод") Then
            Dim managers As Object
            Set managers = DataObj("Руковод")
            If TypeName(managers) = "Collection" And managers.Count > 0 Then
                Dim firstManager As Object
                Set firstManager = managers(1)
                If firstManager.Exists("ФИО") Then
                    directorFIO = firstManager("ФИО")
                End If
            End If
        End If
        ' Если руководителей нет — проверяем УпрОрг
        If Len(directorFIO) = 0 And DataObj.Exists("УпрОрг") Then
            Dim uprOrg As Object
            Set uprOrg = DataObj("УпрОрг")
            If uprOrg.Exists("НаимСокр") Then
                directorFIO = Replace(uprOrg("НаимСокр"), """", "") ' Убираем кавычки
            End If
        End If
        targetSheet.Cells(i, "D").Value = directorFIO

                 ' === 4. Учредители > столбец E ===
        Dim foundersList As String
        foundersList = ""
        Dim separator As String
        separator = ""

        If DataObj.Exists("Учред") Then
            Dim founders As Object
            Set founders = DataObj("Учред")

            ' Обрабатываем ФЛ (физические лица)
            If founders.Exists("ФЛ") Then
                Dim flList As Object
                Set flList = founders("ФЛ")
                If TypeName(flList) = "Collection" Then
                    Dim j As Long
                    For j = 1 To flList.Count
                        Dim founder As Object
                        Set founder = flList(j)
                        Dim fio As String, share As Double
                        fio = ""
                        share = 0
                        If founder.Exists("ФИО") Then fio = founder("ФИО")
                        
                        ' ИСПРАВЛЕНО: Безопасное преобразование доли для русского Excel
                        If founder.Exists("Доля") Then
                            If TypeName(founder("Доля")) = "Dictionary" Then
                                If founder("Доля").Exists("Процент") Then
                                    Dim flPercVal As Variant
                                    flPercVal = founder("Доля")("Процент")
                                    If Not IsNull(flPercVal) Then
                                        ' Пробуем прямое преобразование
                                        On Error Resume Next
                                        share = CDbl(flPercVal)
                                        If Err.Number <> 0 Then
                                            ' Если не получилось, заменяем точку на запятую
                                            Err.Clear
                                            Dim flPercStr As String
                                            flPercStr = Replace(CStr(flPercVal), ".", ",")
                                            share = CDbl(flPercStr)
                                        End If
                                        On Error GoTo ErrorHandler
                                    End If
                                End If
                            End If
                        End If
                        
                        If foundersList <> "" Then foundersList = foundersList & ", "
                        foundersList = foundersList & fio & " " & Round(share, 0) & "%"
                    Next j
                End If
            End If

            ' Обрабатываем РосОрг (российские организации)
            If founders.Exists("РосОрг") Then
                Dim rosOrgList As Object
                Set rosOrgList = founders("РосОрг")
                If TypeName(rosOrgList) = "Collection" Then
                    Dim k As Long
                    For k = 1 To rosOrgList.Count
                        Dim org As Object
                        Set org = rosOrgList(k)
                        Dim orgName As String, orgShare As Double
                        orgName = ""
                        orgShare = 0
                        If org.Exists("НаимСокр") Then orgName = Replace(org("НаимСокр"), """", "")
                        
                        ' ИСПРАВЛЕНО: Безопасное преобразование доли для русского Excel
                        If org.Exists("Доля") Then
                            Dim orgDollar As Object
                            Set orgDollar = org("Доля")
                            If TypeName(orgDollar) = "Dictionary" Then
                                If orgDollar.Exists("Процент") Then
                                    Dim rosPercVal As Variant
                                    rosPercVal = orgDollar("Процент")
                                    If Not IsNull(rosPercVal) Then
                                        ' Пробуем прямое преобразование
                                        On Error Resume Next
                                        orgShare = CDbl(rosPercVal)
                                        If Err.Number <> 0 Then
                                            ' Если не получилось, заменяем точку на запятую
                                            Err.Clear
                                            Dim rosPercStr As String
                                            rosPercStr = Replace(CStr(rosPercVal), ".", ",")
                                            orgShare = CDbl(rosPercStr)
                                        End If
                                        On Error GoTo ErrorHandler
                                    End If
                                End If
                            End If
                        End If
                        
                        If foundersList <> "" Then foundersList = foundersList & ", "
                        foundersList = foundersList & orgName & " " & Round(orgShare, 0) & "%"
                    Next k
                End If
            End If

            ' Обрабатываем ИнОрг (иностранные организации)
            If founders.Exists("ИнОрг") Then
                Dim inOrgList As Object
                Set inOrgList = founders("ИнОрг")
                If TypeName(inOrgList) = "Collection" Then
                    Dim m As Long
                    For m = 1 To inOrgList.Count
                        Dim inOrg As Object
                        Set inOrg = inOrgList(m)
                        Dim inOrgName As String, inOrgShare As Double
                        inOrgName = ""
                        inOrgShare = 0
                        If inOrg.Exists("НаимСокр") Then inOrgName = Replace(inOrg("НаимСокр"), """", "")
                        
                        ' ИСПРАВЛЕНО: Безопасное преобразование доли для русского Excel
                        If inOrg.Exists("Доля") Then
                            Dim inOrgDollar As Object
                            Set inOrgDollar = inOrg("Доля")
                            If TypeName(inOrgDollar) = "Dictionary" Then
                                If inOrgDollar.Exists("Процент") Then
                                    Dim inPercVal As Variant
                                    inPercVal = inOrgDollar("Процент")
                                    If Not IsNull(inPercVal) Then
                                        ' Пробуем прямое преобразование
                                        On Error Resume Next
                                        inOrgShare = CDbl(inPercVal)
                                        If Err.Number <> 0 Then
                                            ' Если не получилось, заменяем точку на запятую
                                            Err.Clear
                                            Dim inPercStr As String
                                            inPercStr = Replace(CStr(inPercVal), ".", ",")
                                            inOrgShare = CDbl(inPercStr)
                                        End If
                                        On Error GoTo ErrorHandler
                                    End If
                                End If
                            End If
                        End If
                        
                        If foundersList <> "" Then foundersList = foundersList & ", "
                        foundersList = foundersList & inOrgName & " " & Round(inOrgShare, 0) & "%"
                    Next m
                End If
            End If
        End If
        targetSheet.Cells(i, "E").Value = foundersList
        

        ' === 5. Населенный пункт > столбец G ===
        Dim nasPunkt As String
        nasPunkt = ""
        If DataObj.Exists("ЮрАдрес") Then
            Dim addrObj As Object
            Set addrObj = DataObj("ЮрАдрес")
            If addrObj.Exists("НасПункт") Then
                nasPunkt = addrObj("НасПункт")
            End If
        End If
        targetSheet.Cells(i, "G").Value = nasPunkt

NextIteration:
    Next i

    MsgBox "Данные успешно загружены!" & vbCrLf & _
           "Выполнено запросов: " & requestCount, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Произошла ошибка: " & Err.Description & vbCrLf & _
           "Номер ошибки: " & Err.Number & vbCrLf & _
           "Строка: " & Erl, vbCritical, "Ошибка выполнения"
End Sub

' === Вспомогательная функция для безопасного получения значения из JSON ===
Function GetJsonValue(report As Object, key As String) As Double
    On Error Resume Next
    If report.Exists(key) Then
        If IsNumeric(report(key)) Then
            GetJsonValue = CDbl(report(key))
        Else
            GetJsonValue = 0
        End If
    Else
        GetJsonValue = 0
    End If
    On Error GoTo 0
End Function

' === Установка нулей для финансовых данных ===
Sub SetZerosForFinancialData(targetSheet As Worksheet, row As Long)
    With targetSheet
        .Cells(row, "H").Value = 0   ' Выручка
        .Cells(row, "I").Value = 0   ' Чистая прибыль
        .Cells(row, "J").Value = 0   ' Активы
    End With
End Sub

' === Очистка строки дополнительной компании ===
Sub ClearAdditionalRow(targetSheet As Worksheet, row As Long)
    targetSheet.Cells(row, "A").Value = ""
    targetSheet.Cells(row, "C").Value = ""
    targetSheet.Cells(row, "D").Value = ""
    targetSheet.Cells(row, "E").Value = ""
    targetSheet.Cells(row, "G").Value = ""
    targetSheet.Cells(row, "H").Value = ""
    targetSheet.Cells(row, "I").Value = ""
    targetSheet.Cells(row, "J").Value = ""
End Sub

