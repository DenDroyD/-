Attribute VB_Name = "ЧекоКлин"
Sub GetOneCompanyFromChecko()
    Dim apiKey As String, inn As String
    Dim targetSheet As Worksheet
    Dim fileDialog As fileDialog
    Dim selectedFile As String
    Dim fileNum As Integer
    Dim fileContent As String

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

    ' === Основная компания (B36) ===
    inn = Trim(targetSheet.Range("B36").Value)
    If Len(inn) > 0 Then
        If Len(inn) = 10 Or Len(inn) = 12 Then
            Call ProcessMainCompany(inn, targetSheet, apiKey)
        Else
            MsgBox "Некорректный ИНН в B36! Должно быть 10 или 12 цифр.", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "ИНН не указан в ячейке B36!", vbExclamation
        Exit Sub
    End If

    ' === Пытаемся импортировать данные из Word файла ===
    Dim wordImportSuccess As Boolean
    wordImportSuccess = ImportDataFromWordForChecko(inn, False)
    
    If wordImportSuccess Then
        MsgBox "Данные успешно загружены! (включая данные из Word файла)", vbInformation
    Else
        MsgBox "Данные успешно загружены! (Word файл не найден или содержит ошибки)", vbInformation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Произошла ошибка: " & Err.Description & vbCrLf & _
           "Номер ошибки: " & Err.Number, vbCritical, "Ошибка выполнения"
End Sub

Function GetCompanyData(inn As String, apiKey As String) As Object
    Dim httpReq As Object, jsonText As String
    Dim companyData As Object
    
    ' Создание HTTP-запроса
    Set httpReq = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If httpReq Is Nothing Then Set httpReq = CreateObject("MSXML2.XMLHTTP")
    If httpReq Is Nothing Then
        Set GetCompanyData = Nothing
        Exit Function
    End If

    ' Формирование URL
    Dim url As String
    url = "https://api.checko.ru/v2/company?key=" & apiKey & "&inn=" & inn
    
    On Error Resume Next
    httpReq.Open "GET", url, False
    httpReq.setRequestHeader "User-Agent", "Excel Company Data Fetcher"
    httpReq.setRequestHeader "Accept", "application/json"
    httpReq.send
    On Error GoTo 0

    ' Проверка статуса
    If httpReq.Status <> 200 Then
        Set GetCompanyData = Nothing
        Exit Function
    End If

    jsonText = httpReq.responseText

    ' Парсим JSON
    On Error Resume Next
    Set companyData = JsonConverter.ParseJson(jsonText)
    If Err.Number <> 0 Then
        Set GetCompanyData = Nothing
        Exit Function
    End If
    
    ' Проверка данных
    If Not companyData.Exists("data") Then
        Set GetCompanyData = Nothing
        Exit Function
    End If

    Set GetCompanyData = companyData("data")
End Function

Sub ProcessMainCompany(inn As String, targetSheet As Worksheet, apiKey As String)
    Dim DataObj As Object
    Set DataObj = GetCompanyData(inn, apiKey)
    
    If DataObj Is Nothing Then
        MsgBox "Не удалось получить данные по основной компании", vbExclamation
        Exit Sub
    End If

    ' === 1. ДатаРег > B39 (формат: dd.mm.yyyy) ===
    If DataObj.Exists("ДатаРег") Then
        Dim regDate As String
        regDate = DataObj("ДатаРег")
        If Not IsNull(regDate) And Len(Trim(regDate)) > 0 Then
            Dim dateParts() As String
            dateParts = Split(regDate, "-")
            If UBound(dateParts) = 2 Then
                targetSheet.Range("B39").Value = dateParts(2) & "." & dateParts(1) & "." & dateParts(0)
            Else
                targetSheet.Range("B39").Value = ""
            End If
        Else
            targetSheet.Range("B39").Value = ""
        End If
    Else
        targetSheet.Range("B39").Value = ""
    End If

    ' === 2. НаимПолн > B34 и B35 ===
    If DataObj.Exists("НаимПолн") Then
        Dim fullName As String
        fullName = DataObj("НаимПолн")
        If IsNull(fullName) Then fullName = ""
        fullName = Trim(fullName)
        
        Dim tempArray() As String
        tempArray = Split(fullName, """")
        Dim orgType As String, shortName As String
        
        If UBound(tempArray) >= 2 Then
            orgType = Trim(tempArray(0))
            shortName = tempArray(1)
        Else
            Dim words() As String
            words = Split(fullName, " ")
            If UBound(words) >= 0 Then
                shortName = words(UBound(words))
                orgType = Trim(Replace(fullName, shortName, ""))
            Else
                orgType = fullName
                shortName = ""
            End If
        End If
        
        If Len(orgType) > 0 And Right(orgType, 1) = " " Then orgType = Left(orgType, Len(orgType) - 1)
        orgType = LCase(orgType)
        If Len(orgType) > 0 Then orgType = UCase(Left(orgType, 1)) & Mid(orgType, 2)
        If Len(shortName) > 0 Then shortName = StrConv(LCase(shortName), vbProperCase)
        
        targetSheet.Range("B34").Value = orgType
        targetSheet.Range("B35").Value = shortName
    Else
        targetSheet.Range("B34").Value = ""
        targetSheet.Range("B35").Value = ""
    End If

    ' === 3. ЮрАдрес.АдресРФ > B45 ===
    If DataObj.Exists("ЮрАдрес") Then
        Dim addrObj As Object
        Set addrObj = DataObj("ЮрАдрес")
        If addrObj.Exists("АдресРФ") Then
            Dim addrText As String
            addrText = addrObj("АдресРФ")
            If IsNull(addrText) Then addrText = ""
            targetSheet.Range("B45").Value = addrText
        Else
            targetSheet.Range("B45").Value = ""
        End If
    Else
        targetSheet.Range("B45").Value = ""
    End If

    ' === 4. ОКВЭД.Наим > B43 ===
    If DataObj.Exists("ОКВЭД") Then
        Dim okvedObj As Object
        Set okvedObj = DataObj("ОКВЭД")
        If okvedObj.Exists("Наим") Then
            Dim okvedText As String
            okvedText = okvedObj("Наим")
            If IsNull(okvedText) Then okvedText = ""
            targetSheet.Range("B43").Value = okvedText
        Else
            targetSheet.Range("B43").Value = ""
        End If
    Else
        targetSheet.Range("B43").Value = ""
    End If

    ' === 5. УстКап.Сумма > B40 (с пробелом и "руб.") ===
    If DataObj.Exists("УстКап") Then
        Dim udkapObj As Object
        Set udkapObj = DataObj("УстКап")
        If udkapObj.Exists("Сумма") Then
            Dim capSumRaw As Variant
            capSumRaw = udkapObj("Сумма")
            If Not IsNull(capSumRaw) And IsNumeric(capSumRaw) Then
                Dim capSum As Double
                capSum = CDbl(capSumRaw)
                Dim udkapSum As String
                udkapSum = Format(capSum, "#,##0") & " руб."
                udkapSum = Replace(udkapSum, ",", " ")
                targetSheet.Range("B40").Value = udkapSum
            Else
                targetSheet.Range("B40").Value = ""
            End If
        Else
            targetSheet.Range("B40").Value = ""
        End If
    Else
        targetSheet.Range("B40").Value = ""
    End If

    ' === 6. Руководитель или УпрОрг > B42 и H42 ===
    Dim directorInfo As String, directorDate As String
    directorInfo = ""
    directorDate = ""

    If DataObj.Exists("Руковод") Then
        Dim managers As Object
        Set managers = DataObj("Руковод")
        If TypeName(managers) = "Collection" And managers.Count > 0 Then
            Dim firstManager As Object
            Set firstManager = managers(1)
            If firstManager.Exists("ФИО") Then
                Dim fioVal As String
                fioVal = firstManager("ФИО")
                If Not IsNull(fioVal) Then
                    Dim должностьVal As String
                    должностьVal = ""
                    If firstManager.Exists("НаимДолжн") Then
                        должностьVal = firstManager("НаимДолжн")
                        If Not IsNull(должностьVal) Then
                            должностьVal = StrConv(LCase(должностьVal), vbProperCase)
                        Else
                            должностьVal = ""
                        End If
                    End If
                    directorInfo = должностьVal & IIf(Len(должностьVal) > 0, " - ", "") & fioVal
                End If
            End If
            ' Дата записи руководителя
            If firstManager.Exists("ДатаЗаписи") Then
                Dim dateVal As String
                dateVal = firstManager("ДатаЗаписи")
                If Not IsNull(dateVal) And Len(Trim(dateVal)) > 0 Then
                    Dim dp() As String
                    dp = Split(dateVal, "-")
                    If UBound(dp) = 2 Then
                        directorDate = dp(2) & "." & dp(1) & "." & dp(0)
                    End If
                End If
            End If
        End If
    End If

    ' Если нет Руковод — проверяем УпрОрг
    If Len(directorInfo) = 0 And DataObj.Exists("УпрОрг") Then
        Dim uprOrg As Object
        Set uprOrg = DataObj("УпрОрг")
        If uprOrg.Exists("НаимСокр") Then
            Dim uprName As String
            uprName = uprOrg("НаимСокр")
            If Not IsNull(uprName) Then
                directorInfo = Replace(uprName, """", "")
            End If
        End If
        ' Дата записи управляющей организации
        If uprOrg.Exists("ДатаЗаписи") Then
            Dim uprDateVal As String
            uprDateVal = uprOrg("ДатаЗаписи")
            If Not IsNull(uprDateVal) And Len(Trim(uprDateVal)) > 0 Then
                Dim udp() As String
                udp = Split(uprDateVal, "-")
                If UBound(udp) = 2 Then
                    directorDate = udp(2) & "." & udp(1) & "." & udp(0)
                End If
            End If
        End If
    End If

    targetSheet.Range("B42").Value = directorInfo
    targetSheet.Range("H42").Value = directorDate

    ' === 7. Учредители > B41 и H41 (с датами и залогами) ===
    Dim foundersList As String, foundersDatesList As String
    foundersList = ""
    foundersDatesList = ""
    Dim sepB As String, sepH As String
    sepB = "": sepH = ""

    If DataObj.Exists("Учред") Then
        Dim founders As Object
        Set founders = DataObj("Учред")

        ' --- ФЛ ---
        If founders.Exists("ФЛ") Then
            Dim flList As Object
            Set flList = founders("ФЛ")
            If TypeName(flList) = "Collection" Then
                Dim i As Long
                For i = 1 To flList.Count
                    Dim fl As Object
                    Set fl = flList(i)
                    ProcessFounder fl, "ФЛ", foundersList, foundersDatesList, sepB, sepH
                Next i
            End If
        End If

        ' --- РосОрг ---
        If founders.Exists("РосОрг") Then
            Dim rosList As Object
            Set rosList = founders("РосОрг")
            If TypeName(rosList) = "Collection" Then
                Dim j As Long
                For j = 1 To rosList.Count
                    Dim ros As Object
                    Set ros = rosList(j)
                    ProcessFounder ros, "РосОрг", foundersList, foundersDatesList, sepB, sepH
                Next j
            End If
        End If

        ' --- ИнОрг ---
        If founders.Exists("ИнОрг") Then
            Dim inList As Object
            Set inList = founders("ИнОрг")
            If TypeName(inList) = "Collection" Then
                Dim k As Long
                For k = 1 To inList.Count
                    Dim inOrg As Object
                    Set inOrg = inList(k)
                    ProcessFounder inOrg, "ИнОрг", foundersList, foundersDatesList, sepB, sepH
                Next k
            End If
        End If
    End If

    targetSheet.Range("B41").Value = foundersList
    targetSheet.Range("H41").Value = foundersDatesList

    ' === 8. Контакты.ВебСайт > B37 ===
    Dim webSite As String: webSite = ""
    If DataObj.Exists("Контакты") Then
        Dim contacts As Object
        Set contacts = DataObj("Контакты")
        If contacts.Exists("ВебСайт") Then
            Dim wsVal As Variant
            wsVal = contacts("ВебСайт")
            If Not IsNull(wsVal) Then
                If TypeName(wsVal) = "String" Then
                    webSite = wsVal
                ElseIf TypeName(wsVal) = "Collection" Then
                    If wsVal.Count > 0 Then webSite = wsVal(1)
                End If
            End If
        End If
    End If
    targetSheet.Range("B37").Value = webSite

    ' === 9. СЧР > B47 ===
    If DataObj.Exists("СЧР") Then
        Dim schrVal As Variant
        schrVal = DataObj("СЧР")
        If Not IsNull(schrVal) Then
            targetSheet.Range("B47").Value = schrVal
        Else
            targetSheet.Range("B47").Value = ""
        End If
    Else
        targetSheet.Range("B47").Value = ""
    End If

    ' === 10. СЧРГод > H47 ===
    If DataObj.Exists("СЧРГод") Then
        Dim schrYearVal As Variant
        schrYearVal = DataObj("СЧРГод")
        If Not IsNull(schrYearVal) Then
            targetSheet.Range("H47").Value = schrYearVal
        Else
            targetSheet.Range("H47").Value = ""
        End If
    Else
        targetSheet.Range("H47").Value = ""
    End If
    
    ' === 11. Дополнительная обработка для Чек-листа ЮЛ ===
    Call ProcessChecklistDataFromMain(DataObj)
    
End Sub

' === Вспомогательная процедура для обработки одного учредителя ===

Sub ProcessFounder(founder As Object, founderType As String, ByRef listB As String, ByRef listH As String, ByRef sepB As String, ByRef sepH As String)
    Dim name As String, share As Double, dateRec As String, pledgeInfo As String
    name = "": share = 0: dateRec = "": pledgeInfo = ""

    ' Имя/название
    If founderType = "ФЛ" Then
        If founder.Exists("ФИО") Then
            Dim fioVal As Variant
            fioVal = founder("ФИО")
            If Not IsNull(fioVal) Then name = fioVal
        End If
    Else
        If founder.Exists("НаимСокр") Then
            Dim nsVal As Variant
            nsVal = founder("НаимСокр")
            If Not IsNull(nsVal) Then name = Replace(nsVal, """", "")
        End If
    End If

    ' Доля
If founder.Exists("Доля") Then
    Dim dolya As Object
    Set dolya = founder("Доля")
    If TypeName(dolya) = "Dictionary" Then
        If dolya.Exists("Процент") Then
            Dim percVal As Variant
            percVal = dolya("Процент")
            If Not IsNull(percVal) Then
                On Error Resume Next
                share = CDbl(percVal)
                If Err.Number <> 0 Then
                    ' Если не удалось, пробуем заменить точку на запятую
                    Dim percStr As String
                    percStr = Replace(CStr(percVal), ".", ",")
                    share = CDbl(percStr)
                End If
                On Error GoTo 0
            End If
        End If
    End If
End If

    ' Дата записи
    If founder.Exists("ДатаЗаписи") Then
        Dim dzVal As Variant
        dzVal = founder("ДатаЗаписи")
        If Not IsNull(dzVal) And Len(Trim(dzVal)) > 0 Then
            Dim dp() As String
            dp = Split(dzVal, "-")
            If UBound(dp) = 2 Then
                dateRec = dp(2) & "." & dp(1) & "." & dp(0)
            End If
        End If
    End If

    ' Обременение (залог)
    If founder.Exists("Обрем") Then
        Dim obremList As Object
        Set obremList = founder("Обрем")
        If TypeName(obremList) = "Collection" Then
            Dim ob As Object
            For Each ob In obremList
                If ob.Exists("Тип") Then
                    Dim tipVal As Variant
                    tipVal = ob("Тип")
                    If Not IsNull(tipVal) And LCase(Trim(tipVal)) = "залог" Then
                        Dim zalogHolder As String, notaryDate As String
                        zalogHolder = "": notaryDate = ""

                        If ob.Exists("Залогодерж") Then
                            Dim zd As Object
                            Set zd = ob("Залогодерж")
                            If zd.Exists("НаимПолн") Then
                                Dim npVal As Variant
                                npVal = zd("НаимПолн")
                                If Not IsNull(npVal) Then
                                    ' === ИЗВЛЕКАЕМ ТОЛЬКО НАЗВАНИЕ В КАВЫЧКАХ ===
                                    Dim parts() As String
                                    parts = Split(npVal, """")
                                    If UBound(parts) >= 1 Then
                                        zalogHolder = parts(1) ' то, что между первой парой кавычек
                                    Else
                                        zalogHolder = npVal ' если кавычек нет — оставляем как есть
                                    End If
                                End If
                            End If
                        End If

                        If ob.Exists("Нотариал") Then
                            Dim notar As Object
                            Set notar = ob("Нотариал")
                            If notar.Exists("Дата") Then
                                Dim ndVal As Variant
                                ndVal = notar("Дата")
                                If Not IsNull(ndVal) And Len(Trim(ndVal)) > 0 Then
                                    Dim ndp() As String
                                    ndp = Split(ndVal, "-")
                                    If UBound(ndp) = 2 Then
                                        notaryDate = ndp(2) & "." & ndp(1) & "." & ndp(0)
                                    End If
                                End If
                            End If
                        End If

                        If Len(zalogHolder) > 0 Then
                            pledgeInfo = " (залогодержатель " & zalogHolder
                            If Len(notaryDate) > 0 Then
                                pledgeInfo = pledgeInfo & " от " & notaryDate
                            End If
                            pledgeInfo = pledgeInfo & ")"
                        End If
                        Exit For ' только первое обременение
                    End If
                End If
            Next ob
        End If
    End If

    ' === ФОРМАТИРОВАНИЕ ПРОЦЕНТА: 92.08%, а не 0% ===
    Dim shareText As String
    If share = 0 Then
        shareText = "0%"
    Else
        ' Округляем до 2 знаков и убираем лишние нули
        shareText = Format(share, "0.##") & "%"
    End If

    ' Формируем строки
    If listB <> "" Then sepB = ", "
    If listH <> "" Then sepH = ", "

    listB = listB & sepB & name & " " & shareText
    listH = listH & sepH & dateRec & pledgeInfo
End Sub
