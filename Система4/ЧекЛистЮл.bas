Attribute VB_Name = "ЧекЛистЮл"
' ====================================================
' МОДУЛЬ ДЛЯ ЧекЛистаЮЛ (ДЛЯ ИСПОЛЬЗОВАНИЯ В ЧЕКОКЛИНЕ)
' ====================================================

Sub ProcessChecklistDataFromMain(DataObj As Object)
    ' Эта процедура вызывается из основного макроса ЧекоКлин
    ' DataObj - это уже полученные данные компании
    
    Dim targetSheet As Worksheet
    
    ' Находим лист "Чек-лист ЮЛ"
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Sheets("Чек-лист ЮЛ")
    On Error GoTo 0
    
    If targetSheet Is Nothing Then
        MsgBox "Лист 'Чек-лист ЮЛ' не найден!", vbExclamation
        Exit Sub
    End If
    
    ' === 1. Статус ликвидации (коды 110, 101, 107, 105) ===
    ProcessLiquidationStatus DataObj, targetSheet
    
    ' === 2. Массовый руководитель ===
    ProcessMassDirector DataObj, targetSheet
    
    ' === 3. Массовый учредитель ===
    ProcessMassFounder DataObj, targetSheet
    
    ' === 4. Санкционные списки ===
    ProcessSanctions DataObj, targetSheet
    
    ' === 5. Дисквалифицированные лица ===
    ProcessDisqualifiedPerson DataObj, targetSheet
End Sub

Sub ProcessLiquidationStatus(companyData As Object, targetSheet As Worksheet)
    Dim statusData As Object
    Dim resultText As String
    Dim statusCode As String
    
    ' Проверяем наличие статуса
    If companyData.Exists("Статус") Then
        Set statusData = companyData("Статус")
        
        If statusData.Exists("Код") Then
            statusCode = Trim(statusData("Код"))
            
            ' Проверяем коды ликвидации
            If statusCode = "110" Or statusCode = "101" Or statusCode = "107" Or statusCode = "105" Then
                ' Формируем текст для ячейки D6
                resultText = "Код: """ & statusCode & """," & vbCrLf
                
                If statusData.Exists("Наим") Then
                    resultText = resultText & "  Наим: """ & Trim(statusData("Наим")) & """," & vbCrLf
                End If
                
                If statusData.Exists("ДатаЗаписи") Then
                    Dim dateStr As String
                    dateStr = Trim(statusData("ДатаЗаписи"))
                    ' Преобразуем дату в формат dd.mm.yyyy
                    If Len(dateStr) > 0 Then
                        Dim dateParts() As String
                        dateParts = Split(dateStr, "-")
                        If UBound(dateParts) = 2 Then
                            dateStr = dateParts(2) & "." & dateParts(1) & "." & dateParts(0)
                        End If
                    End If
                    resultText = resultText & "  ДатаЗаписи: """ & dateStr & """" & vbCrLf & _
                                 "ВНИМАНИЕ: Обнаружен статус ликвидации!"
                End If
                
                ' Записываем в D6
                targetSheet.Range("D6").Value = resultText
                targetSheet.Range("C6").Value = "Да"
            Else
                ' Статус не ликвидационный
                targetSheet.Range("D6").ClearContents
                targetSheet.Range("C6").Value = "Нет"
            End If
        Else
            targetSheet.Range("D6").ClearContents
            targetSheet.Range("C6").Value = "Нет"
        End If
    Else
        targetSheet.Range("D6").ClearContents
        targetSheet.Range("C6").Value = "Нет"
    End If
End Sub

Sub ProcessMassDirector(companyData As Object, targetSheet As Worksheet)
    Dim resultText As String
    Dim hasMassDirector As Boolean
    hasMassDirector = False
    
    ' Проверяем наличие руководителей
    If companyData.Exists("Руковод") Then
        Dim managers As Object
        Set managers = companyData("Руковод")
        
        If TypeName(managers) = "Collection" Then
            Dim i As Long
            For i = 1 To managers.Count
                Dim manager As Object
                Set manager = managers(i)
                
                ' Проверяем флаг МассРуковод
                If manager.Exists("МассРуковод") Then
                    If manager("МассРуковод") = True Then
                        hasMassDirector = True
                        
                        ' Формируем текст
                        resultText = ""
                        If manager.Exists("ФИО") Then
                            resultText = resultText & "ФИО: " & Trim(manager("ФИО")) & vbCrLf
                        End If
                        
                        If manager.Exists("ИНН") Then
                            resultText = resultText & "ИНН: " & Trim(manager("ИНН")) & vbCrLf
                        End If
                        
                        resultText = resultText & "МассРуковод: Да"
                        
                        Exit For
                    End If
                End If
            Next i
        End If
    End If
    
    ' Записываем результат
    If hasMassDirector Then
        targetSheet.Range("D9").Value = resultText
        targetSheet.Range("C9").Value = "Да"
    Else
        targetSheet.Range("D9").ClearContents
        targetSheet.Range("C9").Value = "Нет"
    End If
End Sub

Sub ProcessMassFounder(companyData As Object, targetSheet As Worksheet)
    Dim resultText As String
    Dim hasMassFounder As Boolean
    hasMassFounder = False
    
    ' Проверяем наличие учредителей
    If companyData.Exists("Учред") Then
        Dim founders As Object
        Set founders = companyData("Учред")
        
        ' Проверяем ФЛ (физические лица)
        If founders.Exists("ФЛ") Then
            Dim flList As Object
            Set flList = founders("ФЛ")
            
            If TypeName(flList) = "Collection" Then
                Dim i As Long
                For i = 1 To flList.Count
                    Dim founder As Object
                    Set founder = flList(i)
                    
                    ' Проверяем флаг МассУчред
                    If founder.Exists("МассУчред") Then
                        If founder("МассУчред") = True Then
                            hasMassFounder = True
                            
                            ' Формируем текст
                            resultText = ""
                            If founder.Exists("ФИО") Then
                                resultText = resultText & "ФИО: " & Trim(founder("ФИО")) & vbCrLf
                            End If
                            
                            If founder.Exists("ИНН") Then
                                resultText = resultText & "ИНН: " & Trim(founder("ИНН")) & vbCrLf
                            End If
                            
                            resultText = resultText & "МассУчред: Да"
                            
                            Exit For
                        End If
                    End If
                Next i
            End If
        End If
    End If
    
    ' Записываем результат
    If hasMassFounder Then
        targetSheet.Range("D10").Value = resultText
        targetSheet.Range("C10").Value = "Да"
    Else
        targetSheet.Range("D10").ClearContents
        targetSheet.Range("C10").Value = "Нет"
    End If
End Sub

Sub ProcessSanctions(companyData As Object, targetSheet As Worksheet)
    Dim resultText As String
    Dim hasSanctions As Boolean
    hasSanctions = False
    
    ' Проверяем флаг Санкции
    If companyData.Exists("Санкции") Then
        If companyData("Санкции") = True Then
            hasSanctions = True
            
            ' Формируем текст
            resultText = "Санкции: Да" & vbCrLf
            
            ' Добавляем страны санкций
            If companyData.Exists("СанкцииСтраны") Then
                Dim countries As Object
                Set countries = companyData("СанкцииСтраны")
                
                If TypeName(countries) = "Collection" And countries.Count > 0 Then
                    Dim countryList As String
                    Dim j As Long
                    countryList = ""
                    
                    For j = 1 To countries.Count
                        If j > 1 Then countryList = countryList & ", "
                        countryList = countryList & countries(j)
                    Next j
                    
                    resultText = resultText & "  СанкцииСтраны: " & countryList
                Else
                    resultText = resultText & "  СанкцииСтраны: "
                End If
            Else
                resultText = resultText & "  СанкцииСтраны: "
            End If
        End If
    End If
    
    ' Записываем результат
    If hasSanctions Then
        targetSheet.Range("D11").Value = resultText
        targetSheet.Range("C11").Value = "Да"
    Else
        targetSheet.Range("D11").ClearContents
        targetSheet.Range("C11").Value = "Нет"
    End If
End Sub

Sub ProcessDisqualifiedPerson(companyData As Object, targetSheet As Worksheet)
    Dim resultText As String
    Dim hasDisqualified As Boolean
    hasDisqualified = False
    
    ' Сначала проверяем руководителей
    If companyData.Exists("Руковод") Then
        Dim managers As Object
        Set managers = companyData("Руковод")
        
        If TypeName(managers) = "Collection" Then
            Dim i As Long
            For i = 1 To managers.Count
                Dim manager As Object
                Set manager = managers(i)
                
                ' Проверяем флаг ДисквЛицо
                If manager.Exists("ДисквЛицо") Then
                    If manager("ДисквЛицо") = True Then
                        hasDisqualified = True
                        
                        ' Формируем текст
                        resultText = ""
                        If manager.Exists("ФИО") Then
                            resultText = resultText & "ФИО: " & Trim(manager("ФИО")) & vbCrLf
                        End If
                        
                        If manager.Exists("ИНН") Then
                            resultText = resultText & "ИНН: " & Trim(manager("ИНН")) & vbCrLf
                        End If
                        
                        resultText = resultText & "ДисквЛицо: Да"
                        
                        Exit For
                    End If
                End If
            Next i
        End If
    End If
    
    ' Если не нашли у руководителей, проверяем учредителей
    If Not hasDisqualified And companyData.Exists("Учред") Then
        Dim founders As Object
        Set founders = companyData("Учред")
        
        ' Проверяем ФЛ (физические лица)
        If founders.Exists("ФЛ") Then
            Dim flList As Object
            Set flList = founders("ФЛ")
            
            If TypeName(flList) = "Collection" Then
                Dim j As Long
                For j = 1 To flList.Count
                    Dim founder As Object
                    Set founder = flList(j)
                    
                    ' Проверяем флаг ДисквЛицо (если есть)
                    If founder.Exists("ДисквЛицо") Then
                        If founder("ДисквЛицо") = True Then
                            hasDisqualified = True
                            
                            ' Формируем текст
                            resultText = ""
                            If founder.Exists("ФИО") Then
                                resultText = resultText & "ФИО: " & Trim(founder("ФИО")) & vbCrLf
                            End If
                            
                            If founder.Exists("ИНН") Then
                                resultText = resultText & "ИНН: " & Trim(founder("ИНН")) & vbCrLf
                            End If
                            
                            resultText = resultText & "ДисквЛицо: Да"
                            
                            Exit For
                        End If
                    End If
                Next j
            End If
        End If
    End If
    
    ' Записываем результат
    If hasDisqualified Then
        targetSheet.Range("D12").Value = resultText
        targetSheet.Range("C12").Value = "Да"
    Else
        targetSheet.Range("D12").ClearContents
        targetSheet.Range("C12").Value = "Нет"
    End If
End Sub

