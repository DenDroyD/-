Attribute VB_Name = "Module21"
Function номер(текст As String, Optional проверка As Variant) As String
    On Error GoTo ErrorHandler ' Включаем обработку ошибок
    
    ' === ПРОВЕРКА ВТОРОГО АРГУМЕНТА ===
    If IsMissing(проверка) Or IsEmpty(проверка) Then
        номер = ""
        Exit Function
    End If
    
    Dim triggerText As String
    triggerText = Trim(CStr(проверка))
    
    If triggerText = "" Then
        номер = ""
        Exit Function
    End If
    
    ' Проверяем наличие ключевых слов
    If Not RegExTest(triggerText, "ДЛ|Страхование|Сублизинг|Пени|Штраф|Аренда|Лизинг") Then
        номер = ""
        Exit Function
    End If
    
    ' === ПОДГОТОВКА ТЕКСТА ===
    Dim cleanText As String
    cleanText = ПодготовитьТекст(текст)
    
    ' === ПОИСК НОМЕРА ДОГОВОРА ===
    номер = НайтиНомерДоговора(cleanText)
    
    ' === ДОПОЛНИТЕЛЬНАЯ ПРОВЕРКА ===
    If номер <> "" Then
        ' Удаляем возможные остатки "o " от "No"
        If Left(номер, 2) = "o " Then
            номер = Mid(номер, 3)
        End If
        
        ' Проверка на наличие цифр (минимум 2 цифры)
        If Not RegExTest(номер, "[0-9]{2,}") Then
            номер = ""
        End If
    End If
    
    Exit Function

ErrorHandler:
    номер = "Ошибка: " & Err.Description
End Function

Private Function ПодготовитьТекст(текст As String) As String
    Dim result As String
    result = Replace(текст, Chr(160), " ")  ' Неразрывные пробелы
    result = Replace(result, "  ", " ")     ' Двойные пробелы
    
    ' === ОСОБАЯ ОБРАБОТКА ПРЕФИКСОВ ===
    ' Замена всех вариантов "No" на "№" с удалением "o"
    result = Replace(result, "No ", "№ ")
    result = Replace(result, "No", "№")
    result = Replace(result, "Nо ", "№ ")
    result = Replace(result, "Nо", "№")
    result = Replace(result, "NО ", "№ ")
    result = Replace(result, "NО", "№")
    
    ' Стандартная подготовка
    result = Replace(result, "№", " №")     ' Добавляем пробел перед №
    result = Replace(result, "дог.", "договор ") ' Расширяем сокращения
    result = Replace(result, "N", " №")     ' Заменяем N на №
    result = Replace(result, "n", " №")     ' Заменяем n на №
    result = Replace(result, "ДФЛ", " ДЛ ") ' Унифицируем ДФЛ
    result = Replace(result, "договора", "договора ") ' Для корректного поиска
    result = Replace(result, "договору", "договору ") ' Для корректного поиска
    result = Replace(result, "(", " (")    ' Добавляем пробел перед скобкой
    result = Replace(result, ")", ") ")    ' Добавляем пробел после скобки
    
    ' Удаляем повторяющиеся пробелы
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    result = Trim(result)
    
    ПодготовитьТекст = result
End Function

Private Function НайтиНомерДоговора(текст As String) As String
    Static regex As Object
    If regex Is Nothing Then Set regex = CreateObject("VBScript.RegExp")
    
    ' === КЛЮЧЕВЫЕ СЛОВА ДЛЯ ПОИСКА ===
    Dim ключевыеСлова() As String
    ключевыеСлова = Split("договор,договора,договору,договоре,лизинга,сублизинг,сублизинга,сублизингу,сублизинге,лизинговый,лизинговые,лизингового,лизинговому,лизинговым,лизинговом,страховани,страхованию,страховании,страхования,комисси,комиссии,комиссий,комиссиям,комиссиями,пени,штраф,штрафа,штрафу,штрафом,штрафе,дл,дфл,договор финансовой аренды,финансовой аренды,аренды", ",")
    
    ' === ПОИСК ПО ВСЕМ КЛЮЧЕВЫМ СЛОВАМ ===
    Dim i As Long
    For i = LBound(ключевыеСлова) To UBound(ключевыеСлова)
        Dim pattern As String
        Dim keyword As String
        keyword = Trim(ключевыеСлова(i))
        
        ' === ФОРМИРОВАНИЕ ПАТТЕРНА ДЛЯ КОНКРЕТНОГО КЛЮЧЕВОГО СЛОВА ===
        ' Увеличен лимит до {0,25} для захвата более длинных префиксов
        pattern = "(" & keyword & ")(?:[а-яё]*)?\s*[^№\d]{0,25}\s*[№]?\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,})"
        
        ' Добавляем ограничения для конца номера
        pattern = pattern & "(?=\s*(?:от\s|за\s|г\.|согласно|,|\.|$| \d{2}|;|\(|\)|\d{2}\.\d{2}\.\d{4}|в\s*т\.ч\.|НДС|В\s+том\s+числе|В\s*т\.ч|В\s*т\s*ч))"
        
        regex.pattern = pattern
        regex.IgnoreCase = True
        regex.Global = False
        
        If regex.Test(текст) Then
            Dim matches As Object
            Set matches = regex.Execute(текст)
            
            If matches.Count > 0 Then
                Dim candidate As String
                candidate = matches(0).SubMatches(1) ' Второй подшаблон - сам номер
                
                ' Очистка и проверка номера
                candidate = ОчиститьНомер(candidate, текст)
                
                ' Проверка, что это не страховой полис
                If Not (RegExTest(текст, "страховани[ея]|полис|осаго|каско") And _
                        Not RegExTest(текст, "по\s+договору\s+(?:лизинга|финансовой\s*аренды|сублизинга)")) Then
                    ' Проверка на минимальное количество цифр
                    If RegExTest(candidate, "[0-9]{2,}") Then
                        НайтиНомерДоговора = candidate
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
    
    ' === ДОПОЛНИТЕЛЬНЫЙ ПОИСК ДЛЯ СЛОЖНЫХ СЛУЧАЕВ ===
    НайтиНомерДоговора = НайтиСложныйНомер(текст)
End Function

Private Function ОчиститьНомер(номер As String, текст As String) As String
    ' Удаление лишних пробелов
    номер = Trim(номер)
    Do While InStr(номер, "  ") > 0
        номер = Replace(номер, "  ", " ")
    Loop
    
    ' Удаление точек и запятых в конце
    If Right(номер, 1) = "." Or Right(номер, 1) = "," Then
        номер = Left(номер, Len(номер) - 1)
    End If
    
    ' === ОСОБАЯ ОБРАБОТКА СКОБОК ===
    ' Если есть открывающая скобка, обрезаем до нее
    Dim pos As Long
    pos = InStr(номер, "(")
    If pos > 0 Then
        номер = Left(номер, pos - 1)
    End If
    
    ' === СТОП-СЛОВА ДЛЯ ОБРЕЗКИ ===
    Dim stopWords() As String
    stopWords = Split("от,за,г.,согласно,вх.д.,по,вх.д,на сумму,размере,руб,платеж,оплата,договор,страхов,полис,каско,осаго,лизинга,сублизинг,договору,договора,№,в т.ч.,в т ч,в.т.ч,НДС,НДС20%,В том числе,В т ч,В.т.ч,В т.ч,Вт.ч,В тч,Вт ч,в т.ч,в т ч,в.т.ч,в тч,вт ч,вт.ч", ",")
    
    Dim minPos As Long
    minPos = Len(номер) + 1
    
    Dim i As Long
    For i = LBound(stopWords) To UBound(stopWords)
        Dim stopWord As String
        stopWord = Trim(stopWords(i))
        
        ' Поиск стоп-слова в номере
        pos = InStr(1, номер, " " & stopWord, vbTextCompare)
        If pos > 0 And pos < minPos Then minPos = pos
        
        pos = InStr(1, номер, stopWord & " ", vbTextCompare)
        If pos > 0 And pos < minPos Then minPos = pos
        
        pos = InStr(1, номер, stopWord, vbTextCompare)
        If pos > 0 And pos < minPos Then minPos = pos
    Next i
    
    If minPos <= Len(номер) Then
        номер = Left(номер, minPos - 1)
    End If
    
    ' Удаление отдельных цифр в конце (номера платежей)
    If RegExTest(номер, "\d$") Then
        Dim parts() As String
        parts = Split(номер, " ")
        If UBound(parts) > 0 Then
            If IsNumeric(parts(UBound(parts))) Then
                номер = Trim(Left(номер, Len(номер) - Len(parts(UBound(parts))) - 1))
            End If
        End If
    End If
    
    ' Удаление лишних пробелов вокруг дефисов и слэшей
    номер = Replace(номер, " -", "-")
    номер = Replace(номер, "- ", "-")
    номер = Replace(номер, " /", "/")
    номер = Replace(номер, "/ ", "/")
    
    ' === КРИТИЧЕСКАЯ ИЗМЕНЕННАЯ ОБРАБОТКА ПРОБЕЛОВ ===
    ' Сохраняем пробелы только в специфических форматах, но не удаляем буквы из начала
    If RegExTest(номер, "^[A-ZА-Я]{1,3}\s+\d") Then
        ' Сохраняем пробелы для форматов вроде "АЛ 231178/04-23"
        Do While InStr(номер, "  ") > 0
            номер = Replace(номер, "  ", " ")
        Loop
    Else
        ' Удаляем ВСЕ пробелы, кроме случаев с буквами в начале
        номер = Replace(номер, " ", "")
    End If
    
    номер = Trim(номер)
    
    ОчиститьНомер = номер
End Function

Private Function НайтиСложныйНомер(текст As String) As String
    Static regex As Object
    If regex Is Nothing Then Set regex = CreateObject("VBScript.RegExp")
    
    ' === ДОПОЛНИТЕЛЬНЫЕ ПАТТЕРНЫ ДЛЯ СЛОЖНЫХ СЛУЧАЕВ ===
    Dim patterns() As String
    patterns = Split( _
        "по договору\s+([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "договор[а-яё]*[^№\d]{0,25}[№]?\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "лизинга[^№\d]{0,25}[№]?\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "сублизинг[а-яё]*[^№\d]{0,25}[№]?\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "по\s+договор[уа][^№\d]{0,25}[№]?\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "дл[^№\d]{0,25}[№]?\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "дог\.?[^№\d]{0,25}[№]?\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "страховани[ея][^№\d]{0,25}[№]?\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "комисси[яи][^№\d]{0,25}[№]?\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "пени[^№\d]{0,25}[№]?\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "штраф[аы]?[^№\d]{0,25}[№]?\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "финансовой\s*аренды[^№\d]{0,25}[№]?\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "(?:ДЛ|ДФЛ)[\-\s]*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "[ЛL]\s*[-–]\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,});" & _
        "[№]\s*([a-zA-Zа-яА-Я0-9][a-zA-Zа-яА-Я0-9\s\/\-–\.]{2,})", ";")
    
    Dim i As Long
    For i = LBound(patterns) To UBound(patterns)
        regex.pattern = patterns(i)
        regex.IgnoreCase = True
        regex.Global = False
        
        If regex.Test(текст) Then
            Dim matches As Object
            Set matches = regex.Execute(текст)
            Dim candidate As String
            candidate = matches(0).SubMatches(0)
            
            ' Очистка и проверка номера
            candidate = ОчиститьНомер(candidate, текст)
            
            ' Добавляем префиксы для специальных форматов
            If i = 11 And Not Left(candidate, 3) = "ДЛ-" Then ' ДЛ-XXXX
                candidate = "ДЛ-" & candidate
            ElseIf i = 12 And Not Left(candidate, 2) = "Л-" Then ' Л-XXXX
                candidate = "Л-" & candidate
            End If
            
            ' Проверка на минимальное количество цифр
            If RegExTest(candidate, "[0-9]{2,}") Then
                ' Проверка, что это не страховой полис
                If Not (RegExTest(текст, "страховани[ея]|полис|осаго|каско") And _
                        Not RegExTest(текст, "по\s+договору\s+(?:лизинга|финансовой\s*аренды|сублизинга)")) Then
                    НайтиСложныйНомер = candidate
                    Exit Function
                End If
            End If
        End If
    Next i
    
    НайтиСложныйНомер = ""
End Function

Private Function RegExTest(text As String, pattern As String) As Boolean
    Static regex As Object
    If regex Is Nothing Then Set regex = CreateObject("VBScript.RegExp")
    
    regex.pattern = pattern
    regex.IgnoreCase = True
    RegExTest = regex.Test(text)
End Function

