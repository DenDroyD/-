Attribute VB_Name = "Module1"
Function Лизинг(текст As String, Optional компания As Range) As String
    Dim cleanText As String
    Dim списокЛК As Variant
    Dim i As Integer
    Dim currentCompany As String
    Dim foundLeasing As Boolean
    
    ' Полный список лизинговых компаний
    списокЛК = Array( _
        "ВТБ Лизинг", "ГТЛК", "СберЛизинг", "Европлан", "Росагролизинг", _
        "РЕСО-Лизинг", "Газпромбанк Автолизинг", "Газпромбанк Лизинг", "Каркаде", "Балтийский лизинг", _
        "Лизинг-Премиум", "Регион", "ТрансЛизингКом", "Скания Лизинг", "Альянс-Лизинг", _
        "Аренза-Про", "Бизнес Кар Лизинг", "ВСП-Лизинг", "БМВ Лизинг", "ВелКор", "Вольво Финанс Восток", _
        "Восток Лизинг", "ВЭБ-Лизинг", "Ураллизинг", "Альфамобиль", "ГК Регион", "ЛК АЗИЯ КОРПОРЕЙШН", _
        "Де Лаге Ланден Лизинг", "Ивеко капитал Руссия", "Икс Лизинг", "Интерлизинг", "Катерпиллар Файнэншл", _
        "КВАЗАР Лизинг", "Комацу БОТЛ Финанс СНГ", "Контрол Лизинг", "КузбассФинансЛизинг", "Лизинг-Трейд", _
        "Ликонс", "ФЛК КАМАЗ", "ЛК Эволюция", "Микро-Капитал Руссия", "Мэйджор Лизинг", "НАО Финансовые системы", _
        "Национальная Лизинговая Компания", "Нацпромлизинг", "НГМЛ Финанс", "РАФТ-ЛИЗИНГ", "ПЕАК Лизинг", _
        "Пр-Лизинг", "Проминвестлизинг", "ПСБ Лизинг", "РБ Лизинг", "Регион Лизинг", "РСХБ Лизинг", _
        "Сибирская Лизинговая Компания", "Сименс Финанс", "Система Лизинг 24", "СОБИ Лизинг", "Совкомбанк Лизинг", _
        "Солид Лизинг", "Столичный Лизинг", "СТОУН-XXI", "ТаймЛизинг", "Техно Лизинг", "Техтранслизинг", _
        "Универсальная Лизинговая Компания", "УралБизнесЛизинг", "Уралпромлизинг", "Фольксваген", "ЭкономЛизинг", _
        "Эксперт-Лизинг", "Элемент Лизинг", "Южноуральский Лизинговый Центр", "ЮниКредит Лизинг", "Петербургснаб", _
        "Роделен", "Райффайзен Лизинг", "Кредит Европа Лизинг", "ИНВЕСТ-Лизинг", "АО МСП Лизинг", "АО Талк Лизинг", _
        "Абсолют лизинг", "Балтонэксим Лизинг", "Тяжпромлизинг", "Восточный Ветер Финанс", "ЛК Аспект", "Авторелиз", _
        "Лизинговая компания КАМАЗ", "Лизинговая компания" _
    )
    
    ' Предварительная обработка текста
    cleanText = LCase(Replace(текст, Chr(160), " "))
    cleanText = Replace(cleanText, "  ", " ")
    
    ' Определяем тип платежа по приоритету
    If InStr(cleanText, "пени") > 0 Or InStr(cleanText, "пеня") > 0 Then
        Лизинг = "Пени"
        Exit Function
    ElseIf InStr(cleanText, "штраф") > 0 Then
        Лизинг = "Штраф"
        Exit Function
    ElseIf InStr(cleanText, "страхован") > 0 Or _
           InStr(cleanText, "каско") > 0 Or _
           InStr(cleanText, "осаго") > 0 Or _
           InStr(cleanText, "полис") > 0 Then
        Лизинг = "Страхование"
        Exit Function
    ElseIf InStr(cleanText, "сублизинг") > 0 Then
        Лизинг = "Сублизинг"
        Exit Function
    
    ' Уточненные условия для лизинга
    ElseIf InStr(cleanText, "лизингов") > 0 Or _
           InStr(cleanText, "лизинга") > 0 Or _
           InStr(cleanText, "лизинг") > 0 Or _
           InStr(cleanText, "договору лизинга") > 0 Or _
           InStr(cleanText, "платеж по договору лизинга") > 0 Or _
           InStr(cleanText, "платежа по договору лизинга") > 0 Then
        Лизинг = "ДЛ"
        Exit Function
    End If
    
    ' Проверяем наличие слова "аренда"
    Dim hasArenta As Boolean
    hasArenta = (InStr(cleanText, "аренд") > 0)
    
    ' Проверяем соседнюю ячейку с названием компании
    foundLeasing = False
    If Not компания Is Nothing Then
        currentCompany = компания.Value
        ' Очистка названия компании
        currentCompany = Replace(currentCompany, Chr(160), " ")
        currentCompany = Replace(currentCompany, Chr(9), " ")   ' Удаляем табуляцию
        currentCompany = Replace(currentCompany, "  ", " ")
        currentCompany = Trim(currentCompany)                   ' Удаляем пробелы по краям
        
        ' Проверка по списку лизинговых компаний
        For i = LBound(списокЛК) To UBound(списокЛК)
            If InStr(1, currentCompany, списокЛК(i), vbTextCompare) > 0 Then
                foundLeasing = True
                Exit For
            End If
        Next i
        
        ' Дополнительная проверка по ключевым словам
        If Not foundLeasing Then
            If InStr(1, currentCompany, "лизинг", vbTextCompare) > 0 Or _
               InStr(1, currentCompany, "лизинговая", vbTextCompare) > 0 Or _
               InStr(1, currentCompany, "лизинговый", vbTextCompare) > 0 Then
                foundLeasing = True
            End If
        End If
    End If
    
    ' Определяем результат
    If hasArenta Then
        If foundLeasing Then
            Лизинг = "ДЛ"
        Else
            Лизинг = "Аренда"
        End If
    Else
        If foundLeasing Then
            Лизинг = "ДЛ"
        Else
            Лизинг = ""
        End If
    End If
End Function

