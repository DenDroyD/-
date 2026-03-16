Attribute VB_Name = "ОСВсчета"
Function счета(rightCell As Range, topCell As Range) As Variant
    On Error GoTo ErrorHandler
    
    If rightCell Is Nothing Or topCell Is Nothing Then
        счета = CVErr(xlErrRef)
        Exit Function
    End If
    
    Dim cellValue As String
    Dim commaPos As Integer
    Dim numericPart As String
    
    ' Проверка на пустую ячейку
    If IsEmpty(rightCell.Value) Then
        счета = topCell.Value
        Exit Function
    End If
    
    ' Получаем значение как строку
    cellValue = CStr(rightCell.Value)
    
    ' Проверка на наличие запятой в структуре "XX.XX, Текст"
    commaPos = InStr(1, cellValue, ",")
    
    ' Извлекаем числовую часть если есть запятая
    If commaPos > 0 Then
        numericPart = Trim(Left(cellValue, commaPos - 1))
        
        ' Проверяем, является ли извлеченная часть числом
        If IsNumeric(Replace(numericPart, ".", Application.DecimalSeparator)) Then
            счета = numericPart
            Exit Function
        End If
    End If
    
    ' Стандартная проверка для чисел
    If IsNumeric(Replace(cellValue, ".", Application.DecimalSeparator)) Then
        счета = Replace(cellValue, Application.DecimalSeparator, ".")
    Else
        счета = topCell.Value
    End If
    
    Exit Function
    
ErrorHandler:
    ' Логирование ошибки или другое действие
    счета = topCell.Value
End Function

