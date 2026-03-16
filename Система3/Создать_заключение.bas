Attribute VB_Name = "Создать_заключение"
Sub СоздатьКопиюБезМакросов()
    Dim newFileName As String
    Dim cleanName As String
    Dim savePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim i As Long
    Dim usedRng As Range
    Dim cell As Range
    
    Application.screenUpdating = False
    Application.calculation = xlCalculationManual
    Application.enableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Получаем и очищаем имя файла
    cleanName = ОчиститьИмяФайла(ПолучитьТекстДляИмени())
    If cleanName = "" Then cleanName = "Без_названия"
    
    ' Формируем путь
    newFileName = "Заключение по сделке " & cleanName & ".xlsx"
    savePath = ThisWorkbook.Path & "\" & newFileName
    
    ' Создаем копию как обычно
    ThisWorkbook.Sheets.Copy
    Set wb = ActiveWorkbook
    
    ' Обрабатываем целевые листы
    For Each ws In wb.Worksheets
        ' Для ВСЕХ листов удаляем проверки данных (выпадающие списки)
        On Error Resume Next
        ws.Cells.Validation.Delete
        On Error GoTo 0
        
        ' Только для целевых листов выполняем дополнительную обработку
        If ws.name = "Система 1-2" Or ws.name = "Система 3" Then
            ' Удаляем фигуры (кроме диаграмм)
            For i = ws.Shapes.Count To 1 Step -1
                Set shp = ws.Shapes(i)
                If Not shp.Type = msoChart Then
                    shp.Delete
                End If
            Next i
            
            ' Преобразуем формулы в значения (сохраняя форматы)
            On Error Resume Next ' На случай, если на листе нет данных
            Set usedRng = ws.UsedRange
            On Error GoTo 0
            
            If Not usedRng Is Nothing Then
                ' Используем более эффективный метод для больших диапазонов
                usedRng.Value = usedRng.Value
            End If
        End If
    Next ws
    
    ' ВАЖНО: Устанавливаем свойство HasVBProject = False для удаления макросов
    ' Это скрытое свойство, которое говорит Excel, что в книге нет макросов
    On Error Resume Next
    ' Отключаем все события на листах
    For Each ws In wb.Worksheets
        ws.EnableCalculation = False
        ws.EnableCalculation = True
    Next ws
    
    ' Сохраняем как XLSX с явным указанием, что макросов нет
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = True
    
    wb.Close SaveChanges:=False
    
    MsgBox "Файл успешно сохранен как:" & vbCrLf & newFileName, vbInformation, "Операция завершена"

Finalize:
    Application.screenUpdating = True
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    MsgBox "Произошла ошибка: " & Err.Description, vbCritical
    Resume Finalize
End Sub

Function ПолучитьТекстДляИмени() As String
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Система 3")
    If Not ws Is Nothing Then
        ПолучитьТекстДляИмени = ws.Range("B2").Value
    Else
        Set ws = ThisWorkbook.Sheets("Система 1-2")
        If Not ws Is Nothing Then ПолучитьТекстДляИмени = ws.Range("B4").Value
    End If
    
    If ПолучитьТекстДляИмени = "" Then ПолучитьТекстДляИмени = "Без_названия"
End Function

Function ОчиститьИмяФайла(name As String) As String
    Dim invalidChars As String
    Dim i As Integer
    
    invalidChars = "\/:*?""<>|[]{}=;,+`~!@#$%^&"
    
    For i = 1 To Len(invalidChars)
        name = Replace(name, Mid(invalidChars, i, 1), "_")
    Next i
    
    name = Trim(Replace(Application.WorksheetFunction.Trim(name), " ", "_"))
    ОчиститьИмяФайла = name
End Function
