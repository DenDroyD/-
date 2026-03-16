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
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Получаем и очищаем имя файла
    cleanName = ОчиститьИмяФайла(ПолучитьТекстДляИмени())
    If cleanName = "" Then cleanName = "Без_названия"
    
    ' Формируем путь
    newFileName = "Заключение по сделке " & cleanName & ".xlsx"
    savePath = ThisWorkbook.Path & "\" & newFileName
    
    ' Создаем копию
    ThisWorkbook.Sheets.Copy
    Set wb = ActiveWorkbook
    
    ' Обрабатываем ВСЕ листы
    For Each ws In wb.Worksheets
        ' Удаляем фигуры (кроме диаграмм, картинок, флажков) - только для целевых листов
        If ws.name = "Система 1-2" Or ws.name = "Система 3" Or ws.name = "Система4" Then
            For i = ws.Shapes.Count To 1 Step -1
                Set shp = ws.Shapes(i)
                
                ' Если это диаграмма — оставляем
                If shp.Type = msoChart Then
                    GoTo SkipDelete
                End If
                
                ' Если это картинка — оставляем
                If shp.Type = msoPicture Then
                    GoTo SkipDelete
                End If
                
                ' Если это элемент управления формы (флажок)
                If shp.Type = msoFormControl Then
                    If shp.FormControlType = xlCheckBox Then
                        GoTo SkipDelete
                    End If
                End If
                
                ' Проверяем, находится ли фигура в столбцах L:Q (12–17)
                If shp.TopLeftCell.Column >= 12 And shp.TopLeftCell.Column <= 17 Then
                    shp.Delete
                Else
                    GoTo SkipDelete
                End If
                
SkipDelete:
            Next i
        End If
        
        ' Преобразуем формулы в значения на ВСЕХ листах
        On Error Resume Next
        Set usedRng = ws.UsedRange
        On Error GoTo 0
        
        If Not usedRng Is Nothing Then
            usedRng.Value = usedRng.Value
        End If
        
        ' Сбрасываем переменную для следующей итерации
        Set usedRng = Nothing
    Next ws
    
     ' Скрываем указанные листы в новой книге, если они существуют
    On Error Resume Next
    wb.Sheets("Чек-лист для МКБ").Visible = xlSheetHidden
    wb.Sheets("Справочник").Visible = xlSheetHidden
    On Error GoTo 0
    
    ' Сохраняем как XLSX
    wb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    
    MsgBox "Файл успешно сохранен как:" & vbCrLf & newFileName, vbInformation, "Операция завершена"

Finalize:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
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
        GoTo CleanUp
    End If
    
    Set ws = ThisWorkbook.Sheets("Система 1-2")
    If Not ws Is Nothing Then
        ПолучитьТекстДляИмени = ws.Range("B4").Value
        GoTo CleanUp
    End If
    
    Set ws = ThisWorkbook.Sheets("Система4")
    If Not ws Is Nothing Then
        ПолучитьТекстДляИмени = ws.Range("A2").Value
        GoTo CleanUp
    End If
    
    ПолучитьТекстДляИмени = "Без_названия"

CleanUp:
    On Error GoTo 0
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
