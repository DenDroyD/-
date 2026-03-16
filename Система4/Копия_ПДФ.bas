Attribute VB_Name = "Копия_ПДФ"
' НОВЫЙ КОД - Без зависимости от Adobe Acrobat
Option Explicit

' Главная процедура для извлечения данных
Sub ExtractDataFromSPARK()
    Dim wsSystem As Worksheet, wsChecklist As Worksheet
    Dim targetINN As String, pdfPath As String
    Dim blockedInfo As String, managerInfo As String
    
    ' Установка ссылок на рабочие листы
    Set wsSystem = ThisWorkbook.Worksheets("Система4")
    Set wsChecklist = ThisWorkbook.Worksheets("Чек-лист ЮЛ")
    
    ' Получение ИНН из ячейки B36
    targetINN = Trim(wsSystem.Range("B36").Value)
    
    ' Проверка наличия ИНН
    If targetINN = "" Then
        MsgBox "ИНН не найден в ячейке B36!", vbExclamation
        Exit Sub
    End If
    
    ' Поиск PDF файла по ИНН
    pdfPath = FindPDFByINN(targetINN)
    
    If pdfPath = "" Then
        MsgBox "PDF файл для ИНН " & targetINN & " не найден в папке!" & vbCrLf & _
               "Имя файла должно содержать: СПАРК_*_" & targetINN & "_*.pdf", vbExclamation
        Exit Sub
    End If
    
    ' Метод 1: Попробовать использовать Reader через другую библиотеку
    Dim pdfText As String
    pdfText = ExtractTextWithReader(pdfPath)
    
    ' Если не получилось, попробовать метод с командной строкой
    If pdfText = "" Then
        pdfText = ExtractTextViaCommandLine(pdfPath)
    End If
    
    If pdfText = "" Then
        MsgBox "Не удалось извлечь текст из PDF файла!" & vbCrLf & _
               "Возможные решения:" & vbCrLf & _
               "1. Установите Adobe Reader" & vbCrLf & _
               "2. Установите программу для чтения PDF" & vbCrLf & _
               "3. Используйте текстовую версию PDF", vbExclamation
        Exit Sub
    End If
    
    ' Поиск информации
    blockedInfo = FindBlockedAccountsInfo(pdfText)
    managerInfo = FindManagerRiskInfo(pdfText)
    
    ' Заполнение данных
    FillChecklistData wsChecklist, blockedInfo, managerInfo
    
    MsgBox "Данные успешно извлечены и заполнены!", vbInformation
End Sub

' Метод 1: Попробовать использовать Adobe Reader через другую библиотеку
Function ExtractTextWithReader(pdfPath As String) As String
    On Error Resume Next
    
    ' Попробуем использовать Adobe Reader Type Library
    Dim readerApp As Object, readerDoc As Object
    Dim page As Object
    Dim text As String
    Dim i As Integer
    
    ' Пробуем разные способы создания объектов
    Set readerApp = CreateObject("AcroExch.App")
    If readerApp Is Nothing Then
        Set readerApp = CreateObject("Acrobat.AcroApp")
    End If
    If readerApp Is Nothing Then
        Set readerApp = CreateObject("AcroExch.AVDoc")
    End If
    
    If Not readerApp Is Nothing Then
        ' Пробуем открыть документ
        Set readerDoc = CreateObject("AcroExch.PDDoc")
        If readerDoc Is Nothing Then
            Set readerDoc = CreateObject("Acrobat.AcroPDDoc")
        End If
        
        If Not readerDoc Is Nothing Then
            If readerDoc.Open(pdfPath) Then
                ' Получаем текст с первой страницы
                Set page = readerDoc.AcquirePage(0)
                If Not page Is Nothing Then
                    text = page.GetText()
                    readerDoc.Close
                End If
            End If
        End If
    End If
    
    ExtractTextWithReader = text
End Function

' Метод 2: Использовать командную строку и утилиты
Function ExtractTextViaCommandLine(pdfPath As String) As String
    Dim cmd As String, outputFile As String
    Dim fso As Object, ts As Object
    Dim text As String
    
    ' Создаем временный файл для текста
    outputFile = ThisWorkbook.Path & "\temp_pdf_text.txt"
    
    ' Попробуем использовать pdftotext (если установлен)
    ' Это часть пакета XPDF или Poppler
    cmd = "pdftotext -layout -nopgbrk """ & pdfPath & """ """ & outputFile & """"
    
    ' Запускаем команду скрытно
    shell "cmd.exe /c " & cmd, vbHide
    
    ' Ждем немного
    Application.Wait Now + TimeValue("00:00:03")
    
    ' Читаем результат
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(outputFile) Then
        Set ts = fso.OpenTextFile(outputFile, 1)
        text = ts.ReadAll
        ts.Close
        
        ' Удаляем временный файл
        fso.DeleteFile outputFile
    End If
    
    ExtractTextViaCommandLine = text
End Function

' Функция поиска PDF файла по ИНН
Function FindPDFByINN(inn As String) As String
    Dim fso As Object, folder As Object, file As Object
    Dim filePath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(ThisWorkbook.Path)
    
    For Each file In folder.Files
        If LCase(file.name) Like "спарк*" And _
           LCase(file.name) Like "*_" & inn & "_*.pdf" Then
            FindPDFByINN = file.Path
            Exit Function
        End If
    Next file
    
    FindPDFByINN = ""
End Function

' Функция поиска информации о заблокированных счетах
Function FindBlockedAccountsInfo(pdfText As String) As String
    Dim searchText As String, result As String
    Dim startPos As Long, endPos As Long
    Dim lines() As String
    Dim i As Long
    
    ' Разделяем текст на строки
    lines = Split(pdfText, vbCrLf)
    
    ' Ищем строку с "Заблокированные счета"
    For i = 0 To UBound(lines)
        If InStr(1, lines(i), "Заблокированные счета", vbTextCompare) > 0 Then
            ' Берем всю строку или следующую
            result = Trim(lines(i))
            
            ' Если в этой строке только заголовок, берем следующую
            If Len(result) < 30 Then
                If i + 1 <= UBound(lines) Then
                    result = Trim(lines(i + 1))
                End If
            End If
            
            Exit For
        End If
    Next i
    
    FindBlockedAccountsInfo = CleanText(result)
End Function

' Функция поиска информации о руководителе/учредителе
Function FindManagerRiskInfo(pdfText As String) As String
    Dim searchText As String, result As String
    Dim startPos As Long
    Dim lines() As String
    Dim i As Long, found As Boolean
    
    ' Ищем основную фразу
    searchText = "Руководитель/учредитель компании являлся руководителем/учредителем юрлица, исключенного из ЕГРЮЛ"
    
    lines = Split(pdfText, vbCrLf)
    
    For i = 0 To UBound(lines)
        If InStr(1, lines(i), searchText, vbTextCompare) > 0 Then
            result = Trim(lines(i))
            
            ' Ищем имя руководителя в следующих строках
            Dim j As Long
            For j = i + 1 To i + 3
                If j <= UBound(lines) Then
                    Dim lineText As String
                    lineText = Trim(lines(j))
                    
                    ' Проверяем, содержит ли строка имя (предполагаем, что имя содержит дефис или запятую)
                    If Len(lineText) > 10 And _
                       (InStr(1, lineText, "-") > 0 Or _
                        InStr(1, lineText, ",") > 0 Or _
                        InStr(1, lineText, "генеральный") > 0) Then
                        result = result & vbCrLf & lineText
                        Exit For
                    End If
                End If
            Next j
            
            Exit For
        End If
    Next i
    
    FindManagerRiskInfo = CleanText(result)
End Function

' Функция очистки текста
Function CleanText(text As String) As String
    If text = "" Then Exit Function
    
    ' Удаляем лишние пробелы и переносы
    text = Replace(text, vbCrLf & vbCrLf, vbCrLf)
    text = Replace(text, "  ", " ")
    text = Trim(text)
    
    CleanText = text
End Function

' Процедура заполнения данных в чек-листе
Sub FillChecklistData(ws As Worksheet, blockedInfo As String, managerInfo As String)
    ' Заполнение данных о заблокированных счетах (ячейка D8)
    If blockedInfo <> "" Then
        ws.Range("D8").Value = blockedInfo
        
        ' Определяем статус (ячейка C8)
        If InStr(1, blockedInfo, "имеются") > 0 Or _
           InStr(1, blockedInfo, "имелись") > 0 Then
            ws.Range("C8").Value = "Да"
        Else
            ws.Range("C8").Value = "Нет"
        End If
    Else
        ws.Range("D8").Value = ""
        ws.Range("C8").Value = "Нет"
    End If
    
    ' Заполнение данных о руководителе (ячейка D7)
    If managerInfo <> "" Then
        ws.Range("D7").Value = managerInfo
        ws.Range("C7").Value = "Да"
    Else
        ws.Range("D7").Value = ""
        ws.Range("C7").Value = "Нет"
    End If
End Sub

