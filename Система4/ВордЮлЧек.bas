Attribute VB_Name = "ВордЮлЧек"
Option Explicit

' ====================================================
' МОДУЛЬ ДЛЯ ИМПОРТА ИЗ WORD (ДЛЯ ИСПОЛЬЗОВАНИЯ В ЧЕКОКЛИНЕ)
' ====================================================

Function ImportDataFromWordForChecko(inn As String, Optional showMessages As Boolean = False) As Boolean
    ' Эта процедура вызывается из основного макроса ЧекоКлин
    ' inn - ИНН компании
    ' showMessages - показывать ли сообщения об ошибках
    ' Возвращает True если данные успешно импортированы, False если нет
    
    Dim wsChecklist As Worksheet
    Dim wordFileName As String, wordFilePath As String
    Dim wordApp As Object, wordDoc As Object
    Dim wordContent As String, extractedText As String
    Dim foundPos As Integer, endPos As Integer
    Dim importSuccess As Boolean
    
    ' Устанавливаем флаг успешности
    importSuccess = False
    
    ' Устанавливаем обработку ошибок
    On Error GoTo ErrorHandler
    
    ' Получаем рабочий лист
    Set wsChecklist = ThisWorkbook.Sheets("Чек-лист ЮЛ")
    
    ' Проверяем, что ИНН получен
    If inn = "" Then
        If showMessages Then
            MsgBox "ИНН не передан для поиска Word файла", vbExclamation
        End If
        ImportDataFromWordForChecko = False
        Exit Function
    End If
    
    ' Ищем Word файл с ИНН в названии
    wordFileName = ""
    Dim fso As Object, folder As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(ThisWorkbook.Path)
    
    ' Сначала ищем .docx файлы
    For Each file In folder.Files
        If LCase(file.name) Like "*спарк*" And _
           LCase(file.name) Like "*" & inn & "*" And _
           LCase(Right(file.name, 5)) = ".docx" Then
            wordFileName = file.name
            Exit For
        End If
    Next file
    
    ' Если файл не найден, ищем .doc файлы
    If wordFileName = "" Then
        For Each file In folder.Files
            If LCase(file.name) Like "*спарк*" And _
               LCase(file.name) Like "*" & inn & "*" And _
               LCase(Right(file.name, 4)) = ".doc" Then
                wordFileName = file.name
                Exit For
            End If
        Next file
    End If
    
    ' Если файл не найден
    If wordFileName = "" Then
        If showMessages Then
            MsgBox "Файл Word с ИНН " & inn & " не найден в папке", vbInformation
        End If
        ImportDataFromWordForChecko = False
        Exit Function
    End If
    
    wordFilePath = folder.Path & "\" & wordFileName
    
    ' Открываем Word документ
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False ' Скрываем Word
    Set wordDoc = wordApp.Documents.Open(wordFilePath)
    
    ' Получаем содержимое первой страницы (приблизительно)
    ' Берем первые 4000 символов - обычно это первая страница
    wordContent = Left(wordDoc.Content.text, 4000)
    
    ' 1. Поиск информации о заблокированных счетах
    wsChecklist.Range("D8").ClearContents
    wsChecklist.Range("C8").ClearContents
    
    foundPos = InStr(1, wordContent, "Заблокированные счета", vbTextCompare)
    If foundPos > 0 Then
        ' Ищем следующий перенос строки после найденного текста
        endPos = InStr(foundPos, wordContent, vbCr, vbTextCompare)
        If endPos = 0 Then endPos = InStr(foundPos, wordContent, vbLf, vbTextCompare)
        If endPos = 0 Then endPos = Len(wordContent)
        
        ' Извлекаем текст до следующего переноса строки
        extractedText = Mid(wordContent, foundPos, endPos - foundPos)
        
        ' Ищем информацию после "Заблокированные счета"
        Dim nextLinePos As Integer
        nextLinePos = InStr(endPos, wordContent, vbCr, vbTextCompare)
        If nextLinePos = 0 Then nextLinePos = InStr(endPos, wordContent, vbLf, vbTextCompare)
        
        If nextLinePos > 0 Then
            ' Ищем конец этой строки
            Dim nextLineEnd As Integer
            nextLineEnd = InStr(nextLinePos + 1, wordContent, vbCr, vbTextCompare)
            If nextLineEnd = 0 Then nextLineEnd = InStr(nextLinePos + 1, wordContent, vbLf, vbTextCompare)
            If nextLineEnd = 0 Then nextLineEnd = Len(wordContent)
            
            ' Извлекаем строку с информацией
            extractedText = Mid(wordContent, nextLinePos, nextLineEnd - nextLinePos)
            extractedText = Trim(Replace(Replace(extractedText, vbCr, ""), vbLf, ""))
            
            ' Записываем в ячейку D8
            wsChecklist.Range("D8").Value = extractedText
            
            ' Определяем, есть ли блокировка
            If InStr(1, extractedText, "имеются", vbTextCompare) > 0 Or _
               InStr(1, extractedText, "имелись", vbTextCompare) > 0 Then
                wsChecklist.Range("C8").Value = "Да"
            ElseIf InStr(1, extractedText, "нет действующих", vbTextCompare) > 0 Then
                wsChecklist.Range("C8").Value = "Нет"
            Else
                wsChecklist.Range("C8").Value = "Требует проверки"
            End If
        End If
    End If
    
    ' 2. Поиск информации о руководителе/учредителе
    wsChecklist.Range("D7").ClearContents
    wsChecklist.Range("C7").ClearContents
    
    Dim searchPhrase As String
    searchPhrase = "Руководитель/учредитель компании являлся руководителем/учредителем юрлица, исключенного из ЕГРЮЛ по инициативе ФНС или долгами перед бюджетом на момент его исключения (за последние 3 года)"
    
    foundPos = InStr(1, wordContent, searchPhrase, vbTextCompare)
    If foundPos > 0 Then
        ' Записываем основную фразу в D7
        wsChecklist.Range("D7").Value = searchPhrase
        
        ' Ищем имя руководителя после этой фразы
        ' Ищем следующий перенос строки
        endPos = InStr(foundPos, wordContent, vbCr, vbTextCompare)
        If endPos = 0 Then endPos = InStr(foundPos, wordContent, vbLf, vbTextCompare)
        If endPos = 0 Then endPos = Len(wordContent)
        
        ' Ищем следующую строку (где может быть имя)
        Dim nameStart As Integer
        nameStart = endPos + 1
        If nameStart < Len(wordContent) Then
            ' Ищем конец этой строки
            Dim nameEnd As Integer
            nameEnd = InStr(nameStart, wordContent, vbCr, vbTextCompare)
            If nameEnd = 0 Then nameEnd = InStr(nameStart, wordContent, vbLf, vbTextCompare)
            If nameEnd = 0 Then nameEnd = Len(wordContent)
            
            ' Извлекаем имя
            Dim managerName As String
            managerName = Mid(wordContent, nameStart, nameEnd - nameStart)
            managerName = Trim(Replace(Replace(managerName, vbCr, ""), vbLf, ""))
            
            ' Добавляем имя к существующей ячейке D7
            If managerName <> "" Then
                wsChecklist.Range("D7").Value = wsChecklist.Range("D7").Value & vbCrLf & managerName
            End If
        End If
        
        ' Записываем "Да" в C7
        wsChecklist.Range("C7").Value = "Да"
    Else
        ' Если фраза не найдена
        wsChecklist.Range("C7").Value = "Нет"
    End If
    
    ' Устанавливаем флаг успешного импорта
    importSuccess = True
    
CleanUp:
    ' Закрываем Word документ
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    On Error GoTo 0
    
    ' Освобождаем объекты
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set fso = Nothing
    
    ' Возвращаем результат
    ImportDataFromWordForChecko = importSuccess
    
    If importSuccess And showMessages Then
        MsgBox "Данные успешно импортированы из файла: " & wordFileName, vbInformation
    End If
    
    Exit Function
    
ErrorHandler:
    ' Ловим ошибки и продолжаем работу
    importSuccess = False
    
    ' Закрываем Word в случае ошибки
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    On Error GoTo 0
    
    ' Освобождаем объекты
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set fso = Nothing
    
    If showMessages Then
        MsgBox "Ошибка при импорте из Word: " & Err.Description, vbExclamation
    End If
    
    ImportDataFromWordForChecko = False
End Function

' Старая функция для обратной совместимости
Sub ImportDataFromWord()
    Dim wsSystem As Worksheet
    Dim targetINN As String
    
    ' Получаем рабочий лист
    Set wsSystem = ThisWorkbook.Sheets("Система4")
    
    ' Получаем ИНН из ячейки B36
    targetINN = Trim(CStr(wsSystem.Range("B36").Value))
    
    ' Вызываем новую функцию с показом сообщений
    Dim result As Boolean
    result = ImportDataFromWordForChecko(targetINN, True)
End Sub

' Функция для запуска макроса из меню
Sub RunImport()
    Call ImportDataFromWord
End Sub

