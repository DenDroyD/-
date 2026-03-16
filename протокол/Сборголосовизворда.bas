Attribute VB_Name = "Сборголосовизворда"
Sub ImportApprovalFromWord()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wdApp As Object, wdDoc As Object
    Dim folderPath As String, fileName As String
    Dim fileFound As Boolean
    Dim i As Long, j As Long
    Dim wordData As Collection
    Dim nameInExcel As String, foundMatch As Boolean

    ' Настройка производительности
    Application.screenUpdating = False
    Application.calculation = xlCalculationManual
    Application.enableEvents = False

    Set wb = ThisWorkbook
    Set ws = wb.Sheets(1)
    folderPath = wb.Path & "\"

    ' Поиск файла Word
    fileFound = False
    fileName = Dir(folderPath & "*.*")
    
    Do While fileName <> ""
        If LCase(Right(fileName, 5)) = ".docx" Or LCase(Right(fileName, 4)) = ".doc" Then
            If InStr(1, fileName, "Лист", vbTextCompare) > 0 Or _
               InStr(1, fileName, "согласования", vbTextCompare) > 0 Then
                fileFound = True
                fileName = folderPath & fileName
                Exit Do
            End If
        End If
        fileName = Dir()
    Loop

    If Not fileFound Then GoTo Cleanup ' Если не найден — выходим тихо

    ' Открытие Word
    On Error Resume Next
    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Open(fileName, ReadOnly:=True)
    On Error GoTo 0

    If wdDoc Is Nothing Then GoTo Cleanup

    ' Сбор данных из всех таблиц Word
    Set wordData = New Collection
    Dim tbl As Object, row As Object
    Dim nameInWord As String, resultText As String, commentText As String

    For Each tbl In wdDoc.Tables
        For Each row In tbl.Rows
            If row.Cells.count >= 3 Then
                nameInWord = CleanText(row.Cells(1).Range.text)
                resultText = CleanText(row.Cells(2).Range.text)
                commentText = CleanText(row.Cells(3).Range.text)

                If nameInWord <> "" And nameInWord <> "Согласующий" Then
                    wordData.Add Array(nameInWord, resultText, commentText)
                End If
            End If
        Next row
    Next tbl

    ' Заполнение Excel (строки 38–45)
    For i = 38 To 45
        nameInExcel = CleanText(ws.Cells(i, 4).Value)
        If nameInExcel = "" Then GoTo SkipRow

        foundMatch = False
        For j = 1 To wordData.count
            Dim dataEntry As Variant
            dataEntry = wordData(j)
            Dim wordName As String
            wordName = dataEntry(0)

            ' Сравнение по фамилии: Excel (последнее слово), Word (первое слово)
            Dim excelLastName As String, wordLastName As String
            excelLastName = ExtractLastName(nameInExcel, False)
            wordLastName = ExtractLastName(wordName, True)

            If LCase(excelLastName) = LCase(wordLastName) Then
                ws.Cells(i, 2).Value = IIf(dataEntry(2) <> "", dataEntry(2), dataEntry(1))
                foundMatch = True
                Exit For
            End If
        Next j

SkipRow:
    Next i

Cleanup:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing: Set wdApp = Nothing

    Application.screenUpdating = True
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    
MsgBox "Голоса согласующих импортированы!", vbInformation
End Sub

' Очистка текста
Function CleanText(text As String) As String
    If text = "" Then CleanText = "": Exit Function
    text = Replace(text, Chr(13), " ")
    text = Replace(text, Chr(10), " ")
    text = Replace(text, Chr(7), "")
    text = Replace(text, ".", " ")
    text = Replace(text, ",", " ")
    text = Application.Trim(text)
    CleanText = text
End Function

' Извлечение фамилии
Function ExtractLastName(fullName As String, isFromWord As Boolean) As String
    Dim parts() As String
    parts = Split(Application.Trim(fullName), " ")
    If UBound(parts) >= 0 Then
        ExtractLastName = parts(IIf(isFromWord, 0, UBound(parts)))
    Else
        ExtractLastName = fullName
    End If
End Function


