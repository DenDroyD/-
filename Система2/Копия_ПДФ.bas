Attribute VB_Name = "Копия_ПДФ"
Sub InsertMultiplePDFPagesWithGhostscript()
    Dim pdfPath As String
    Dim pdfName As String
    Dim excelPath As String
    Dim tempFolder As String
    Dim ws As Worksheet
    Dim gsPath As String
    Dim cmd As String
    Dim shellResult As Long
    Dim fileCount As Integer
    Dim topPosition As Double
    Dim fileNames As Collection
    Dim i As Integer
    Dim img1 As Shape
    Dim img2 As Shape
    Dim leftPosition As Double
    
    ' Настройки
    Set ws = ThisWorkbook.Worksheets("Оценка")
    excelPath = ThisWorkbook.Path
    tempFolder = excelPath & "\TempPDFImages\"
    
    ' Ищем Ghostscript (обычно устанавливается с PDF24)
    gsPath = "C:\Program Files\PDF24\gs\bin\gswin64c.exe"
    If Dir(gsPath) = "" Then
        gsPath = "C:\Program Files\PDF24\gs\bin\gswin32c.exe"
        If Dir(gsPath) = "" Then
            MsgBox "Ghostscript не найден. Убедитесь, что PDF24 установлен правильно."
            Exit Sub
        End If
    End If
    
    ' Создаем временную папку
    If Dir(tempFolder, vbDirectory) = "" Then MkDir tempFolder
    
    ' Собираем все PDF файлы с словом "Оценка" в названии
    Set fileNames = New Collection
    pdfName = Dir(excelPath & "\*Оценка*.pdf")
    
    While pdfName <> ""
        fileNames.Add pdfName
        pdfName = Dir()
    Wend
    
    If fileNames.Count = 0 Then
        MsgBox "PDF файлы с словом 'Оценка' в названии не найдены"
        Exit Sub
    End If
    
    ' Очищаем лист перед вставкой
    ws.Pictures.Delete
    
    ' Начальная позиция для вставки
    topPosition = 10
    
    ' Обрабатываем каждый файл
    For i = 1 To fileNames.Count
        pdfPath = excelPath & "\" & fileNames(i)
        
        ' Конвертируем первые две страницы PDF в JPG с помощью Ghostscript
        cmd = Chr(34) & gsPath & Chr(34) & " -dNOPAUSE -dBATCH -sDEVICE=jpeg -r200 -dFirstPage=1 -dLastPage=2 -sOutputFile=" & _
              Chr(34) & tempFolder & "page-%d.jpg" & Chr(34) & " " & Chr(34) & pdfPath & Chr(34)
        
        ' Запускаем команду в скрытом режиме
        shellResult = Shell(cmd, vbHide)
        
        ' Ждем завершения конвертации
        Application.Wait Now + TimeValue("00:00:03")
        
        ' Вставляем изображения на лист слева направо
        leftPosition = 10 ' Начальная позиция по горизонтали
        
        ' Вставляем первую страницу как встроенное изображение
        If Dir(tempFolder & "page-1.jpg") <> "" Then
            Set img1 = ws.Shapes.AddPicture( _
                Filename:=tempFolder & "page-1.jpg", _
                LinkToFile:=False, _
                SaveWithDocument:=True, _
                Left:=leftPosition, _
                Top:=topPosition, _
                Width:=-1, _
                Height:=-1)
            leftPosition = leftPosition + img1.Width + 10
        End If
        
        ' Вставляем вторую страницу как встроенное изображение
        If Dir(tempFolder & "page-2.jpg") <> "" Then
            Set img2 = ws.Shapes.AddPicture( _
                Filename:=tempFolder & "page-2.jpg", _
                LinkToFile:=False, _
                SaveWithDocument:=True, _
                Left:=leftPosition, _
                Top:=topPosition, _
                Width:=-1, _
                Height:=-1)
            
            ' Определяем высоту для следующей строки
            Dim maxHeight As Double
            maxHeight = Application.Max(img1.Height, img2.Height)
            topPosition = topPosition + maxHeight + 20
        ElseIf Not img1 Is Nothing Then
            ' Если второй страницы нет, используем высоту первого изображения
            topPosition = topPosition + img1.Height + 20
        End If
        
        ' Удаляем временные файлы
        On Error Resume Next
        Kill tempFolder & "page-1.jpg"
        Kill tempFolder & "page-2.jpg"
        On Error GoTo 0
    Next i
    
    ' Удаляем временную папку
    On Error Resume Next
    RmDir tempFolder
    On Error GoTo 0
    
    MsgBox "Обработано файлов: " & fileNames.Count & ". Изображения успешно добавлены!"
End Sub

