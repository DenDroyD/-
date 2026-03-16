Attribute VB_Name = "Менее5"
Option Explicit

' =====================================================
' НАСТРОЙКИ ДЛЯ ЛИСТА "ДЕБИТОРЫ"
' =====================================================
Private Const DEBTORS_GROUP1_START As Long = 4
Private Const DEBTORS_GROUP1_END   As Long = 8
Private Const DEBTORS_AFFILIATED1  As Long = 9
Private Const DEBTORS_OTHER1       As Long = 10

Private Const DEBTORS_GROUP2_START As Long = 12
Private Const DEBTORS_GROUP2_END   As Long = 16
Private Const DEBTORS_AFFILIATED2  As Long = 17
Private Const DEBTORS_OTHER2       As Long = 18

Private Const DEBTORS_GROUP3_START As Long = 20
Private Const DEBTORS_GROUP3_END   As Long = 24
Private Const DEBTORS_AFFILIATED3  As Long = 25
Private Const DEBTORS_OTHER3       As Long = 26

Private Const DEBTORS_GROUP4_START As Long = 28
Private Const DEBTORS_GROUP4_END   As Long = 32
Private Const DEBTORS_AFFILIATED4  As Long = 33
Private Const DEBTORS_OTHER4       As Long = 34

' =====================================================
' НАСТРОЙКИ ДЛЯ ЛИСТА "КРЕДИТОРЫ"
' =====================================================
Private Const CREDITORS_GROUP1_START As Long = 4
Private Const CREDITORS_GROUP1_END   As Long = 8
Private Const CREDITORS_AFFILIATED1  As Long = 9
Private Const CREDITORS_OTHER1       As Long = 10

Private Const CREDITORS_GROUP2_START As Long = 12
Private Const CREDITORS_GROUP2_END   As Long = 16
Private Const CREDITORS_AFFILIATED2  As Long = 17
Private Const CREDITORS_OTHER2       As Long = 18

Private Const CREDITORS_GROUP3_START As Long = 20
Private Const CREDITORS_GROUP3_END   As Long = 24
Private Const CREDITORS_AFFILIATED3  As Long = 25
Private Const CREDITORS_OTHER3       As Long = 26

Private Const CREDITORS_GROUP4_START As Long = 28
Private Const CREDITORS_GROUP4_END   As Long = 32
Private Const CREDITORS_AFFILIATED4  As Long = 33
Private Const CREDITORS_OTHER4       As Long = 34

' =====================================================
' НАСТРОЙКИ ДЛЯ ЛИСТА "ОСНОВНЫЕ КА"
' =====================================================
Private Const MAIN_GROUP1_START As Long = 3
Private Const MAIN_GROUP1_END   As Long = 7
Private Const MAIN_AFFILIATED1  As Long = 8
Private Const MAIN_OTHER1       As Long = 9

Private Const MAIN_GROUP2_START As Long = 13
Private Const MAIN_GROUP2_END   As Long = 17
Private Const MAIN_AFFILIATED2  As Long = 18
Private Const MAIN_OTHER2       As Long = 19

Private Const MAIN_COL_AMOUNT   As Long = 2   ' B
Private Const MAIN_COL_PERCENT  As Long = 3   ' C

' =====================================================
' КОЛОНКИ (для Дебиторов / Кредиторов)
' =====================================================
Private Const COL_AMOUNT_2024 As Long = 2   ' B
Private Const COL_AMOUNT_2025 As Long = 4   ' D
Private Const COL_PERCENT_2025 As Long = 5  ' E

Private wsData As Worksheet
Private wsStorage As Worksheet

' =====================================================
' ПРОЦЕДУРЫ-ОБЁРТКИ ДЛЯ КАЖДОГО ЛИСТА
' =====================================================
Public Sub HideDebtors()
    ApplyHiding ThisWorkbook.Worksheets("Дебиторы"), _
                DEBTORS_GROUP1_START, DEBTORS_GROUP1_END, DEBTORS_AFFILIATED1, DEBTORS_OTHER1, _
                DEBTORS_GROUP2_START, DEBTORS_GROUP2_END, DEBTORS_AFFILIATED2, DEBTORS_OTHER2, _
                DEBTORS_GROUP3_START, DEBTORS_GROUP3_END, DEBTORS_AFFILIATED3, DEBTORS_OTHER3, _
                DEBTORS_GROUP4_START, DEBTORS_GROUP4_END, DEBTORS_AFFILIATED4, DEBTORS_OTHER4
End Sub

Public Sub HideCreditors()
    ApplyHiding ThisWorkbook.Worksheets("Кредиторы"), _
                CREDITORS_GROUP1_START, CREDITORS_GROUP1_END, CREDITORS_AFFILIATED1, CREDITORS_OTHER1, _
                CREDITORS_GROUP2_START, CREDITORS_GROUP2_END, CREDITORS_AFFILIATED2, CREDITORS_OTHER2, _
                CREDITORS_GROUP3_START, CREDITORS_GROUP3_END, CREDITORS_AFFILIATED3, CREDITORS_OTHER3, _
                CREDITORS_GROUP4_START, CREDITORS_GROUP4_END, CREDITORS_AFFILIATED4, CREDITORS_OTHER4
End Sub

Public Sub HideMainKA()
    ApplyHidingMain ThisWorkbook.Worksheets("Основные КА"), _
                    MAIN_GROUP1_START, MAIN_GROUP1_END, MAIN_AFFILIATED1, MAIN_OTHER1, _
                    MAIN_GROUP2_START, MAIN_GROUP2_END, MAIN_AFFILIATED2, MAIN_OTHER2
End Sub

' =====================================================
' УНИВЕРСАЛЬНАЯ ПРОЦЕДУРА – ПЕРЕКЛЮЧАТЕЛЬ (для Дебиторы/Кредиторы)
' =====================================================
Private Sub ApplyHiding(sht As Worksheet, _
                        ByVal g1s, ByVal g1e, ByVal aff1, ByVal other1, _
                        ByVal g2s, ByVal g2e, ByVal aff2, ByVal other2, _
                        ByVal g3s, ByVal g3e, ByVal aff3, ByVal other3, _
                        ByVal g4s, ByVal g4e, ByVal aff4, ByVal other4)
    
    Set wsData = sht
    
    ' 1. Убедимся, что лист-хранилище состояний существует
    If Not StorageExists Then CreateStorageSheet
    
    ' 2. Текущее состояние макроса
    Dim isHidden As Boolean
    isHidden = GetMacroState(wsData.Name)
    
    If Not isHidden Then
        ' --- СОСТОЯНИЕ "НЕ СКРЫТО" > выполняем скрытие ---
        PerformHide wsData.Name, _
            g1s, g1e, aff1, other1, _
            g2s, g2e, aff2, other2, _
            g3s, g3e, aff3, other3, _
            g4s, g4e, aff4, other4
    Else
        ' --- СОСТОЯНИЕ "СКРЫТО" > восстанавливаем (бэкап должен быть) ---
        If BackupExists(wsData.Name) Then
            PerformRestore wsData.Name, _
                g1s, g1e, aff1, other1, _
                g2s, g2e, aff2, other2, _
                g3s, g3e, aff3, other3, _
                g4s, g4e, aff4, other4
        Else
            ' Аварийный случай – бэкап пропал, просто скрываем заново
            PerformHide wsData.Name, _
                g1s, g1e, aff1, other1, _
                g2s, g2e, aff2, other2, _
                g3s, g3e, aff3, other3, _
                g4s, g4e, aff4, other4
        End If
    End If
End Sub

' =====================================================
' УНИВЕРСАЛЬНАЯ ПРОЦЕДУРА – ПЕРЕКЛЮЧАТЕЛЬ (для Основные КА)
' =====================================================
Private Sub ApplyHidingMain(sht As Worksheet, _
                            ByVal g1s, ByVal g1e, ByVal aff1, ByVal other1, _
                            ByVal g2s, ByVal g2e, ByVal aff2, ByVal other2)
    
    Set wsData = sht
    
    If Not StorageExists Then CreateStorageSheet
    
    Dim isHidden As Boolean
    isHidden = GetMacroState(wsData.Name)
    
    If Not isHidden Then
        PerformHideMain wsData.Name, _
            g1s, g1e, aff1, other1, _
            g2s, g2e, aff2, other2
    Else
        If BackupExists(wsData.Name) Then
            PerformRestoreMain wsData.Name, _
                g1s, g1e, aff1, other1, _
                g2s, g2e, aff2, other2
        Else
            PerformHideMain wsData.Name, _
                g1s, g1e, aff1, other1, _
                g2s, g2e, aff2, other2
        End If
    End If
End Sub

' ------------------------------------------------------------------
' ПОЛНОЕ СКРЫТИЕ (Дебиторы/Кредиторы)
' ------------------------------------------------------------------
Private Sub PerformHide(sheetName As String, _
                        ByVal g1s, ByVal g1e, ByVal aff1, ByVal other1, _
                        ByVal g2s, ByVal g2e, ByVal aff2, ByVal other2, _
                        ByVal g3s, ByVal g3e, ByVal aff3, ByVal other3, _
                        ByVal g4s, ByVal g4e, ByVal aff4, ByVal other4)
    
    ' Создаём свежий бэкап
    CreateBackupForSheet sheetName, _
        g1s, g1e, aff1, other1, _
        g2s, g2e, aff2, other2, _
        g3s, g3e, aff3, other3, _
        g4s, g4e, aff4, other4
    
    ' Применяем логику скрытия и обнуления
    ProcessGroup g1s, g1e, aff1, other1
    ProcessGroup g2s, g2e, aff2, other2
    ProcessGroup g3s, g3e, aff3, other3
    ProcessGroup g4s, g4e, aff4, other4
    
    ' Устанавливаем состояние "СКРЫТО"
    SetMacroState sheetName, True
    MsgBox "Лист '" & sheetName & "' – строки с долей <5% скрыты.", vbInformation
End Sub

' ------------------------------------------------------------------
' ПОЛНОЕ СКРЫТИЕ (Основные КА)
' ------------------------------------------------------------------
Private Sub PerformHideMain(sheetName As String, _
                            ByVal g1s, ByVal g1e, ByVal aff1, ByVal other1, _
                            ByVal g2s, ByVal g2e, ByVal aff2, ByVal other2)
    
    ' Создаём свежий бэкап (передаём две группы, остальные как 0)
    CreateBackupForSheet sheetName, _
        g1s, g1e, aff1, other1, _
        g2s, g2e, aff2, other2, _
        0, 0, 0, 0, _
        0, 0, 0, 0
    
    ' Применяем логику скрытия для Основные КА
    ProcessGroupMain g1s, g1e, aff1, other1
    ProcessGroupMain g2s, g2e, aff2, other2
    
    SetMacroState sheetName, True
    MsgBox "Лист '" & sheetName & "' – строки с долей <5% скрыты.", vbInformation
End Sub

' ------------------------------------------------------------------
' ПОЛНОЕ ВОССТАНОВЛЕНИЕ (Дебиторы/Кредиторы)
' ------------------------------------------------------------------
Private Sub PerformRestore(sheetName As String, _
                        ByVal g1s, ByVal g1e, ByVal aff1, ByVal other1, _
                        ByVal g2s, ByVal g2e, ByVal aff2, ByVal other2, _
                        ByVal g3s, ByVal g3e, ByVal aff3, ByVal other3, _
                        ByVal g4s, ByVal g4e, ByVal aff4, ByVal other4)
    
    ' Восстанавливаем данные из бэкапа
    RestoreFromBackup sheetName, _
        g1s, g1e, aff1, other1, _
        g2s, g2e, aff2, other2, _
        g3s, g3e, aff3, other3, _
        g4s, g4e, aff4, other4
    
    ' Удаляем лист бэкапа
    DeleteBackup sheetName
    
    ' Сбрасываем состояние
    SetMacroState sheetName, False
    MsgBox "Лист '" & sheetName & "' – исходное состояние восстановлено.", vbInformation
End Sub

' ------------------------------------------------------------------
' ПОЛНОЕ ВОССТАНОВЛЕНИЕ (Основные КА)
' ------------------------------------------------------------------
Private Sub PerformRestoreMain(sheetName As String, _
                                ByVal g1s, ByVal g1e, ByVal aff1, ByVal other1, _
                                ByVal g2s, ByVal g2e, ByVal aff2, ByVal other2)
    
    RestoreFromBackup sheetName, _
        g1s, g1e, aff1, other1, _
        g2s, g2e, aff2, other2, _
        0, 0, 0, 0, _
        0, 0, 0, 0
    
    DeleteBackup sheetName
    SetMacroState sheetName, False
    MsgBox "Лист '" & sheetName & "' – исходное состояние восстановлено.", vbInformation
End Sub

' ------------------------------------------------------------------
' Обработка одной группы для Дебиторы/Кредиторы
' ------------------------------------------------------------------
Private Sub ProcessGroup(ByVal startRow As Long, ByVal endRow As Long, _
                         ByVal affRow As Long, ByVal otherRow As Long)
    
    Dim i As Long
    Dim changed As Boolean
    
    Do
        changed = False
        For i = startRow To endRow
            If Not wsData.Rows(i).Hidden Then
                If IsPercentLessThan5(wsData.Cells(i, COL_PERCENT_2025)) Then
                    wsData.Rows(i).Hidden = True
                    wsData.Cells(i, COL_AMOUNT_2024).Value = 0
                    wsData.Cells(i, COL_AMOUNT_2025).Value = 0
                    changed = True
                End If
            End If
        Next i
        Application.Calculate
    Loop While changed
    
    ' Аффилированные
    If Not wsData.Rows(affRow).Hidden Then
        If IsPercentLessThan5(wsData.Cells(affRow, COL_PERCENT_2025)) Then
            wsData.Rows(affRow).Hidden = True
            wsData.Cells(affRow, COL_AMOUNT_2024).Value = 0
            wsData.Cells(affRow, COL_AMOUNT_2025).Value = 0
        End If
    End If
    
    ' Прочие (только скрываем)
    If Not wsData.Rows(otherRow).Hidden Then
        If IsPercentLessThan5(wsData.Cells(otherRow, COL_PERCENT_2025)) Then
            wsData.Rows(otherRow).Hidden = True
        End If
    End If
    
    Application.Calculate
End Sub

' ------------------------------------------------------------------
' Обработка одной группы для Основные КА
' ------------------------------------------------------------------
Private Sub ProcessGroupMain(ByVal startRow As Long, ByVal endRow As Long, _
                             ByVal affRow As Long, ByVal otherRow As Long)
    
    Dim i As Long
    Dim changed As Boolean
    
    Do
        changed = False
        For i = startRow To endRow
            If Not wsData.Rows(i).Hidden Then
                If IsPercentLessThan5(wsData.Cells(i, MAIN_COL_PERCENT)) Then
                    wsData.Rows(i).Hidden = True
                    wsData.Cells(i, MAIN_COL_AMOUNT).Value = 0
                    changed = True
                End If
            End If
        Next i
        Application.Calculate
    Loop While changed
    
    ' Аффилированные
    If Not wsData.Rows(affRow).Hidden Then
        If IsPercentLessThan5(wsData.Cells(affRow, MAIN_COL_PERCENT)) Then
            wsData.Rows(affRow).Hidden = True
            wsData.Cells(affRow, MAIN_COL_AMOUNT).Value = 0
        End If
    End If
    
    ' Прочие (только скрываем)
    If Not wsData.Rows(otherRow).Hidden Then
        If IsPercentLessThan5(wsData.Cells(otherRow, MAIN_COL_PERCENT)) Then
            wsData.Rows(otherRow).Hidden = True
        End If
    End If
    
    Application.Calculate
End Sub

' ------------------------------------------------------------------
' Проверка доли < 5%
' ------------------------------------------------------------------
Private Function IsPercentLessThan5(cell As Range) As Boolean
    Dim val As Variant
    val = cell.Value
    If IsNumeric(val) Then
        IsPercentLessThan5 = (val < 0.045)
    Else
        IsPercentLessThan5 = False
    End If
End Function

' =====================================================
' РАБОТА С РЕЗЕРВНЫМИ КОПИЯМИ
' =====================================================

' ------------------------------------------------------------------
' Проверка существования резервного листа
' ------------------------------------------------------------------
Private Function BackupExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Backup_" & sheetName)
    On Error GoTo 0
    BackupExists = Not ws Is Nothing
End Function

' ------------------------------------------------------------------
' Удаление резервного листа
' ------------------------------------------------------------------
Private Sub DeleteBackup(sheetName As String)
    Dim wsBackup As Worksheet
    On Error Resume Next
    Set wsBackup = ThisWorkbook.Worksheets("Backup_" & sheetName)
    If Not wsBackup Is Nothing Then
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        wsBackup.Visible = xlSheetVisible
        wsBackup.Delete
        
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------
' Создание резервной копии ВСЕХ строк группы
' ------------------------------------------------------------------
Private Sub CreateBackupForSheet(sheetName As String, _
                        ByVal g1s, ByVal g1e, ByVal aff1, ByVal other1, _
                        ByVal g2s, ByVal g2e, ByVal aff2, ByVal other2, _
                        ByVal g3s, ByVal g3e, ByVal aff3, ByVal other3, _
                        ByVal g4s, ByVal g4e, ByVal aff4, ByVal other4)
    
    Dim currentSheet As Worksheet
    Set currentSheet = ThisWorkbook.ActiveSheet
    
    If BackupExists(sheetName) Then DeleteBackup sheetName
    
    Dim wsBackup As Worksheet
    Set wsBackup = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsBackup.Name = "Backup_" & sheetName
    wsBackup.Visible = xlSheetVeryHidden
    
    currentSheet.Activate
    
    Dim destRow As Long
    destRow = 1
    wsBackup.Cells(destRow, 1).Value = "Sheet: " & sheetName
    destRow = destRow + 1
    
    ' Группа 1 (всегда копируем, даже если номера 0)
    If g1s > 0 And g1e > 0 Then
        CopyRangeToBackup wsBackup, destRow, g1s, g1e
        destRow = destRow + (g1e - g1s + 1)
    End If
    If aff1 > 0 Then
        CopyRangeToBackup wsBackup, destRow, aff1, aff1
        destRow = destRow + 1
    End If
    If other1 > 0 Then
        CopyRangeToBackup wsBackup, destRow, other1, other1
        destRow = destRow + 1
    End If
    
    ' Группа 2
    If g2s > 0 And g2e > 0 Then
        CopyRangeToBackup wsBackup, destRow, g2s, g2e
        destRow = destRow + (g2e - g2s + 1)
    End If
    If aff2 > 0 Then
        CopyRangeToBackup wsBackup, destRow, aff2, aff2
        destRow = destRow + 1
    End If
    If other2 > 0 Then
        CopyRangeToBackup wsBackup, destRow, other2, other2
        destRow = destRow + 1
    End If
    
    ' Группа 3
    If g3s > 0 And g3e > 0 Then
        CopyRangeToBackup wsBackup, destRow, g3s, g3e
        destRow = destRow + (g3e - g3s + 1)
    End If
    If aff3 > 0 Then
        CopyRangeToBackup wsBackup, destRow, aff3, aff3
        destRow = destRow + 1
    End If
    If other3 > 0 Then
        CopyRangeToBackup wsBackup, destRow, other3, other3
        destRow = destRow + 1
    End If
    
    ' Группа 4
    If g4s > 0 And g4e > 0 Then
        CopyRangeToBackup wsBackup, destRow, g4s, g4e
        destRow = destRow + (g4e - g4s + 1)
    End If
    If aff4 > 0 Then
        CopyRangeToBackup wsBackup, destRow, aff4, aff4
        destRow = destRow + 1
    End If
    If other4 > 0 Then
        CopyRangeToBackup wsBackup, destRow, other4, other4
    End If
End Sub

' ------------------------------------------------------------------
' Копирование диапазона строк на лист резервной копии
' ------------------------------------------------------------------
Private Sub CopyRangeToBackup(wsBackup As Worksheet, ByVal destRow As Long, _
                              ByVal rowFrom As Long, ByVal rowTo As Long)
    
    Dim srcRange As Range
    Dim destRange As Range
    
    Set srcRange = wsData.Rows(rowFrom & ":" & rowTo)
    Set destRange = wsBackup.Rows(destRow & ":" & destRow + (rowTo - rowFrom))
    
    srcRange.Copy
    destRange.PasteSpecial xlPasteAll
    Application.CutCopyMode = False
End Sub

' ------------------------------------------------------------------
' Восстановление из резервной копии
' ------------------------------------------------------------------
Private Sub RestoreFromBackup(sheetName As String, _
                        ByVal g1s, ByVal g1e, ByVal aff1, ByVal other1, _
                        ByVal g2s, ByVal g2e, ByVal aff2, ByVal other2, _
                        ByVal g3s, ByVal g3e, ByVal aff3, ByVal other3, _
                        ByVal g4s, ByVal g4e, ByVal aff4, ByVal other4)
    
    Dim wsBackup As Worksheet
    Set wsBackup = ThisWorkbook.Worksheets("Backup_" & sheetName)
    
    ' Раскрываем все строки
    On Error Resume Next
    wsData.Rows(g1s & ":" & other1).Hidden = False
    wsData.Rows(g2s & ":" & other2).Hidden = False
    wsData.Rows(g3s & ":" & other3).Hidden = False
    wsData.Rows(g4s & ":" & other4).Hidden = False
    On Error GoTo 0
    
    Dim srcRow As Long
    srcRow = 2
    
    ' Группа 1
    If g1s > 0 And g1e > 0 Then
        CopyRangeFromBackup wsBackup, srcRow, g1s, g1e
        srcRow = srcRow + (g1e - g1s + 1)
    End If
    If aff1 > 0 Then
        CopyRangeFromBackup wsBackup, srcRow, aff1, aff1
        srcRow = srcRow + 1
    End If
    If other1 > 0 Then
        CopyRangeFromBackup wsBackup, srcRow, other1, other1
        srcRow = srcRow + 1
    End If
    
    ' Группа 2
    If g2s > 0 And g2e > 0 Then
        CopyRangeFromBackup wsBackup, srcRow, g2s, g2e
        srcRow = srcRow + (g2e - g2s + 1)
    End If
    If aff2 > 0 Then
        CopyRangeFromBackup wsBackup, srcRow, aff2, aff2
        srcRow = srcRow + 1
    End If
    If other2 > 0 Then
        CopyRangeFromBackup wsBackup, srcRow, other2, other2
        srcRow = srcRow + 1
    End If
    
    ' Группа 3
    If g3s > 0 And g3e > 0 Then
        CopyRangeFromBackup wsBackup, srcRow, g3s, g3e
        srcRow = srcRow + (g3e - g3s + 1)
    End If
    If aff3 > 0 Then
        CopyRangeFromBackup wsBackup, srcRow, aff3, aff3
        srcRow = srcRow + 1
    End If
    If other3 > 0 Then
        CopyRangeFromBackup wsBackup, srcRow, other3, other3
        srcRow = srcRow + 1
    End If
    
    ' Группа 4
    If g4s > 0 And g4e > 0 Then
        CopyRangeFromBackup wsBackup, srcRow, g4s, g4e
        srcRow = srcRow + (g4e - g4s + 1)
    End If
    If aff4 > 0 Then
        CopyRangeFromBackup wsBackup, srcRow, aff4, aff4
        srcRow = srcRow + 1
    End If
    If other4 > 0 Then
        CopyRangeFromBackup wsBackup, srcRow, other4, other4
    End If
    
    Application.Calculate
End Sub

' ------------------------------------------------------------------
' Копирование диапазона с листа резервной копии на исходный лист
' ------------------------------------------------------------------
Private Sub CopyRangeFromBackup(wsBackup As Worksheet, ByVal srcRow As Long, _
                                ByVal destRowFrom As Long, ByVal destRowTo As Long)
    
    Dim srcRange As Range
    Dim destRange As Range
    
    Set srcRange = wsBackup.Rows(srcRow & ":" & srcRow + (destRowTo - destRowFrom))
    Set destRange = wsData.Rows(destRowFrom & ":" & destRowTo)
    
    srcRange.Copy
    destRange.PasteSpecial xlPasteAll
    Application.CutCopyMode = False
End Sub

' =====================================================
' ХРАНИЛИЩЕ СОСТОЯНИЙ
' =====================================================
Private Function StorageExists() As Boolean
    On Error Resume Next
    Set wsStorage = ThisWorkbook.Worksheets("Storage_DebtCred")
    On Error GoTo 0
    StorageExists = Not wsStorage Is Nothing
End Function

Private Sub CreateStorageSheet()
    Dim currentSheet As Worksheet
    Set currentSheet = ThisWorkbook.ActiveSheet
    
    Set wsStorage = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsStorage.Name = "Storage_DebtCred"
    wsStorage.Visible = xlSheetVeryHidden
    
    currentSheet.Activate
    
    With wsStorage
        .Cells(1, 1).Value = "SheetName"
        .Cells(1, 2).Value = "State"
    End With
End Sub

Private Function GetMacroState(sheetName As String) As Boolean
    Dim r As Long
    If wsStorage Is Nothing Then Set wsStorage = ThisWorkbook.Worksheets("Storage_DebtCred")
    For r = 2 To wsStorage.Cells(wsStorage.Rows.count, 1).End(xlUp).row
        If wsStorage.Cells(r, 1).Value = sheetName Then
            GetMacroState = CBool(wsStorage.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
    GetMacroState = False
End Function

Private Sub SetMacroState(sheetName As String, state As Boolean)
    Dim r As Long
    If wsStorage Is Nothing Then Set wsStorage = ThisWorkbook.Worksheets("Storage_DebtCred")
    
    For r = 2 To wsStorage.Cells(wsStorage.Rows.count, 1).End(xlUp).row
        If wsStorage.Cells(r, 1).Value = sheetName Then
            wsStorage.Cells(r, 2).Value = state
            Exit Sub
        End If
    Next r
    
    Dim i As Long
    i = wsStorage.Cells(wsStorage.Rows.count, 1).End(xlUp).row + 1
    wsStorage.Cells(i, 1).Value = sheetName
    wsStorage.Cells(i, 2).Value = state
End Sub

' =====================================================
' СОЗДАНИЕ КНОПОК
' =====================================================
Public Sub AddButtonsToSheets()
    Dim btn As Button
    Dim sht As Worksheet
    Dim rng As Range
    
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name = "Дебиторы" Or sht.Name = "Кредиторы" Then
            On Error Resume Next
            sht.Buttons.Delete
            On Error GoTo 0
            
            Set rng = sht.Range("A1")
            Set btn = sht.Buttons.Add(Left:=rng.Left, Top:=rng.Top, Width:=rng.Width, Height:=rng.Height)
            With btn
                .Caption = "Скрыть / восстановить (<5%)"
                .OnAction = IIf(sht.Name = "Дебиторы", "HideDebtors", "HideCreditors")
                .Font.Size = 10
                .Font.Bold = True
                .Placement = xlMoveAndSize
            End With
        ElseIf sht.Name = "Основные КА" Then
            On Error Resume Next
            sht.Buttons.Delete
            On Error GoTo 0
            
            Set rng = sht.Range("A1")   ' Кнопка в A1
            Set btn = sht.Buttons.Add(Left:=rng.Left, Top:=rng.Top, Width:=rng.Width, Height:=rng.Height)
            With btn
                .Caption = "Скрыть / восстановить (<5%)"
                .OnAction = "HideMainKA"
                .Font.Size = 10
                .Font.Bold = True
                .Placement = xlMoveAndSize
            End With
        End If
    Next sht
End Sub

