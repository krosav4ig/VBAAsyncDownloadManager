VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDownloadLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' Модуль: CDownloadLogger.cls
' Назначение: Реализация IDownloadCallback – логирование в Excel и текстовый файл
' ============================================================================
Option Explicit
Implements IDownloadCallback

Private m_logSheet As Worksheet
Private m_nextRow As Long
Private m_backupLogPath As String
Private m_taskRowMap As Object    ' Dictionary для быстрого поиска строк по TaskID

' Константы для индексов колонок (для читаемости)
Private Const COL_TASKID As Long = 1
Private Const COL_URL As Long = 2
Private Const COL_DEST As Long = 3
Private Const COL_STATUS As Long = 4
Private Const COL_PROGRESS As Long = 5
Private Const COL_START_TIME As Long = 6
Private Const COL_END_TIME As Long = 7
Private Const COL_ERROR As Long = 8

' Инициализация логгера
Public Sub Init(ByVal logSheet As Worksheet, Optional ByVal backupLogPath As String = "")
    Set m_logSheet = logSheet
    m_backupLogPath = backupLogPath
    m_nextRow = 2
    
    ' Создаём словарь для быстрого доступа к строкам (безопаснее чем Find)
    Set m_taskRowMap = CreateObject("Scripting.Dictionary")
    
    ' Заголовки таблицы
    With m_logSheet
        .Cells(1, COL_TASKID).Value = "TaskID"
        .Cells(1, COL_URL).Value = "URL"
        .Cells(1, COL_DEST).Value = "Destination"
        .Cells(1, COL_STATUS).Value = "Status"
        .Cells(1, COL_PROGRESS).Value = "Progress %"
        .Cells(1, COL_START_TIME).Value = "Start Time"
        .Cells(1, COL_END_TIME).Value = "End Time"
        .Cells(1, COL_ERROR).Value = "Error"
        .Columns("A:H").AutoFit
    End With
End Sub

' Начало загрузки
Private Sub IDownloadCallback_OnStart(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String)
    ' Добавляем новую строку в таблицу
    Dim newRow As Long
    newRow = m_nextRow
    
    Dim rowData As Variant
    rowData = Array(TaskId, url, destPath, "Downloading", 0, Now, Empty, Empty)
    
    On Error Resume Next
    m_logSheet.Cells(newRow, COL_TASKID).Resize(1, 8).Value = rowData
    On Error GoTo 0
    
    ' Сохраняем соответствие TaskID -> номер строки
    m_taskRowMap.Add CStr(TaskId), newRow
    m_nextRow = newRow + 1
    
    ' Резервное логирование в текстовый файл
    If m_backupLogPath <> "" Then
        On Error Resume Next
        Open m_backupLogPath For Append As #1
        Print #1, Now & " | Задача " & TaskId & " | НАЧАЛО | " & url & " -> " & destPath
        Close #1
        On Error GoTo 0
    End If
End Sub

' Прогресс загрузки (БЕЗОПАСНАЯ ВЕРСИЯ)
Private Sub IDownloadCallback_OnProgress(ByVal TaskId As Long, ByVal bytesReceived As Currency, _
                                          ByVal totalBytes As Currency, ByVal percent As Integer)
    On Error GoTo ErrorHandler
    
    ' Способ 1: Использование Dictionary (рекомендуется)
    ' Находим строку (автоматически с кэшированием)
    Dim rowNumber As Variant
    rowNumber = FindTaskRow(TaskId)
    
    If Not IsNull(rowNumber) And IsNumeric(rowNumber) Then
        ' Обновляем только ячейку прогресса
        m_logSheet.Cells(CLng(rowNumber), COL_PROGRESS).Value = percent
    End If
    
    ' Отладочный вывод (можно закомментировать для production)
    Debug.Print "Задача " & TaskId & ": " & percent & "% (" & bytesReceived & "/" & totalBytes & ")"
    
    Exit Sub
    
ErrorHandler:
    ' Тихая обработка ошибки - не прерываем загрузку из-за ошибки логирования
    Debug.Print "Ошибка в OnProgress для задачи " & TaskId & ": " & Err.Description
End Sub

' Завершение загрузки (БЕЗОПАСНАЯ ВЕРСИЯ)
Private Sub IDownloadCallback_OnComplete(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String, _
                                          ByVal success As Boolean, ByVal errorMsg As String)
    On Error GoTo ErrorHandler
    
    Dim status As String
    status = IIf(success, "Completed", "Error")
    
     ' Находим строку (автоматически с кэшированием)
    Dim rowNumber As Variant
    rowNumber = FindTaskRow(TaskId)
    
    If Not IsNull(rowNumber) And IsNumeric(rowNumber) Then
        Dim row As Long
        row = CLng(rowNumber)
        
        ' Обновляем существующую строку (только изменившиеся поля)
        m_logSheet.Cells(row, COL_STATUS).Value = status
        m_logSheet.Cells(row, COL_PROGRESS).Value = IIf(success, 100, m_logSheet.Cells(row, COL_PROGRESS).Value)
        m_logSheet.Cells(row, COL_END_TIME).Value = Now
        If Not success Then
            m_logSheet.Cells(row, COL_ERROR).Value = errorMsg
        End If
    Else
        ' Аварийный случай - создаём новую строку массовой записью
        Dim newRow As Long
        newRow = m_nextRow
        
        Dim rowData As Variant
        rowData = Array(TaskId, url, destPath, status, IIf(success, 100, 0), Now, Now, IIf(Not success, errorMsg, Empty))
        
        On Error Resume Next
        m_logSheet.Cells(newRow, COL_TASKID).Resize(1, 8).Value = rowData
        On Error GoTo 0
        
        m_taskRowMap.Add CStr(TaskId), newRow
        m_nextRow = newRow + 1
    End If
    
    ' Резервное логирование
    If m_backupLogPath <> "" Then
        On Error Resume Next
        Open m_backupLogPath For Append As #1
        Print #1, Now & " | Задача " & TaskId & " | " & status & " | " & url
        If Not success Then Print #1, "Ошибка: " & errorMsg
        Close #1
        On Error GoTo 0
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Ошибка в OnComplete для задачи " & TaskId & ": " & Err.Description
End Sub

' Добавьте в CDownloadLogger следующий публичный метод для принудительного обновления лога
Public Sub ForceComplete(ByVal TaskId As Long, ByVal success As Boolean, Optional ByVal errorMsg As String = "")
    Dim rowNumber As Variant
    rowNumber = FindTaskRow(TaskId)
    If Not IsNull(rowNumber) And IsNumeric(rowNumber) Then
        Dim row As Long: row = CLng(rowNumber)
        m_logSheet.Cells(row, COL_STATUS).Value = IIf(success, "Completed", "Error")
        m_logSheet.Cells(row, COL_PROGRESS).Value = IIf(success, 100, m_logSheet.Cells(row, COL_PROGRESS).Value)
        m_logSheet.Cells(row, COL_END_TIME).Value = Now
        If Not success Then m_logSheet.Cells(row, COL_ERROR).Value = errorMsg
    End If
End Sub

' Вспомогательная функция для безопасного поиска строки
Private Function FindTaskRow(ByVal TaskId As Long) As Variant
    Dim taskKey As String
    taskKey = CStr(TaskId)
    
    ' 1. Проверяем словарь
    If m_taskRowMap.Exists(taskKey) Then
        FindTaskRow = m_taskRowMap(taskKey)
        Exit Function
    End If
    
    ' 2. Если нет в словаре - ищем через Find
    On Error GoTo ErrorHandler
    
    Dim searchRange As Range
    Set searchRange = m_logSheet.Range("A:A")
    
    Dim foundCell As Range
    Set foundCell = searchRange.Find( _
        What:=taskKey, _
        After:=m_logSheet.Cells(1, 1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
    
    If Not foundCell Is Nothing Then
        ' ДОБАВЛЯЕМ В СЛОВАРЬ
        m_taskRowMap.Add taskKey, foundCell.row
        FindTaskRow = foundCell.row
        Debug.Print "Найдена и закэширована задача " & TaskId & " в строке " & foundCell.row
    Else
        FindTaskRow = Null  ' Не найдена
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Ошибка поиска задачи " & TaskId & ": " & Err.Description
    FindTaskRow = Null
End Function

' ----------------------------------------------------------------------------
' Перезагрузка словаря из существующих данных
' ----------------------------------------------------------------------------
Public Sub ReloadDictionary()
    On Error Resume Next
    
    If m_taskRowMap Is Nothing Then
        Set m_taskRowMap = CreateObject("Scripting.Dictionary")
    Else
        m_taskRowMap.RemoveAll
    End If
    
    Dim lastRow As Long
    lastRow = m_logSheet.Cells(m_logSheet.Rows.Count, 1).End(xlUp).row
    
    Dim i As Long
    For i = 2 To lastRow
        Dim taskIdValue As Variant
        taskIdValue = m_logSheet.Cells(i, 1).Value
        
        If IsNumeric(taskIdValue) Then
            Dim TaskId As Long
            TaskId = CLng(taskIdValue)
            Dim taskKey As String
            taskKey = CStr(TaskId)
            
            If Not m_taskRowMap.Exists(taskKey) Then
                m_taskRowMap.Add taskKey, i
            End If
        End If
    Next i
    
    m_nextRow = lastRow + 1
    
    Debug.Print "Загружено в кэш: " & m_taskRowMap.Count & " записей"
    On Error GoTo 0
End Sub

' Очистка ресурсов
Private Sub Class_Terminate()
    ' Очищаем словарь
    If Not m_taskRowMap Is Nothing Then
        On Error Resume Next
        m_taskRowMap.RemoveAll
        Set m_taskRowMap = Nothing
        On Error GoTo 0
    End If
    
    ' Освобождаем ссылки на объекты Excel
    Set m_logSheet = Nothing
End Sub



