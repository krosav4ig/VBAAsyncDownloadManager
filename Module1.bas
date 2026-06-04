Attribute VB_Name = "Module1"
' ============================================================================
' Модуль: Module1
' Назначение: Точка входа, глобальные обработчики
' ============================================================================
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private g_manager As CAsyncDownloadManager
Private g_logger As CDownloadLogger
Private g_criticalErrorLogPath As String

' ============================================================================
' Глобальный таймер (вызывается менеджером)
' ============================================================================
Public Sub CAsyncDownloadManager_TimerCheck()
    On Error GoTo ErrorHandler
    
    If Not g_manager Is Nothing Then
        g_manager.TimerCheck
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Критическая ошибка - попытка логирования
    On Error Resume Next
    Dim logFile As Integer
    logFile = FreeFile
    Open g_criticalErrorLogPath For Append As #logFile
    Print #logFile, Now & " [CRITICAL] TimerCheck error: " & Err.Description
    Close #logFile
    On Error GoTo 0
End Sub

' ============================================================================
' Основная процедура запуска
' ============================================================================
Public Sub StartBatchDownload(Optional ByVal logLevel As Long = LL_INFO)
    On Error GoTo CriticalError
    
    ' Инициализация критического лога
    g_criticalErrorLogPath = ThisWorkbook.Path & "\critical_errors.log"
    
    ' Подготовка лог-листа
    Dim wsLog As Worksheet
    Set wsLog = GetOrCreateLogSheet()
    
    ' Инициализация логгера
    Set g_logger = New CDownloadLogger
    g_logger.Init wsLog, ThisWorkbook.Path & "\download_log.txt", logLevel
    g_logger.ReloadDictionary
    
    g_logger.LogMessage LL_INFO, 0, "=== НАЧАЛО СЕССИИ ЗАГРУЗКИ ==="
    
    ' Инициализация менеджера
    Set g_manager = New CAsyncDownloadManager
    g_manager.Init maxConcurrent:=5, callback:=g_logger, _
                   resolveTimeoutMs:=30000, connectTimeoutMs:=60000, _
                   sendTimeoutMs:=30000, receiveTimeoutMs:=60000, _
                   bufferSizeBytes:=131072, checkIntervalSec:=3, _
                   maxRetries:=3
    
    ' Подключение обработчиков событий
    HookManagerEvents
    
    ' Чтение задач
    Dim taskCount As Long
    taskCount = LoadTasksFromTable()
    
    If taskCount = 0 Then
        g_logger.LogMessage LL_WARNING, 0, "Нет задач для загрузки"
        MsgBox "Нет задач для загрузки. Проверьте таблицу 'Ссылки'.", vbExclamation
        Cleanup
        Exit Sub
    End If
    
    g_logger.LogMessage LL_INFO, 0, "Загружено задач: " & taskCount
    
    ' Запуск
    g_manager.Start
    
    ' Цикл ожидания с обработкой событий
    Dim lastStatusTime As Date
    lastStatusTime = Now
    
    Do While g_manager.IsBusy
        DoEvents
        Sleep 50
        
        ' Периодический вывод статуса (каждые 5 секунд)
        If DateDiff("s", lastStatusTime, Now) >= 5 Then
            lastStatusTime = Now
            Debug.Print Format(Now, "HH:MM:SS") & " - Активно: " & g_manager.ActiveCount & _
                          ", Ожидает: " & g_manager.PendingCount & _
                          ", Завершено: " & g_manager.CompletedCount & _
                          ", Ошибок: " & g_manager.FailedCount
        End If
    Loop
    
    ' Итоговое сообщение
    Dim resultMsg As String
    resultMsg = "Все загрузки завершены!" & vbNewLine & _
                "Успешно: " & g_manager.CompletedCount & vbNewLine & _
                "Ошибок: " & g_manager.FailedCount
    
    If g_manager.FailedCount > 0 Then
        resultMsg = resultMsg & vbNewLine & vbNewLine & "Проверьте лог для деталей."
        MsgBox resultMsg, vbExclamation
    Else
        MsgBox resultMsg, vbInformation
    End If
    
    g_logger.LogMessage LL_INFO, 0, "=== СЕССИЯ ЗАВЕРШЕНА ==="
    
    Cleanup
    Exit Sub
    
CriticalError:
    Dim critMsg As String
    critMsg = "КРИТИЧЕСКАЯ ОШИБКА: " & Err.Description & vbNewLine & _
              "Код: " & Err.Number & vbNewLine & _
              "Место: " & Erl & vbNewLine & _
              "Пожалуйста, перезапустите приложение."
    
    MsgBox critMsg, vbCritical, "Критическая ошибка"
    
    On Error Resume Next
    Dim critFile As Integer
    critFile = FreeFile
    Open g_criticalErrorLogPath For Append As #critFile
    Print #critFile, Now & " " & critMsg
    Close #critFile
    On Error GoTo 0
    
    Cleanup
End Sub

' ============================================================================
' Получение или создание лог-листа
' ============================================================================
Private Function GetOrCreateLogSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateLogSheet = ThisWorkbook.Sheets("DownloadLog")
    If GetOrCreateLogSheet Is Nothing Then
        Set GetOrCreateLogSheet = ThisWorkbook.Sheets.Add
        GetOrCreateLogSheet.Name = "DownloadLog"
    End If
    On Error GoTo 0
End Function

' ============================================================================
' Загрузка задач из таблицы
' ============================================================================
Private Function LoadTasksFromTable() As Long
    Dim tbl As ListObject
    Dim urlCol As Long, pathCol As Long
    Dim dataRange As Range
    Dim i As Long
    Dim count As Long
    
    On Error Resume Next
    Set tbl = ThisWorkbook.Sheets("Лист1").ListObjects("Ссылки")
    If tbl Is Nothing Then
        g_logger.LogMessage LL_ERROR, 0, "Таблица 'Ссылки' не найдена на листе Sheet1"
        LoadTasksFromTable = 0
        Exit Function
    End If
    
    ' Поиск колонок
    urlCol = GetColumnIndex(tbl, "Ссылка")
    pathCol = GetColumnIndex(tbl, "Путь для сохранения")
    
    If urlCol = 0 Then
        g_logger.LogMessage LL_ERROR, 0, "Колонка 'URL' не найдена"
        LoadTasksFromTable = 0
        Exit Function
    End If
    
    If pathCol = 0 Then
        g_logger.LogMessage LL_ERROR, 0, "Колонка 'Путь для сохранения' не найдена"
        LoadTasksFromTable = 0
        Exit Function
    End If
    
    Set dataRange = tbl.DataBodyRange
    If dataRange Is Nothing Then
        LoadTasksFromTable = 0
        Exit Function
    End If
    
    For i = 1 To dataRange.Rows.count
        Dim url As String
        Dim destPath As String
        
        url = Trim(CStr(dataRange.Cells(i, urlCol).Value))
        destPath = Trim(CStr(dataRange.Cells(i, pathCol).Value))
        
        If url <> "" And destPath <> "" Then
            g_manager.AddTask url, destPath
            count = count + 1
        End If
    Next i
    
    LoadTasksFromTable = count
    On Error GoTo 0
End Function

' ============================================================================
' Получение индекса колонки по имени
' ============================================================================
Private Function GetColumnIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name = colName Then
            GetColumnIndex = col.Index
            Exit Function
        End If
    Next col
    GetColumnIndex = 0
End Function

' ============================================================================
' Подключение событий менеджера
' ============================================================================
Private Sub HookManagerEvents()
    ' В VBA нельзя использовать WithEvents на переменной, объявленной через New
    ' Поэтому используем вспомогательный класс-обёртку
    ' Для упрощения: события обрабатываются через таймер
End Sub

' ============================================================================
' Очистка ресурсов
' ============================================================================
Private Sub Cleanup()
    On Error Resume Next
    
    If Not g_manager Is Nothing Then
        g_manager.StopAll
        Set g_manager = Nothing
    End If
    
    If Not g_logger Is Nothing Then
        g_logger.CloseLogFile
        Set g_logger = Nothing
    End If
    
    On Error GoTo 0
End Sub

' ============================================================================
' Обработчик потери состояния (вызывается StateLossCallback)
' ============================================================================
Public Sub OnWorkerLostState(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String)
    On Error Resume Next
    
    Dim logMsg As String
    logMsg = Now & " [STATE_LOSS] Worker #" & TaskId & " lost state. URL: " & url
    
    ' Запись в критический лог
    Dim critFile As Integer
    critFile = FreeFile
    Open g_criticalErrorLogPath For Append As #critFile
    Print #critFile, logMsg
    Close #critFile
    
    ' Попытка восстановления через менеджер
    If Not g_manager Is Nothing Then
        ' Перезапуск задачи
        g_manager.AddTask url, destPath
    End If
    
    On Error GoTo 0
End Sub

' ============================================================================
' Остановка всех загрузок (можно вызвать из кнопки)
' ============================================================================
Public Sub StopAllDownloads()
    If Not g_manager Is Nothing Then
        g_manager.StopAll
        Debug.Print "Все загрузки остановлены пользователем"
    End If
End Sub

' ============================================================================
' Приостановка загрузок
' ============================================================================
Public Sub PauseDownloads()
    If Not g_manager Is Nothing Then
        g_manager.Pause
        Debug.Print "Загрузки приостановлены"
    End If
End Sub

' ============================================================================
' Возобновление загрузок
' ============================================================================
Public Sub Resume_Downloads()
    If Not g_manager Is Nothing Then
        g_manager.Resume_
        Debug.Print "Загрузки возобновлены"
    End If
End Sub

