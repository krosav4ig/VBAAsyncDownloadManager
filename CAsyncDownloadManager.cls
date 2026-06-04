VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAsyncDownloadManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' Класс: CAsyncDownloadManager
' Назначение: Управление очередью загрузок с пулом воркеров
' ============================================================================
Option Explicit

Private m_maxConcurrent As Long
Private m_callback As IDownloadCallback
Private m_pendingTasks As Collection
Private m_activeWorkers As Collection
Private m_nextTaskId As Long
Private m_isRunning As Boolean
Private m_isPaused As Boolean

' Конфигурация
Private m_resolveTimeout As Long
Private m_connectTimeout As Long
Private m_sendTimeout As Long
Private m_receiveTimeout As Long
Private m_bufferSize As Long
Private m_checkIntervalSeconds As Double
Private m_maxRetriesPerTask As Integer

' Таймер
Private m_nextTimer As Date
Private m_timerScheduled As Boolean
Private m_timerCallbackName As String

' Синхронизация (имитация)
Private m_syncLock As Object

' Статистика
Private m_completedCount As Long
Private m_failedCount As Long
Private m_totalTasks As Long

Public Event AllTasksCompleted()
Public Event TaskCompleted(ByVal TaskId As Long, ByVal success As Boolean)
Public Event ProgressUpdated(ByVal ActiveCount As Long, ByVal PendingCount As Long)

' ============================================================================
' Инициализация
' ============================================================================
Public Sub Init(ByVal maxConcurrent As Long, ByVal callback As IDownloadCallback, _
                Optional ByVal resolveTimeoutMs As Long = 30000, _
                Optional ByVal connectTimeoutMs As Long = 60000, _
                Optional ByVal sendTimeoutMs As Long = 30000, _
                Optional ByVal receiveTimeoutMs As Long = 60000, _
                Optional ByVal bufferSizeBytes As Long = 65536, _
                Optional ByVal checkIntervalSec As Double = 5, _
                Optional ByVal maxRetries As Integer = 3)
    
    On Error GoTo ErrorHandler
    
    m_maxConcurrent = maxConcurrent
    If m_maxConcurrent < 1 Then m_maxConcurrent = 1
    If m_maxConcurrent > 10 Then m_maxConcurrent = 10  ' Ограничение
    
    Set m_callback = callback
    Set m_pendingTasks = New Collection
    Set m_activeWorkers = New Collection
    Set m_syncLock = CreateObject("Scripting.Dictionary")
    m_syncLock.Add "lock", 1
    
    m_nextTaskId = 1
    m_isRunning = False
    m_isPaused = False
    m_timerScheduled = False
    m_timerCallbackName = "CAsyncDownloadManager_TimerCheck"
    
    m_resolveTimeout = resolveTimeoutMs
    m_connectTimeout = connectTimeoutMs
    m_sendTimeout = sendTimeoutMs
    m_receiveTimeout = receiveTimeoutMs
    m_bufferSize = bufferSizeBytes
    m_checkIntervalSeconds = checkIntervalSec
    m_maxRetriesPerTask = maxRetries
    
    m_completedCount = 0
    m_failedCount = 0
    m_totalTasks = 0
    
    LogMessage LL_INFO, "Менеджер инициализирован: maxConcurrent=" & m_maxConcurrent
    
    Exit Sub
    
ErrorHandler:
    LogMessage LL_CRITICAL, "Ошибка инициализации: " & Err.Description, Err.Number
End Sub

' ============================================================================
' Добавление задачи
' ============================================================================
Public Sub AddTask(ByVal url As String, ByVal destPath As String)
    On Error GoTo ErrorHandler
    
    If url = "" Or destPath = "" Then
        LogMessage LL_WARNING, "Пропущена задача с пустым URL или путём"
        Exit Sub
    End If
    
    ' Нормализация пути
    Dim normalizedPath As String
    normalizedPath = NormalizePath(destPath)
    
    Dim taskData As Variant
    taskData = Array(m_nextTaskId, url, normalizedPath)
    m_pendingTasks.Add taskData
    m_totalTasks = m_totalTasks + 1
    
    LogMessage LL_DEBUG, "Добавлена задача #" & m_nextTaskId & ": " & url
    
    m_nextTaskId = m_nextTaskId + 1
    
    Exit Sub
    
ErrorHandler:
    LogMessage LL_ERROR, "Ошибка добавления задачи: " & Err.Description, Err.Number
End Sub

' ============================================================================
' Добавление нескольких задач из массива
' ============================================================================
Public Sub AddTasks(ByVal urls As Variant, ByVal destPaths As Variant)
    Dim i As Long
    Dim count As Long
    
    If IsArray(urls) And IsArray(destPaths) Then
        count = UBound(urls) - LBound(urls) + 1
        For i = 0 To count - 1
            If i < UBound(destPaths) - LBound(destPaths) + 1 Then
                AddTask urls(i), destPaths(i)
            End If
        Next i
    End If
End Sub

' ============================================================================
' Запуск загрузок
' ============================================================================
Public Sub Start()
    If m_isRunning Then
        LogMessage LL_WARNING, "Менеджер уже запущен"
        Exit Sub
    End If
    
    If m_pendingTasks.count = 0 And m_activeWorkers.count = 0 Then
        LogMessage LL_WARNING, "Нет задач для выполнения"
        Exit Sub
    End If
    
    m_isRunning = True
    m_isPaused = False
    
    LogMessage LL_INFO, "Запуск менеджера. Всего задач: " & m_totalTasks
    
    ' Запуск начальной партии воркеров
    Dim i As Long
    For i = 1 To m_maxConcurrent
        If Not LaunchNextWorker() Then Exit For
    Next i
    
    ' Запуск таймера
    ScheduleTimer
End Sub

' ============================================================================
' Пауза
' ============================================================================
Public Sub Pause()
    If Not m_isRunning Then Exit Sub
    m_isPaused = True
    LogMessage LL_INFO, "Поставлен на паузу"
End Sub

' ============================================================================
' Возобновление
' ============================================================================
Public Sub Resume_()
    If Not m_isRunning Then Exit Sub
    m_isPaused = False
    LogMessage LL_INFO, "Возобновлён"
    
    ' Заполнение пула
    Dim i As Long
    For i = m_activeWorkers.count To m_maxConcurrent - 1
        If Not LaunchNextWorker() Then Exit For
    Next i
End Sub

' ============================================================================
' Остановка всех загрузок
' ============================================================================
Public Sub StopAll()
    LogMessage LL_INFO, "Остановка всех загрузок"
    
    m_isRunning = False
    m_isPaused = False
    
    ' Отмена всех активных воркеров
    Dim worker As CAsyncDownloadWorker
    For Each worker In m_activeWorkers
        On Error Resume Next
        worker.Cancel
        On Error GoTo 0
    Next worker
    
    ' Очистка коллекций
    Set m_activeWorkers = New Collection
    Set m_pendingTasks = New Collection
    
    ' Отмена таймера
    CancelTimer
    
    LogMessage LL_INFO, "Все загрузки остановлены"
End Sub

' ============================================================================
' Планирование таймера
' ============================================================================
Private Sub ScheduleTimer()
    If Not m_isRunning Then Exit Sub
    If m_timerScheduled Then Exit Sub
    
    m_nextTimer = Now + m_checkIntervalSeconds / 86400
    m_timerScheduled = True
    
    On Error Resume Next
    Application.OnTime m_nextTimer, m_timerCallbackName, , True
    If Err.Number <> 0 Then
        LogMessage LL_ERROR, "Ошибка планирования таймера: " & Err.Description, Err.Number
        m_timerScheduled = False
    End If
    On Error GoTo 0
End Sub

' ============================================================================
' Отмена таймера
' ============================================================================
Private Sub CancelTimer()
    If m_timerScheduled Then
        On Error Resume Next
        Application.OnTime m_nextTimer, m_timerCallbackName, , False
        On Error GoTo 0
        m_timerScheduled = False
    End If
End Sub

' ============================================================================
' Публичный метод для вызова по таймеру (вызывается из Module1)
' ============================================================================
Public Sub TimerCheck()
    If Not m_isRunning Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    ' Сброс флага для следующего планирования
    m_timerScheduled = False
    
    ' Проверка зависших воркеров
    CheckStuckWorkers
    
    ' Проверка завершённых (тихих) воркеров
    CheckSilentCompletion
    
    ' Запуск новых воркеров при наличии места
    If Not m_isPaused Then
        Dim i As Long
        For i = m_activeWorkers.count To m_maxConcurrent - 1
            If Not LaunchNextWorker() Then Exit For
        Next i
    End If
    
    ' Проверка полного завершения
    If m_activeWorkers.count = 0 And m_pendingTasks.count = 0 And m_isRunning Then
        m_isRunning = False
        LogMessage LL_INFO, "Все задачи завершены. Успешно: " & m_completedCount & ", Ошибок: " & m_failedCount
        RaiseEvent AllTasksCompleted
    End If
    
    ' Уведомление о прогрессе
    RaiseEvent ProgressUpdated(m_activeWorkers.count, m_pendingTasks.count)
    
    ' Планирование следующего вызова
    ScheduleTimer
    
    Exit Sub
    
ErrorHandler:
    LogMessage LL_ERROR, "Ошибка в TimerCheck: " & Err.Description, Err.Number
    ScheduleTimer
End Sub

' ============================================================================
' Проверка зависших воркеров
' ============================================================================
Private Sub CheckStuckWorkers()
    Dim worker As CAsyncDownloadWorker
    Dim keys As Collection
    Dim key As Variant
    
    Set keys = New Collection
    
    ' Сбор ключей (безопасная итерация)
    On Error Resume Next
    For Each worker In m_activeWorkers
        keys.Add CStr(worker.TaskId)
    Next worker
    On Error GoTo 0
    
    For Each key In keys
        On Error Resume Next
        Set worker = m_activeWorkers(key)
        On Error GoTo 0
        
        If Not worker Is Nothing Then
            If worker.IsStuck() Then
                LogMessage LL_WARNING, "Обнаружен зависший воркер #" & key & ", рестарт"
                worker.Restart
            End If
        End If
    Next key
End Sub

' ============================================================================
' Проверка "тихого" завершения (когда событие не пришло)
' ============================================================================
Private Sub CheckSilentCompletion()
    Dim worker As CAsyncDownloadWorker
    Dim keys As Collection
    Dim key As Variant
    
    Set keys = New Collection
    
    On Error Resume Next
    For Each worker In m_activeWorkers
        keys.Add CStr(worker.TaskId)
    Next worker
    On Error GoTo 0
    
    For Each key In keys
        On Error Resume Next
        Set worker = m_activeWorkers(key)
        On Error GoTo 0
        
        If Not worker Is Nothing Then
            If worker.CheckAndCompleteIfFileExists() Then
                ' Воркер сам вызовет OnWorkerComplete
                LogMessage LL_INFO, "Тихое завершение задачи #" & key
            End If
        End If
    Next key
End Sub

' ============================================================================
' Запуск следующего воркера из очереди
' ============================================================================
Private Function LaunchNextWorker() As Boolean
    If m_isPaused Then
        LaunchNextWorker = False
        Exit Function
    End If
    
    If m_pendingTasks.count = 0 Then
        LaunchNextWorker = False
        Exit Function
    End If
    
    ' Безопасное извлечение
    Dim taskData As Variant
    On Error Resume Next
    taskData = m_pendingTasks.Item(1)
    If Err.Number <> 0 Then
        LaunchNextWorker = False
        Exit Function
    End If
    m_pendingTasks.Remove 1
    On Error GoTo 0
    
    Dim TaskId As Long: TaskId = taskData(0)
    Dim url As String: url = taskData(1)
    Dim destPath As String: destPath = taskData(2)
    
    ' Создание и запуск воркера
    Dim worker As CAsyncDownloadWorker
    Set worker = New CAsyncDownloadWorker
    
    ' Добавление в активные
    On Error Resume Next
    m_activeWorkers.Add worker, CStr(TaskId)
    On Error GoTo 0
    
    worker.Init TaskId, url, destPath, m_callback, m_bufferSize, _
                m_resolveTimeout, m_connectTimeout, m_sendTimeout, m_receiveTimeout, _
                m_maxRetriesPerTask
    worker.SetManager Me
    worker.Start
    
    
    
    LogMessage LL_DEBUG, "Запущен воркер #" & TaskId & " (активно: " & m_activeWorkers.count & ")"
    
    LaunchNextWorker = True
End Function

' ============================================================================
' Обработка завершения воркера
' ============================================================================
Public Sub OnWorkerComplete(ByVal TaskId As Long, ByVal success As Boolean)
    On Error GoTo ErrorHandler
    
    ' Обновление статистики
    If success Then
        m_completedCount = m_completedCount + 1
    Else
        m_failedCount = m_failedCount + 1
    End If
    
    ' Удаление из активных
    On Error Resume Next
    m_activeWorkers.Remove CStr(TaskId)
    On Error GoTo 0
    
    LogMessage LL_DEBUG, "Завершён воркер #" & TaskId & ", успех=" & success & " (активно: " & m_activeWorkers.count & ")"
    
    RaiseEvent TaskCompleted(TaskId, success)
    
    Exit Sub
    
ErrorHandler:
    LogMessage LL_ERROR, "Ошибка в OnWorkerComplete: " & Err.Description, Err.Number
End Sub

' ============================================================================
' Свойства
' ============================================================================
Public Property Get IsBusy() As Boolean
    IsBusy = (m_activeWorkers.count > 0) Or (m_pendingTasks.count > 0)
End Property

Public Property Get IsRunning() As Boolean
    IsRunning = m_isRunning
End Property

Public Property Get ActiveCount() As Long
    ActiveCount = m_activeWorkers.count
End Property

Public Property Get PendingCount() As Long
    PendingCount = m_pendingTasks.count
End Property

Public Property Get CompletedCount() As Long
    CompletedCount = m_completedCount
End Property

Public Property Get FailedCount() As Long
    FailedCount = m_failedCount
End Property

' ============================================================================
' Вспомогательные методы
' ============================================================================
Private Function NormalizePath(ByVal filePath As String) As String
    Dim result As String
    result = Trim(filePath)
    
    ' Замена слешей
    result = Replace(result, "/", "\")
    
    ' Удаление лишних обратных слешей
    Do While InStr(result, "\\") > 0
        result = Replace(result, "\\", "\")
    Loop
    
    NormalizePath = result
End Function

Private Sub LogMessage(ByVal level As Long, ByVal message As String, _
                       Optional ByVal errorCode As Long = 0)
    If Not m_callback Is Nothing Then
        m_callback.OnLog level, 0, "[Manager] " & message, errorCode
    End If
End Sub

' ============================================================================
' Деструктор
' ============================================================================
Private Sub Class_Terminate()
    StopAll
    Set m_pendingTasks = Nothing
    Set m_activeWorkers = Nothing
    Set m_callback = Nothing
    Set m_syncLock = Nothing
End Sub

