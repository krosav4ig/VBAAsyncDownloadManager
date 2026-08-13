VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAsyncDownloadManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_netMonitor As CNetworkMonitor
Private m_maxConcurrent As Long
Private m_callback As IDownloadCallback
Private m_pendingTasks As Collection
Private m_activeWorkers As Collection
Private m_nextTaskId As Long
Private m_isRunning As Boolean
Private m_isPaused As Boolean

Private m_resolveTimeout As Long
Private m_connectTimeout As Long
Private m_sendTimeout As Long
Private m_receiveTimeout As Long
Private m_bufferSize As Long
Private m_checkIntervalSeconds As Double
Private m_maxRetriesPerTask As Integer

Private m_nextTimer As Date
Private m_timerScheduled As Boolean
Private m_timerCallbackName As String

Private m_postProcessMacroName As String
Public m_postProcessQueue As Collection

Public m_completedCount As Long
Public m_failedCount As Long
Public m_totalTasks As Long

Public m_xlApp As Application

Public Event AllTasksCompleted()
Public Event TaskCompleted(ByVal TaskId As Long, ByVal success As Boolean)
Public Event ProgressUpdated()

Public Sub Init(ByVal maxConcurrent As Long, ByVal callback As IDownloadCallback, _
                Optional ByVal resolveTimeoutMs As Long = 30000, _
                Optional ByVal connectTimeoutMs As Long = 60000, _
                Optional ByVal sendTimeoutMs As Long = 30000, _
                Optional ByVal receiveTimeoutMs As Long = 60000, _
                Optional ByVal bufferSizeBytes As Long = 65536, _
                Optional ByVal checkIntervalSec As Double = 5, _
                Optional ByVal maxRetries As Integer = 3, _
                Optional ByVal maxNetworkUsagePercent As Long = 0, _
                Optional ByVal networkInterfaceFilter As String = "", _
                Optional ByVal postProcessMacroName As String = "")
    
    On Error GoTo ErrorHandler
    
    m_maxConcurrent = maxConcurrent
    If m_maxConcurrent < 1 Then m_maxConcurrent = 1
    If m_maxConcurrent > 10 Then m_maxConcurrent = 10
    
    Set m_callback = callback
    Set m_pendingTasks = New Collection
    Set m_activeWorkers = New Collection
    Set m_postProcessQueue = New Collection
    
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
    m_postProcessMacroName = postProcessMacroName
    
    
    If postProcessMacroName = "" Then
        Set m_xlApp = CreateObject("Excel.application")
        m_xlApp.Visible = 1
        With m_xlApp.Workbooks.Add
            .SaveAs Environ("tmp") & "\fileconverter.xlsm", 52, AddToMru:=False
            With .VBProject.VBComponents.Add(vbext_ct_StdModule)
                .Name = "modFileConverter"
                With .CodeModule
                    .InsertLines .CountOfLines + 1, "'test"
                End With
            End With
            
        End With
    End If
    
    m_completedCount = 0
    m_failedCount = 0
    m_totalTasks = 0
    
    Set m_netMonitor = New CNetworkMonitor
    m_netMonitor.Init maxNetworkUsagePercent, networkInterfaceFilter
    
    Dim limitStr As String
    If maxNetworkUsagePercent > 0 Then
        limitStr = maxNetworkUsagePercent & "%"
    Else
        limitStr = "Нет"
    End If
    
    LogMessage LL_INFO, "Менеджер инициализирован: maxConcurrent=" & m_maxConcurrent & _
                        ", Лимит сети=" & limitStr
    Exit Sub
ErrorHandler:
    LogMessage LL_CRITICAL, "Ошибка инициализации: " & Err.Description, Err.Number
End Sub

Public Sub AddTask(ByVal url As String, ByVal destPath As String)
    On Error GoTo ErrorHandler
    
    If url = "" Or destPath = "" Then
        LogMessage LL_WARNING, "Пропущена задача с пустым URL или путём"
        Exit Sub
    End If
    
    Dim normalizedPath As String
    normalizedPath = NormalizePath(destPath)
    
    Dim taskData As Variant
    taskData = Array(m_nextTaskId, url, normalizedPath)
    m_pendingTasks.Add taskData
    m_totalTasks = m_totalTasks + 1
    
    m_nextTaskId = m_nextTaskId + 1
    Exit Sub
ErrorHandler:
    LogMessage LL_ERROR, "Ошибка добавления задачи: " & Err.Description, Err.Number
End Sub

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
    
    Dim i As Long
    For i = 1 To m_maxConcurrent
        If Not LaunchNextWorker() Then Exit For
    Next i
    
    ScheduleTimer
End Sub

Public Sub Pause()
    If Not m_isRunning Then Exit Sub
    m_isPaused = True
    LogMessage LL_INFO, "Поставлен на паузу"
End Sub

Public Sub Resume_()
    If Not m_isRunning Then Exit Sub
    m_isPaused = False
    LogMessage LL_INFO, "Возобновлён"
    
    Dim i As Long
    For i = m_activeWorkers.count + 1 To m_maxConcurrent
        If Not LaunchNextWorker() Then Exit For
    Next i
End Sub

Public Sub StopAll()
    LogMessage LL_INFO, "Остановка всех загрузок"
    
    m_isRunning = False
    m_isPaused = False
    
    Dim worker As CAsyncDownloadWorker
    For Each worker In m_activeWorkers
        On Error Resume Next
        worker.Cancel
        On Error GoTo 0
    Next worker
    
    Set m_activeWorkers = New Collection
    Set m_pendingTasks = New Collection
    
    ' Отменяем все задачи постобработки
    On Error Resume Next
    While m_postProcessQueue.count > 0
        Dim cb As StateLossCallback
        Set cb = m_postProcessQueue.item(1)
        m_postProcessQueue.Remove 1
        Set cb = Nothing
    Wend
    On Error GoTo 0
    
    CancelTimer
    LogMessage LL_INFO, "Все загрузки остановлены"
End Sub

Private Sub ScheduleTimer()
    If Not m_isRunning Or m_timerScheduled Then Exit Sub
    
    m_nextTimer = Now + m_checkIntervalSeconds / 86400#
    m_timerScheduled = True
    
    On Error Resume Next
    Application.OnTime m_nextTimer, m_timerCallbackName, , True
    If Err.Number <> 0 Then
        LogMessage LL_ERROR, "Ошибка планирования таймера: " & Err.Description, Err.Number
        m_timerScheduled = False
    End If
    On Error GoTo 0
End Sub

Private Sub CancelTimer()
    If m_timerScheduled Then
        On Error Resume Next
        Application.OnTime m_nextTimer, m_timerCallbackName, , False
        On Error GoTo 0
        m_timerScheduled = False
    End If
End Sub

Public Sub TimerCheck()
    If Not m_isRunning Then Exit Sub
    On Error GoTo ErrorHandler
    
    m_timerScheduled = False
    
    CheckStuckWorkers
    CheckSilentCompletion
    
    If m_postProcessQueue.count > 0 Then
        Dim cb As StateLossCallback
        Set cb = m_postProcessQueue.item(1)
        m_postProcessQueue.Remove 1
        Set cb = Nothing
    End If
    
    CheckAsyncMarkers
    
    If Not m_isPaused Then
        Dim canLaunchNew As Boolean
        canLaunchNew = True
        
        If Not m_netMonitor Is Nothing And m_netMonitor.MaxUsagePercent > 0 Then
            m_netMonitor.Refresh
            If m_netMonitor.IsOverLimit Then
                canLaunchNew = False
                LogMessage LL_DEBUG, "Сеть перегружена: " & Format(m_netMonitor.CurrentUsagePercent, "0.0") & _
                                   "% (Лимит: " & m_netMonitor.MaxUsagePercent & "%). Ожидание..."
            End If
        End If
        
        If canLaunchNew Then
            Dim i As Long
            For i = m_activeWorkers.count + 1 To m_maxConcurrent
                If Not LaunchNextWorker() Then Exit For
            Next i
        End If
    End If
    
    ' Проверка завершения всех задач (включая постобработку)
    Dim pendingPostProcess As Long
    pendingPostProcess = 0
    If Not g_pendingAsyncMarkers Is Nothing Then pendingPostProcess = g_pendingAsyncMarkers.count
    
    If m_activeWorkers.count = 0 And m_pendingTasks.count = 0 And _
       m_postProcessQueue.count = 0 And pendingPostProcess = 0 And m_isRunning Then
        m_isRunning = False
        RaiseEvent AllTasksCompleted
    End If
    
    RaiseEvent ProgressUpdated
    ScheduleTimer
    Exit Sub
ErrorHandler:
    LogMessage LL_ERROR, "Ошибка в TimerCheck: " & Err.Description, Err.Number
    On Error Resume Next
    ScheduleTimer
    On Error GoTo 0
End Sub

Private Sub CheckStuckWorkers()
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
        Set worker = m_activeWorkers(CStr(key))
        On Error GoTo 0
        
        If Not worker Is Nothing Then
            On Error Resume Next
            If worker.IsStuck() Then
                LogMessage LL_WARNING, "Обнаружен зависший воркер #" & key & ", рестарт"
                worker.Restart
            End If
            On Error GoTo 0
        End If
    Next key
End Sub

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
        Set worker = m_activeWorkers(CStr(key))
        On Error GoTo 0
        
        If Not worker Is Nothing Then
            On Error Resume Next
            If worker.CheckAndCompleteIfFileExists() Then
                LogMessage LL_INFO, "Тихое завершение задачи #" & key
            End If
            On Error GoTo 0
        End If
    Next key
End Sub

Private Function LaunchNextWorker() As Boolean
    If m_isPaused Or m_pendingTasks.count = 0 Then
        LaunchNextWorker = False
        Exit Function
    End If
    
    Dim taskData As Variant
    On Error Resume Next
    taskData = m_pendingTasks.item(1)
    If Err.Number <> 0 Then
        LaunchNextWorker = False
        Exit Function
    End If
    On Error GoTo 0
    
    Dim TaskId As Long: TaskId = taskData(0)
    Dim url As String: url = taskData(1)
    Dim destPath As String: destPath = taskData(2)
    
    Dim worker As CAsyncDownloadWorker
    Set worker = New CAsyncDownloadWorker
    
    worker.Init TaskId, url, destPath, m_callback, m_bufferSize, _
                m_resolveTimeout, m_connectTimeout, m_sendTimeout, m_receiveTimeout, _
                m_maxRetriesPerTask, m_postProcessMacroName
    worker.SetManager Me
    
    On Error Resume Next
    m_pendingTasks.Remove 1
    m_activeWorkers.Add worker, CStr(TaskId)
    On Error GoTo 0
    
    worker.Start
    LaunchNextWorker = True
End Function

' ============================================================================
' Вызывается воркером сразу после скачивания файла
' Освобождает слот для новых загрузок, не дожидаясь постобработки!
' ============================================================================
Public Sub OnDownloadComplete(ByVal TaskId As Long, ByVal destPath As String, ByVal macroName As String)
    On Error GoTo ErrorHandler
    
    LogMessage LL_INFO, "OnDownloadComplete: TaskId='" & TaskId & ",  macroName='" & macroName & "', destPath='" & destPath & "'"
    
    ' 1. Мгновенно освобождаем слот и запускаем следующую задачу
    On Error Resume Next
    m_activeWorkers.Remove CStr(TaskId)
    On Error GoTo 0
    
    Dim i As Long
    For i = m_activeWorkers.count + 1 To m_maxConcurrent
        If Not LaunchNextWorker() Then Exit For
    Next i
    
    ' 2. Если нужна постобработка - ставим в очередь через StateLossCallback
    If macroName <> "" And destPath <> "" Then
        Dim cb As StateLossCallback
        Set cb = New StateLossCallback
        cb.InitByMacroName "ExecutePostProcess", TaskId, destPath, macroName
        
        m_postProcessQueue.Add cb
        LogMessage LL_INFO, "Задача #" & TaskId & " скачана. Ожидает постобработки."
    Else
        FinalizeTask TaskId, True
    End If
    Exit Sub
ErrorHandler:
    LogMessage LL_ERROR, "Ошибка в OnDownloadComplete: " & Err.Description, Err.Number
    FinalizeTask TaskId, False
End Sub

' ============================================================================
' Финализация задачи (вызывается после успешной постобработки или при ошибке)
' ============================================================================
Public Sub FinalizeTask(ByVal TaskId As Long, ByVal success As Boolean)
    On Error GoTo ErrorHandler
    
    If success Then
        m_completedCount = m_completedCount + 1
    Else
        m_failedCount = m_failedCount + 1
    End If
    
    LogMessage LL_DEBUG, "Финализация задачи #" & TaskId & ", успех=" & success
    RaiseEvent TaskCompleted(TaskId, success)
    Exit Sub
ErrorHandler:
    LogMessage LL_ERROR, "Ошибка в FinalizeTask: " & Err.Description, Err.Number
End Sub

Public Property Get IsBusy() As Boolean
    Dim pendingPostProcess As Long
    pendingPostProcess = 0
    If Not g_pendingAsyncMarkers Is Nothing Then pendingPostProcess = g_pendingAsyncMarkers.count
    IsBusy = (m_activeWorkers.count > 0) Or (m_pendingTasks.count > 0) Or _
             (m_postProcessQueue.count > 0) Or (pendingPostProcess > 0)
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

Private Function NormalizePath(ByVal filePath As String) As String
    Dim result As String
    result = Trim$(filePath)
    
    Dim isUNC As Boolean
    isUNC = (Left$(result, 2) = "\\")
    
    result = Replace$(result, "/", "\")
    
    Do While InStr(result, "\\") > 0
        result = Replace$(result, "\\", "\")
    Loop
    
    If isUNC And Left$(result, 2) <> "\\" Then result = "\" & result
    
    NormalizePath = result
End Function

Private Sub LogMessage(ByVal level As Long, ByVal message As String, _
                       Optional ByVal errorCode As Long = 0)
    If Not m_callback Is Nothing Then
        m_callback.OnLog level, 0, "[Manager] " & message, errorCode
    End If
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If m_timerScheduled Then
        Application.OnTime m_nextTimer, m_timerCallbackName, , False
    End If
    m_isRunning = False
    
    Dim worker As CAsyncDownloadWorker
    For Each worker In m_activeWorkers
        worker.Cancel
    Next worker
    
    Set m_netMonitor = Nothing
    Set m_pendingTasks = Nothing
    Set m_activeWorkers = Nothing
    Set m_postProcessQueue = Nothing
    Set m_callback = Nothing
    
    'm_xlApp.Quit
    
    On Error GoTo 0
End Sub

