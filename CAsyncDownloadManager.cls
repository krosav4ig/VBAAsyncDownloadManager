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
' Класс: CAsyncDownloadManager.cls
' Назначение: Управление очередью загрузок, пул воркеров, таймер проверки зависаний.
' ============================================================================
Option Explicit

Private m_maxConcurrent   As Long
Private m_callback        As IDownloadCallback
Private m_pendingTasks    As Collection
Private m_activeWorkers   As Collection
Private m_nextTaskId      As Long
Private m_isRunning       As Boolean

' Параметры конфигурации
Private m_resolveTimeout As Long
Private m_connectTimeout As Long
Private m_sendTimeout As Long
Private m_receiveTimeout As Long
Private m_bufferSize As Long
Private m_checkIntervalSeconds As Double

Private m_nextTimer As Date

Public Event AllTasksCompleted()

' Инициализация
Public Sub Init(ByVal maxConcurrent As Long, ByVal callback As IDownloadCallback, _
                Optional ByVal resolveTimeoutMs As Long = 5000, _
                Optional ByVal connectTimeoutMs As Long = 60000, _
                Optional ByVal sendTimeoutMs As Long = 30000, _
                Optional ByVal receiveTimeoutMs As Long = 30000, _
                Optional ByVal bufferSizeBytes As Long = 65536, _
                Optional ByVal checkIntervalSec As Double = 5)
    m_maxConcurrent = maxConcurrent
    Set m_callback = callback
    Set m_pendingTasks = New Collection
    Set m_activeWorkers = New Collection
    m_nextTaskId = 1
    m_isRunning = False
    
    m_resolveTimeout = resolveTimeoutMs
    m_connectTimeout = connectTimeoutMs
    m_sendTimeout = sendTimeoutMs
    m_receiveTimeout = receiveTimeoutMs
    m_bufferSize = bufferSizeBytes
    m_checkIntervalSeconds = checkIntervalSec
End Sub

Public Sub AddTask(ByVal url As String, ByVal destPath As String)
    Dim taskData As Variant
    taskData = Array(m_nextTaskId, url, destPath)
    m_pendingTasks.Add taskData
    m_nextTaskId = m_nextTaskId + 1
End Sub

Public Sub Start()
    If m_isRunning Then Exit Sub
    m_isRunning = True
    
    ' Запуск первой партии воркеров
    Dim i As Long
    For i = 1 To m_maxConcurrent
        If Not LaunchNextWorker() Then Exit For
    Next i
    
    ' Запуск таймера проверки
    ScheduleTimer
End Sub

' Планирование следующего вызова таймера
Private Sub ScheduleTimer()
    If Not m_isRunning Then Exit Sub
    m_nextTimer = Now + m_checkIntervalSeconds / 86400
    ' Используем Application.OnTime для вызова публичной процедуры в Module1
    On Error Resume Next
    Application.OnTime m_nextTimer, "CAsyncDownloadManager_TimerCheck", , True
    On Error GoTo 0
End Sub

' Публичный метод для вызова по таймеру
Public Sub TimerCheck()
    If Not m_isRunning Then Exit Sub
    
    ' Проверка всех активных воркеров на "тихое" завершение (файл есть, события нет)
    CheckAllWorkersCompletion
    
    ' Проверка на зависания и рестарт
    CheckStuckWorkers
    
    ' Продолжаем цикл таймера
    ScheduleTimer
End Sub

' Проверка соответствия размера файла (восстановление после сбоя событий)
Private Sub CheckAllWorkersCompletion()
    Dim worker As CAsyncDownloadWorker
    Dim keys As Variant
    Dim i As Long
    
    ' Копируем ключи, чтобы безопасно итерироваться
    ReDim keys(1 To m_activeWorkers.Count)
    i = 1
    For Each worker In m_activeWorkers
        keys(i) = CStr(worker.TaskId)
        i = i + 1
    Next
    
    For i = 1 To UBound(keys)
        On Error Resume Next
        Set worker = m_activeWorkers(keys(i))
        On Error GoTo 0
        
        If Not worker Is Nothing Then
            ' Если файл уже скачан полностью (например, предыдущая сессия упала в конце)
            If worker.CheckAndCompleteIfFileExists Then
                ' Воркер сам вызовет OnWorkerComplete и удалится из коллекции
            End If
        End If
    Next
End Sub

' Проверка на зависания и инициирование рестарта
Private Sub CheckStuckWorkers()
    Dim worker As CAsyncDownloadWorker
    Dim keys As Variant
    Dim i As Long
    
    ReDim keys(1 To m_activeWorkers.Count)
    i = 1
    For Each worker In m_activeWorkers
        keys(i) = CStr(worker.TaskId)
        i = i + 1
    Next
    
    For i = 1 To UBound(keys)
        On Error Resume Next
        Set worker = m_activeWorkers(keys(i))
        On Error GoTo 0
        
        If Not worker Is Nothing Then
            If worker.IsStuck() Then
                Debug.Print "Manager: Задача " & worker.TaskId & " зависла. Инициирую рестарт."
                worker.Restart
            End If
        End If
    Next
End Sub

' Запуск следующего воркера из очереди
Private Function LaunchNextWorker() As Boolean
    If m_pendingTasks.Count = 0 Then
        LaunchNextWorker = False
        Exit Function
    End If
    
    Dim taskData As Variant
    taskData = m_pendingTasks.Item(1)
    m_pendingTasks.Remove 1
    
    Dim TaskId As Long: TaskId = taskData(0)
    Dim url As String: url = taskData(1)
    Dim destPath As String: destPath = taskData(2)

    Dim worker As CAsyncDownloadWorker
    Set worker = New CAsyncDownloadWorker
    
    worker.Init TaskId, url, destPath, m_callback, m_bufferSize, _
                m_resolveTimeout, m_connectTimeout, m_sendTimeout, m_receiveTimeout
    worker.SetManager Me
    worker.Start

    m_activeWorkers.Add worker, CStr(TaskId)
    LaunchNextWorker = True
End Function

' Обработка завершения воркера
Public Sub OnWorkerComplete(ByVal TaskId As Long, ByVal success As Boolean)
    On Error Resume Next
    m_activeWorkers.Remove CStr(TaskId)
    On Error GoTo 0
    
    ' Запускаем следующую задачу, если есть места
    LaunchNextWorker
    
    ' Проверка полного завершения
    If m_activeWorkers.Count = 0 And m_pendingTasks.Count = 0 Then
        m_isRunning = False
        RaiseEvent AllTasksCompleted
    End If
End Sub

Public Sub StopAll()
    Dim worker As CAsyncDownloadWorker
    For Each worker In m_activeWorkers
        worker.Cancel
    Next worker
    Set m_activeWorkers = New Collection
    Set m_pendingTasks = New Collection
    m_isRunning = False
End Sub

Public Property Get IsBusy() As Boolean
    IsBusy = (m_activeWorkers.Count > 0) Or (m_pendingTasks.Count > 0)
End Property

Private Sub Class_Terminate()
    If m_isRunning Then
        On Error Resume Next
        Application.OnTime m_nextTimer, "CAsyncDownloadManager_TimerCheck", , False
        On Error GoTo 0
    End If
End Sub

