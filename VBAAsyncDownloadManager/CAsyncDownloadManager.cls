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
' Модуль: CAsyncDownloadManager.cls
' Назначение: Управление очередью с таймаутами, буферизацией, таймером проверки
' ============================================================================
Option Explicit

Private m_maxConcurrent   As Long
Private m_callback        As IDownloadCallback
Private m_pendingTasks    As Collection
Private m_activeWorkers   As Collection
Private m_nextTaskId      As Long
Private m_isRunning       As Boolean

' Новые параметры
Private m_resolveTimeout As Long
Private m_connectTimeout As Long
Private m_sendTimeout As Long
Private m_receiveTimeout As Long
Private m_bufferSize As Long
Private m_checkIntervalSeconds As Double   ' интервал таймера проверки файлов

Private m_nextTimer As Date

Public Event AllTasksCompleted()

' Инициализация с параметрами таймаутов и буфера
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
    Dim i As Long
    For i = 1 To m_maxConcurrent
        If Not LaunchNextWorker() Then Exit For
    Next i
    ' Запускаем таймер проверки
    ScheduleTimer
End Sub

' Запуск таймера
Private Sub ScheduleTimer()
    If Not m_isRunning Then Exit Sub
    m_nextTimer = Now + m_checkIntervalSeconds / 86400
    Application.OnTime m_nextTimer, "CAsyncDownloadManager_TimerCheck", , True
End Sub

' Публичная процедура для вызова по таймеру (должна быть в стандартном модуле, но вызовем через Application.Run)
Public Sub TimerCheck()
    If Not m_isRunning Then Exit Sub
    CheckAllWorkersCompletion
    ScheduleTimer
End Sub

' Проверка всех активных воркеров на наличие уже сохранённых файлов
Private Sub CheckAllWorkersCompletion()
    Dim worker As CAsyncDownloadWorker
    Dim keys As Variant
    ReDim keys(1 To m_activeWorkers.Count)
    Dim i As Long: i = 1
    For Each worker In m_activeWorkers
        keys(i) = CStr(worker.TaskId)   ' необходимо добавить свойство taskId в воркер (см. ниже)
        i = i + 1
    Next
    ' Чтобы избежать изменения коллекции во время итерации, проходим по сохранённым ключам
    For i = 1 To UBound(keys)
        On Error Resume Next
        Set worker = m_activeWorkers(keys(i))
        On Error GoTo 0
        If Not worker Is Nothing Then
            If worker.CheckAndCompleteIfFileExists Then
                ' Воркер сам вызовет OnWorkerComplete, он будет удалён из m_activeWorkers
            End If
        End If
    Next
End Sub

' Запуск следующего воркера
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

Public Sub OnWorkerComplete(ByVal TaskId As Long, ByVal success As Boolean)
    On Error Resume Next
    m_activeWorkers.Remove CStr(TaskId)
    On Error GoTo 0
    LaunchNextWorker
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

' Деструктор – отключаем таймер
Private Sub Class_Terminate()
    If m_isRunning Then
        On Error Resume Next
        Application.OnTime m_nextTimer, "CAsyncDownloadManager_TimerCheck", , False
        On Error GoTo 0
    End If
End Sub

