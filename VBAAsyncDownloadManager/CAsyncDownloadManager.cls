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
' Назначение: Управление очередью загрузок с ограничением параллельности
' ============================================================================
Option Explicit

Private m_maxConcurrent   As Long
Private m_callback        As IDownloadCallback
Private m_pendingTasks    As Collection   ' элементы – массивы {taskId, url, destPath}
Private m_activeWorkers   As Collection   ' ключ – taskId, значение – CAsyncDownloadWorker
Private m_nextTaskId      As Long
Private m_isRunning       As Boolean

Public Event AllTasksCompleted()

' Инициализация менеджера
Public Sub Init(ByVal maxConcurrent As Long, ByVal callback As IDownloadCallback)
    m_maxConcurrent = maxConcurrent
    Set m_callback = callback
    Set m_pendingTasks = New Collection
    Set m_activeWorkers = New Collection
    m_nextTaskId = 1
    m_isRunning = False
End Sub

' Добавление задачи в очередь
Public Sub AddTask(ByVal url As String, ByVal destPath As String)
    Dim taskData As Variant
    taskData = Array(m_nextTaskId, url, destPath)
    m_pendingTasks.Add taskData
    m_nextTaskId = m_nextTaskId + 1
End Sub

' Запуск обработки очереди
Public Sub Start()
    If m_isRunning Then Exit Sub
    m_isRunning = True
    Dim i As Long
    For i = 1 To m_maxConcurrent
        If Not LaunchNextWorker() Then Exit For
    Next i
End Sub

' Запускает следующий ожидающий рабочий процесс, если есть
Private Function LaunchNextWorker() As Boolean
    If m_pendingTasks.Count = 0 Then
        LaunchNextWorker = False
        Exit Function
    End If
    
    Dim taskData As Variant
    taskData = m_pendingTasks.Item(1)
    m_pendingTasks.Remove 1
    
    Dim taskId As Long: taskId = taskData(0)
    Dim url As String: url = taskData(1)
    Dim destPath As String: destPath = taskData(2)
    
    Dim worker As CAsyncDownloadWorker
    Set worker = New CAsyncDownloadWorker
    worker.Init taskId, url, destPath, m_callback
    worker.SetManager Me
    worker.Start
    
    ' Сохраняем активный воркер
    m_activeWorkers.Add worker, CStr(taskId)
    LaunchNextWorker = True
End Function

' Вызывается рабочим процессом по завершении (успех/ошибка)
Public Sub OnWorkerComplete(ByVal taskId As Long, ByVal success As Boolean)
    ' Удаляем из активных
    On Error Resume Next
    m_activeWorkers.Remove CStr(taskId)
    On Error GoTo 0
    
    ' Запускаем следующий, если есть
    LaunchNextWorker
    
    ' Если все задачи выполнены – генерируем событие
    If m_activeWorkers.Count = 0 And m_pendingTasks.Count = 0 Then
        m_isRunning = False
        RaiseEvent AllTasksCompleted
    End If
End Sub

' Остановка всех активных загрузок
Public Sub StopAll()
    Dim worker As CAsyncDownloadWorker
    For Each worker In m_activeWorkers
        worker.Cancel
    Next worker
    Set m_activeWorkers = New Collection
    Set m_pendingTasks = New Collection
    m_isRunning = False
End Sub

' Возвращает True, если есть активные или ожидающие задачи
Public Property Get IsBusy() As Boolean
    IsBusy = (m_activeWorkers.Count > 0) Or (m_pendingTasks.Count > 0)
End Property

