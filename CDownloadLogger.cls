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
' Класс: CDownloadLogger
' Назначение: Логирование загрузок в Excel и файл с уровнями детализации
' ============================================================================
Option Explicit
Implements IDownloadCallback

Private Const LOG_BUFFER_SIZE As Long = 100  ' Буфер для текстового лога

' Колонки таблицы
Private Enum LogCol
    colTaskId = 1
    colUrl = 2
    colDest = 3
    colStatus = 4
    colProgress = 5
    colStartTime = 6
    colEndTime = 7
    colError = 8
    colRetryCount = 9
    colDownloadedBytes = 10
End Enum

Private m_logSheet As Worksheet
Private m_logFileHandle As Integer
Private m_logFilePath As String
Private m_logLevel As Long
Private m_taskRowMap As Object
Private m_logBuffer As Collection
Private m_isFileOpen As Boolean
Private m_criticalErrors As Collection

' Синхронизация
Private m_syncLock As Object

' ============================================================================
' Инициализация
' ============================================================================
Public Sub Init(ByVal logSheet As Worksheet, Optional ByVal logFilePath As String = "", _
                Optional ByVal logLevel As Long = LL_INFO)
    On Error GoTo ErrorHandler
    
    Set m_logSheet = logSheet
    m_logFilePath = logFilePath
    m_logLevel = logLevel
    Set m_taskRowMap = CreateObject("Scripting.Dictionary")
    Set m_logBuffer = New Collection
    Set m_criticalErrors = New Collection
    Set m_syncLock = CreateObject("Scripting.Dictionary")
    m_syncLock.Add "lock", 1
    
    ' Инициализация таблицы
    If Not logSheet Is Nothing Then
        InitLogSheet
    End If
    
    ' Открытие файла лога
    If m_logFilePath <> "" Then
        OpenLogFile
    End If
    
    LogMessage LL_INFO, 0, "Логгер инициализирован. Уровень логирования: " & CStr(m_logLevel)
    Exit Sub
    
ErrorHandler:
    Debug.Print "CRITICAL: Ошибка инициализации логгера: " & Err.Description
End Sub

' ============================================================================
Private Sub InitLogSheet()
    On Error Resume Next
    With m_logSheet
        .Cells(1, LogCol.colTaskId).Value = "TaskID"
        .Cells(1, LogCol.colUrl).Value = "URL"
        .Cells(1, LogCol.colDest).Value = "Destination"
        .Cells(1, LogCol.colStatus).Value = "Status"
        .Cells(1, LogCol.colProgress).Value = "Progress %"
        .Cells(1, LogCol.colStartTime).Value = "Start Time"
        .Cells(1, LogCol.colEndTime).Value = "End Time"
        .Cells(1, LogCol.colError).Value = "Error"
        .Cells(1, LogCol.colRetryCount).Value = "Retries"
        .Cells(1, LogCol.colDownloadedBytes).Value = "Downloaded MB"
        
        .Rows(1).Font.Bold = True
        .Columns("A:J").AutoFit
    End With
    On Error GoTo 0
End Sub

' ============================================================================
Private Sub OpenLogFile()
    On Error GoTo ErrorHandler
    m_logFileHandle = FreeFile
    Open m_logFilePath For Append As #m_logFileHandle
    m_isFileOpen = True
    Exit Sub
    
ErrorHandler:
    m_isFileOpen = False
    Debug.Print "ERROR: Не удалось открыть файл лога: " & m_logFilePath
End Sub

' ============================================================================
Public Sub LogMessage(ByVal level As Long, ByVal TaskId As Long, _
                      ByVal message As String, Optional ByVal errorCode As Long = 0)
    On Error GoTo SafeExit
    
    If level > m_logLevel Then Exit Sub
    
    Dim timestamp As String
    timestamp = Format(Now, "yyyy-mm-dd HH:MM:SS")
    Dim prefix As String
    
    Select Case level
        Case LL_CRITICAL: prefix = "[CRITICAL]"
        Case LL_ERROR: prefix = "[ERROR]"
        Case LL_WARNING: prefix = "[WARN]"
        Case LL_INFO: prefix = "[INFO]"
        Case LL_DEBUG: prefix = "[DEBUG]"
    End Select
    
    Dim taskStr As String
    If TaskId > 0 Then taskStr = "T" & CStr(TaskId) Else taskStr = "SYS"
    
    Dim logMsg As String
    logMsg = timestamp & " " & prefix & " " & taskStr & ": " & message
    If errorCode <> 0 Then logMsg = logMsg & " (Err: " & errorCode & ")"
    
    ' Вывод в Immediate окно
    Debug.Print logMsg
    
    ' Запись в файл (буферизованная)
    If m_isFileOpen Then
        On Error Resume Next
        Print #m_logFileHandle, logMsg
        If level = LL_CRITICAL Then
            FlushLogFile
        End If
        On Error GoTo 0
    End If
    
    ' Критические ошибки сохраняем отдельно
    If level = LL_CRITICAL Then
        On Error Resume Next
        m_criticalErrors.Add logMsg
        On Error GoTo 0
    End If
    
SafeExit:
End Sub

' ============================================================================
Public Sub FlushLogFile()
    If m_isFileOpen Then
        On Error Resume Next
        Dim tempNum As Integer
        tempNum = FreeFile
        Close #m_logFileHandle
        Open m_logFilePath For Append As #m_logFileHandle
        On Error GoTo 0
    End If
End Sub

' ============================================================================
Public Sub CloseLogFile()
    If m_isFileOpen Then
        On Error Resume Next
        FlushLogFile
        Close #m_logFileHandle
        On Error GoTo 0
        m_isFileOpen = False
    End If
End Sub

' ============================================================================
Private Function GetTaskRow(ByVal TaskId As Long) As Variant
    Dim key As String
    key = CStr(TaskId)
    
    If m_taskRowMap.Exists(key) Then
        GetTaskRow = m_taskRowMap(key)
        Exit Function
    End If
    
    ' Поиск в таблице
    On Error Resume Next
    Dim cell As Range
    Set cell = m_logSheet.Columns(LogCol.colTaskId).Find(What:=TaskId, LookAt:=xlWhole)
    If Not cell Is Nothing Then
        m_taskRowMap.Add key, cell.row
        GetTaskRow = cell.row
    Else
        GetTaskRow = Null
    End If
    On Error GoTo 0
End Function

' ============================================================================
' Реализация IDownloadCallback
' ============================================================================
Private Sub IDownloadCallback_OnStart(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String)
    On Error GoTo ErrorHandler
    
    LogMessage LL_INFO, TaskId, "Начало загрузки: " & url
    
    If m_logSheet Is Nothing Then Exit Sub
    
    Dim newRow As Long
    newRow = m_logSheet.Cells(m_logSheet.Rows.count, LogCol.colTaskId).End(xlUp).row + 1
    If newRow < 2 Then newRow = 2
    
    With m_logSheet
        .Cells(newRow, LogCol.colTaskId).Value = TaskId
        .Cells(newRow, LogCol.colUrl).Value = url
        .Cells(newRow, LogCol.colDest).Value = destPath
        .Cells(newRow, LogCol.colStatus).Value = "Загрузка..."
        .Cells(newRow, LogCol.colProgress).Value = 0
        .Cells(newRow, LogCol.colStartTime).Value = Now
        .Cells(newRow, LogCol.colRetryCount).Value = 0
    End With
    
    m_taskRowMap.Add CStr(TaskId), newRow
    Exit Sub
    
ErrorHandler:
    LogMessage LL_ERROR, TaskId, "Ошибка OnStart: " & Err.Description, Err.Number
End Sub

' ============================================================================
Private Sub IDownloadCallback_OnProgress(ByVal TaskId As Long, ByVal bytesReceived As Currency, _
                                          ByVal totalBytes As Currency, ByVal percent As Integer)
    If percent < 0 Or percent > 100 Then Exit Sub
    
    ' Обновляем Excel только при изменении процента (оптимизация)
    Static lastPercent As Integer
    
    If percent Mod 5 = 0 Or percent = 100 Then
        Dim rowNumber As Variant
        rowNumber = GetTaskRow(TaskId)
        
        If Not IsNull(rowNumber) And IsNumeric(rowNumber) Then
            On Error Resume Next
            m_logSheet.Cells(CLng(rowNumber), LogCol.colProgress).Value = percent
            If totalBytes > 0 Then
                m_logSheet.Cells(CLng(rowNumber), LogCol.colDownloadedBytes).Value = _
                    Round(CDbl(bytesReceived) / 1048576, 2)
            End If
            On Error GoTo 0
        End If
        
        lastPercent = percent
    End If
    
    ' Логирование прогресса только на DEBUG уровне
    If m_logLevel >= LL_DEBUG Then
        LogMessage LL_DEBUG, TaskId, "Прогресс: " & percent & "%"
    End If
End Sub

' ============================================================================
Private Sub IDownloadCallback_OnComplete(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String, _
                                          ByVal success As Boolean, ByVal errorMsg As String)
    On Error GoTo ErrorHandler
    
    Dim status As String
    Dim logLevel As Long
    
    If success Then
        status = "Завершено"
        logLevel = LL_INFO
        LogMessage logLevel, TaskId, "Загрузка успешно завершена: " & destPath
    Else
        status = "Ошибка"
        logLevel = LL_ERROR
        LogMessage logLevel, TaskId, "Ошибка загрузки: " & errorMsg
    End If
    
    If m_logSheet Is Nothing Then GoTo SkipExcel
    
    Dim rowNumber As Variant
    rowNumber = GetTaskRow(TaskId)
    
    If Not IsNull(rowNumber) And IsNumeric(rowNumber) Then
        Dim row As Long: row = CLng(rowNumber)
        
        On Error Resume Next
        m_logSheet.Cells(row, LogCol.colStatus).Value = status
        If success Then
            m_logSheet.Cells(row, LogCol.colProgress).Value = 100
        End If
        m_logSheet.Cells(row, LogCol.colEndTime).Value = Now
        If Not success And errorMsg <> "" Then
            m_logSheet.Cells(row, LogCol.colError).Value = Left(errorMsg, 500)
        End If
        On Error GoTo 0
    End If
    
SkipExcel:
    Exit Sub
    
ErrorHandler:
    LogMessage LL_CRITICAL, TaskId, "Критическая ошибка в OnComplete: " & Err.Description, Err.Number
End Sub

' ============================================================================
Private Sub IDownloadCallback_OnLog(ByVal level As Long, ByVal TaskId As Long, _
                                     ByVal message As String, Optional ByVal errorCode As Long = 0)
    LogMessage level, TaskId, message, errorCode
End Sub

' ============================================================================
Public Sub UpdateRetryCount(ByVal TaskId As Long, ByVal retryCount As Integer)
    Dim rowNumber As Variant
    rowNumber = GetTaskRow(TaskId)
    
    If Not IsNull(rowNumber) And IsNumeric(rowNumber) Then
        On Error Resume Next
        m_logSheet.Cells(CLng(rowNumber), LogCol.colRetryCount).Value = retryCount
        On Error GoTo 0
    End If
End Sub

' ============================================================================
Public Function GetCriticalErrors() As Collection
    Set GetCriticalErrors = m_criticalErrors
End Function

' ============================================================================
Public Sub ReloadDictionary()
    On Error Resume Next
    m_taskRowMap.RemoveAll
    
    If m_logSheet Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = m_logSheet.Cells(m_logSheet.Rows.count, LogCol.colTaskId).End(xlUp).row
    
    Dim i As Long
    For i = 2 To lastRow
        Dim taskIdValue As Variant
        taskIdValue = m_logSheet.Cells(i, LogCol.colTaskId).Value
        If IsNumeric(taskIdValue) Then
            m_taskRowMap.Add CStr(CLng(taskIdValue)), i
        End If
    Next i
    On Error GoTo 0
End Sub

' ============================================================================
Private Sub Class_Terminate()
    CloseLogFile
    Set m_taskRowMap = Nothing
    Set m_logBuffer = Nothing
    Set m_criticalErrors = Nothing
    Set m_syncLock = Nothing
End Sub


