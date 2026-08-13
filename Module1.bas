Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const SOURCE_SHEET_NAME As String = "Лист1" 
Private Const SOURCE_TABLE_NAME As String = "Ссылки"
Private Const URL_COLUMN_NAME As String = "Ссылка"
Private Const PATH_COLUMN_NAME As String = "Путь для сохранения"

Public g_manager As CAsyncDownloadManager
Public g_logger As CDownloadLogger
Public g_managerEvents As CManagerEvents
Public g_criticalErrorLogPath As String

' ============================================================================
' Глобальный таймер (вызывается менеджером через Application.OnTime)
' ============================================================================
Public Sub CAsyncDownloadManager_TimerCheck()
    On Error GoTo ErrorHandler
    
    If Not g_manager Is Nothing Then
        g_manager.TimerCheck
    End If
    Exit Sub
ErrorHandler:
    On Error Resume Next
    Dim logFile As Integer: logFile = FreeFile
    Open g_criticalErrorLogPath For Append As #logFile
    Print #logFile, Now & " [CRITICAL] TimerCheck error: " & Err.Description
    Close #logFile
    On Error GoTo 0
End Sub

Sub DownloadCSV()
    StartBatchDownload 4, "ConvertXlsxToCsv_Async"
End Sub
' ============================================================================
' Основная процедура запуска
' ============================================================================
Public Sub StartBatchDownload(Optional ByVal logLevel As Long = LL_INFO, _
                              Optional ByVal postProcessMacroName As String = "", Optional maxConcurrent = 3)
    On Error GoTo CriticalError
    
    g_criticalErrorLogPath = ThisWorkbook.Path & "\critical_errors.log"
    
    Dim wsLog As Worksheet
    Set wsLog = GetOrCreateLogSheet()
    
    Set g_logger = New CDownloadLogger
    g_logger.Init wsLog, ThisWorkbook.Path & "\download_log.txt", logLevel
    g_logger.ReloadDictionary
    
    g_logger.LogMessage LL_INFO, 0, "=== НАЧАЛО СЕССИИ ЗАГРУЗКИ ==="
    
    ' Инициализация постобработки
    InitPostProcess
    
    Set g_manager = New CAsyncDownloadManager
    g_manager.Init maxConcurrent:=maxConcurrent, callback:=g_logger, _
                   resolveTimeoutMs:=30000, connectTimeoutMs:=60000, _
                   sendTimeoutMs:=30000, receiveTimeoutMs:=60000, _
                   bufferSizeBytes:=131072, checkIntervalSec:=3, _
                   maxRetries:=3, _
                   maxNetworkUsagePercent:=60, _
                   networkInterfaceFilter:="", _
                   postProcessMacroName:=postProcessMacroName
    
    HookManagerEvents
    
    Dim taskCount As Long
    taskCount = LoadTasksFromTable()
    
    If taskCount = 0 Then
        g_logger.LogMessage LL_WARNING, 0, "Нет задач для загрузки"
        MsgBox "Нет задач для загрузки. Проверьте таблицу '" & SOURCE_TABLE_NAME & "'.", vbExclamation
        Cleanup
        Exit Sub
    End If
    
    g_logger.LogMessage LL_INFO, 0, "Загружено задач: " & taskCount
    
    g_manager.Start
    
    Dim lastStatusTime As Date
    lastStatusTime = Now
    
    Do While g_manager.IsBusy
        DoEvents
        Sleep 50
        
        If DateDiff("s", lastStatusTime, Now) >= 5 Then
            lastStatusTime = Now
            g_manager.TimerCheck
            CheckAsyncMarkers
        End If
        
    Loop
    
    Dim resultMsg As String
    resultMsg = "Все загрузки и постобработка завершены!" & vbNewLine & _
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
    MsgBox "КРИТИЧЕСКАЯ ОШИБКА: " & Err.Description & vbNewLine & _
           "Код: " & Err.Number, vbCritical, "Критическая ошибка"
    On Error Resume Next
    Dim critFile As Integer: critFile = FreeFile
    Open g_criticalErrorLogPath For Append As #critFile
    Print #critFile, Now & " " & Err.Description
    Close #critFile
    On Error GoTo 0
    Cleanup
End Sub

Private Function GetOrCreateLogSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateLogSheet = ThisWorkbook.Sheets("DownloadLog")
    If GetOrCreateLogSheet Is Nothing Then
        Set GetOrCreateLogSheet = ThisWorkbook.Sheets.Add
        GetOrCreateLogSheet.Name = "DownloadLog"
    End If
    On Error GoTo 0
End Function

Private Function FindSourceSheet() As Worksheet
    Dim ws As Worksheet
    Dim namesToTry As Variant
    namesToTry = Array(SOURCE_SHEET_NAME, "Sheet1", "Лист1")
    
    Dim i As Long
    For i = LBound(namesToTry) To UBound(namesToTry)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(namesToTry(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Set FindSourceSheet = ws
            Exit Function
        End If
    Next i
    Set FindSourceSheet = Nothing
End Function

Private Function LoadTasksFromTable() As Long
    Dim tbl As ListObject
    Dim urlCol As Long, pathCol As Long
    Dim dataRange As Range
    Dim i As Long, count As Long
    
    Dim ws As Worksheet
    Set ws = FindSourceSheet()
    
    If ws Is Nothing Then
        LoadTasksFromTable = 0
        Exit Function
    End If
    
    On Error Resume Next
    Set tbl = ws.ListObjects(SOURCE_TABLE_NAME)
    If tbl Is Nothing Then
        LoadTasksFromTable = 0
        Exit Function
    End If
    
    urlCol = GetColumnIndex(tbl, URL_COLUMN_NAME)
    pathCol = GetColumnIndex(tbl, PATH_COLUMN_NAME)
    
    If urlCol = 0 Or pathCol = 0 Then
        LoadTasksFromTable = 0
        Exit Function
    End If
    
    Set dataRange = tbl.DataBodyRange
    If dataRange Is Nothing Then
        LoadTasksFromTable = 0
        Exit Function
    End If
    
    For i = 1 To dataRange.Rows.count
        Dim url As String, destPath As String
        url = Trim$(CStr(dataRange.Cells(i, urlCol).Value))
        destPath = Trim$(CStr(dataRange.Cells(i, pathCol).Value))
        
        If url <> "" And destPath <> "" Then
            g_manager.AddTask url, destPath
            count = count + 1
        End If
    Next i
    
    LoadTasksFromTable = count
    On Error GoTo 0
End Function

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

Private Sub HookManagerEvents()
    Set g_managerEvents = New CManagerEvents
    g_managerEvents.Init g_manager, g_logger
End Sub

Private Sub Cleanup()
    On Error Resume Next
    
    If Not g_manager Is Nothing Then
        g_manager.StopAll
        Set g_manager = Nothing
    End If
    
    If Not g_managerEvents Is Nothing Then
        Set g_managerEvents = Nothing
    End If
    
    ClearAsyncMarkers
    
    If Not g_logger Is Nothing Then
        g_logger.CloseLogFile
        Set g_logger = Nothing
    End If
    
    On Error GoTo 0
End Sub

Public Sub OnWorkerLostState(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String)
    On Error Resume Next
    Dim logMsg As String
    logMsg = Now & " [STATE_LOSS] Worker #" & TaskId & " lost state. URL: " & url
    
    Dim critFile As Integer: critFile = FreeFile
    Open g_criticalErrorLogPath For Append As #critFile
    Print #critFile, logMsg
    Close #critFile
    
    If Not g_manager Is Nothing Then
        g_manager.AddTask url, destPath
    End If
    On Error GoTo 0
End Sub

Public Sub StopAllDownloads()
    If Not g_manager Is Nothing Then
        g_manager.StopAll
        Debug.Print "Все загрузки остановлены пользователем"
    End If
End Sub

Public Sub PauseDownloads()
    If Not g_manager Is Nothing Then
        g_manager.Pause
        Debug.Print "Загрузки приостановлены"
    End If
End Sub

Public Sub Resume_Downloads()
    If Not g_manager Is Nothing Then
        g_manager.Resume_
        Debug.Print "Загрузки возобновлены"
    End If
End Sub

' ============================================================================
' ГЛОБАЛЬНАЯ ФУНКЦИЯ ПОСТОБРАБОТКИ (вызывается через StateLossCallback)
' ============================================================================
Public Sub ExecutePostProcess(ByVal TaskId As Long, ByVal filePath As String, ByVal macroName As String)
    On Error GoTo ErrorHandler
    
    If Dir(filePath) = "" Then
        If Not g_logger Is Nothing Then
            g_logger.LogMessage LL_WARNING, TaskId, "Файл исчез перед постобработкой: " & filePath
        End If
        If Not g_manager Is Nothing Then g_manager.FinalizeTask TaskId, False
        Exit Sub
    End If
    
    If Not g_logger Is Nothing Then
        g_logger.LogMessage LL_INFO, TaskId, "Запуск постобработки: " & macroName & "(" & filePath & ")"
    End If
    
    ' Устанавливаем контекст для асинхронных операций
    g_currentPostProcessTaskId = TaskId
    
    ' Вызов пользовательского макроса
    Dim result As Variant
    result = Application.Run(macroName, filePath)
    
    ' Если макрос вернул 1 - операция асинхронная, FinalizeTask вызовется позже
    If IsNumeric(result) And CLng(result) = 1 Then
        If Not g_logger Is Nothing Then
            g_logger.LogMessage LL_DEBUG, TaskId, "Асинхронная постобработка запущена, ожидание маркера"
        End If
    Else
        ' Синхронная операция - сразу финализируем
        If Not g_manager Is Nothing Then
            g_manager.FinalizeTask TaskId, True
        End If
    End If
    Exit Sub
    
ErrorHandler:
    If Not g_logger Is Nothing Then
        g_logger.LogMessage LL_ERROR, TaskId, "Ошибка в макросе '" & macroName & "': " & Err.Description, Err.Number
    End If
    
    If Not g_manager Is Nothing Then
        g_manager.FinalizeTask TaskId, False
    End If
End Sub

