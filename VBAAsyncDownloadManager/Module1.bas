Attribute VB_Name = "Module1"
' ============================================================================
' Модуль: Module1.bas (стандартный)
' Назначение: Точка входа – добавлены параметры таймаутов и буфера
' ============================================================================
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private g_manager As CAsyncDownloadManager
Private g_logger As CDownloadLogger

' Глобальная процедура для таймера (должна быть в стандартном модуле)
Public Sub CAsyncDownloadManager_TimerCheck()
    If Not g_manager Is Nothing Then
        g_manager.TimerCheck
    End If
End Sub

Sub StartBatchDownload()
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("Log")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add
        wsLog.Name = "Log"
    End If
    On Error GoTo 0
    wsLog.Cells.Clear
    
    Set g_logger = New CDownloadLogger
    g_logger.Init wsLog, ThisWorkbook.Path & "\download_log.txt"
    g_logger.ReloadDictionary
    
    ' Параметры: макс. параллельно, колбэк, таймауты (мс), буфер (байт), интервал проверки (сек)
    Set g_manager = New CAsyncDownloadManager
    g_manager.Init maxConcurrent:=3, callback:=g_logger, _
                   resolveTimeoutMs:=5000, connectTimeoutMs:=60000, _
                   sendTimeoutMs:=30000, receiveTimeoutMs:=30000, _
                   bufferSizeBytes:=65536, checkIntervalSec:=5
    
    Dim col1&, col2&, col3&, arr() As Variant, i&
    With Лист1.ListObjects!Ссылки
        col1 = .ListColumns!Ссылка.Index
        col2 = .ListColumns![Путь для сохранения].Index
        col3 = .ListColumns!Скачано.Index
        arr = .DataBodyRange
        For i = 1 To .ListRows.Count
            g_manager.AddTask arr(i, col1), arr(i, col2)
        Next
    End With
    
    g_manager.Start
    
    Do While g_manager.IsBusy
        DoEvents
        Sleep 100
    Loop
    
    Set g_manager = Nothing
    Set g_logger = Nothing
    MsgBox "Все загрузки завершены!", vbInformation
End Sub

Public Sub OnWorkerLostState(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String)
    On Error Resume Next
    Dim logPath As String
    logPath = ThisWorkbook.Path & "\critical_failure.log"
    Open logPath For Append As #1
    Print #1, Now & " | КРИТИЧЕСКАЯ ОШИБКА: Задача " & TaskId & " потеряла состояние. URL: " & url
    Close #1
    On Error GoTo 0
End Sub

