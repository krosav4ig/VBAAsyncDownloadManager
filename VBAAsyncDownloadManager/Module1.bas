Attribute VB_Name = "Module1"
' ============================================================================
' Модуль: Module1 (стандартный)
' Назначение: Точка входа, инициализация, запуск, обработчик потери состояния
' ============================================================================
'Option Explicit
' ----------------------------------------------------------------------------
' Вспомогательная процедура Sleep (не блокирует события, но приостанавливает поток)
' ----------------------------------------------------------------------------
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Private g_manager As CAsyncDownloadManager
Private g_logger As CDownloadLogger

' ----------------------------------------------------------------------------
' Запуск пакетной загрузки (главная точка входа)
' ----------------------------------------------------------------------------
Sub StartBatchDownload()
    ' 1. Подготовка листа лога
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("Log")
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add
        wsLog.Name = "Log"
    End If
    On Error GoTo 0
    wsLog.Cells.Clear
    
    ' 2. Создаём логгер
    Set g_logger = New CDownloadLogger
    g_logger.Init wsLog, ThisWorkbook.Path & "\download_log.txt"
    g_logger.ReloadDictionary
    
    
    ' 3. Создаём менеджер (максимум 2 одновременных загрузки)
    Set g_manager = New CAsyncDownloadManager
    g_manager.Init maxConcurrent:=3, callback:=g_logger
    
    ' 4. Добавляем задачи (URL, путь сохранения)
    '    При необходимости измените пути на существующие папки на вашем диске
    
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
    
    ' 5. Запускаем загрузки
    g_manager.Start
    
    ' 6. Цикл ожидания с сохранением отзывчивости Excel
    '    Прервать можно нажатием Ctrl+Break (или Esc)
    Do While g_manager.IsBusy
        DoEvents
        Sleep 100   ' небольшая пауза для снижения нагрузки на процессор
    Loop
    
    ' 7. Очистка и завершение
    Set g_manager = Nothing
    Set g_logger = Nothing
    MsgBox "Все загрузки завершены!", vbInformation
End Sub

' ----------------------------------------------------------------------------
' Обработчик потери состояния VBA (вызывается StateLossCallback)
' ----------------------------------------------------------------------------
Public Sub OnWorkerLostState(ByVal taskId As Long, ByVal url As String, ByVal destPath As String)
    ' Этот макрос выполнится даже после нажатия End, Stop или сброса VBA.
    ' Используем только файловый лог, так как объекты Excel могут быть уже разрушены.
    On Error Resume Next
    Dim logPath As String
    logPath = ThisWorkbook.Path & "\critical_failure.log"
    Open logPath For Append As #1
    Print #1, Now & " | КРИТИЧЕСКАЯ ОШИБКА: Задача " & taskId & " потеряла состояние. URL: " & url
    Close #1
    On Error GoTo 0
End Sub