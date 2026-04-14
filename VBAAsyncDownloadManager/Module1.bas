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
    
    ' 4. НАСТРОЙКА ТАЙМАУТОВ (в миллисекундах)
    '    Для избежания ошибки -2147012894 (таймаут соединения) рекомендуется:
    '    - resolveTimeout: время на разрешение DNS (60000 = 60 сек)
    '    - connectTimeout: время на подключение к серверу (120000 = 120 сек)
    '    - sendTimeout: время на отправку запроса (60000 = 60 сек)
    '    - receiveTimeout: время на получение данных (300000 = 5 мин для больших файлов)
    g_manager.SetDefaultTimeouts _
        resolveTimeout:=60000, _
        connectTimeout:=120000, _
        sendTimeout:=60000, _
        receiveTimeout:=300000
    
    ' 5. НАСТРОЙКА БУФЕРИЗИРОВАННОЙ ЗАПИСИ
    '    useBufferedWrite:=True - разбивает сохранение на части, не блокируя события WinHTTP
    '    bufferSize:=65536 - размер буфера в байтах (64 KB по умолчанию)
    '    Для очень больших файлов можно увеличить до 256*1024 или 512*1024
    g_manager.SetDefaultBufferedWriteOptions _
        useBufferedWrite:=True, _
        bufferSize:=65536
    
    ' 6. Добавляем задачи (URL, путь сохранения)
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
    
    ' 7. Запускаем загрузки
    g_manager.Start
    
    ' 8. Цикл ожидания с сохранением отзывчивости Excel
    '    Прервать можно нажатием Ctrl+Break (или Esc)
    Do While g_manager.IsBusy
        DoEvents
        Sleep 100   ' небольшая пауза для снижения нагрузки на процессор
    Loop
    
    ' 9. Очистка и завершение
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