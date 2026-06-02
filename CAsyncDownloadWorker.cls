VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAsyncDownloadWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' Класс: CAsyncDownloadWorker.cls
' Назначение: Асинхронная загрузка с поддержкой докачки (Resume),
'             буферизированной записью, динамическими таймаутами и проверкой целостности.
' ============================================================================
Option Explicit

Private WithEvents m_http As WinHttp.WinHttpRequest
Attribute m_http.VB_VarHelpID = -1
Private m_taskId As Long
Private m_url As String
Private m_destPath As String
Private m_callback As IDownloadCallback
Private m_manager As CAsyncDownloadManager
Private m_fileWriter As CFileWriter
Private m_lossCallback As StateLossCallback

' Статистика загрузки
Private m_totalBytes As Currency       ' Ожидаемый полный размер файла
Private m_receivedBytes As Currency    ' Сколько уже скачано в этой сессии
Private m_startOffset As Currency      ' Смещение, с которого начали качать (для докачки)
Private m_expectedSize As Currency     ' Полный размер файла (из Content-Length или Content-Range)

' Таймауты и буферизация
Private m_bufferSize As Long
Private m_resolveTimeout As Long
Private m_connectTimeout As Long
Private m_sendTimeout As Long
Private m_receiveTimeout As Long

' Механизм обнаружения зависаний (Dynamic Timeout)
Private m_lastActivityTime As Double
Private m_baseIdleTimeout As Double
Private m_speedHistory() As Double
Private m_speedHistoryIndex As Integer
Private m_speedHistoryCount As Integer
Private m_minSpeedThreshold As Double
Private m_safetyFactor As Double

' Двухфазная проверка (Подозрение -> Рестарт)
Private m_suspectedStuck As Boolean
Private m_suspicionStartTime As Double

' Рестарты
Private m_restartCount As Integer
Private m_maxRestarts As Integer

' Флаги состояния
Private m_isCompleted As Boolean
Private m_chunkCounter As Long
Private m_lastReportedPercent As Integer

Public Event WorkerComplete(ByVal TaskId As Long, ByVal success As Boolean)

' ----------------------------------------------------------------------------
' Инициализация воркера
' ----------------------------------------------------------------------------
Public Sub Init(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String, _
                ByVal callback As IDownloadCallback, ByVal bufferSize As Long, _
                ByVal resolveTimeout As Long, ByVal connectTimeout As Long, _
                ByVal sendTimeout As Long, ByVal receiveTimeout As Long)
    
    m_taskId = TaskId
    m_url = url
    m_destPath = destPath
    Set m_callback = callback
    m_bufferSize = bufferSize
    
    ' Настройка таймаутов
    m_resolveTimeout = resolveTimeout
    m_connectTimeout = connectTimeout
    m_sendTimeout = sendTimeout
    m_receiveTimeout = receiveTimeout
    
    Set m_http = New WinHttp.WinHttpRequest
    
    ' Сброс статистики
    m_totalBytes = 0@
    m_receivedBytes = 0@
    m_startOffset = 0@
    m_expectedSize = 0@
    m_isCompleted = False
    m_restartCount = 0
    m_suspectedStuck = False
    
    ' Настройки детекции зависаний
    m_baseIdleTimeout = 120          ' 2 минуты минимум
    m_minSpeedThreshold = 1024       ' 1 КБ/с порог
    m_safetyFactor = 2#              ' Коэффициент запаса
    ReDim m_speedHistory(1 To 10)    ' История последних 10 замеров
    m_speedHistoryIndex = 0
    m_speedHistoryCount = 0
    
    ' Защита от потери состояния
    Set m_lossCallback = New StateLossCallback
    m_lossCallback.InitByMacroName "OnWorkerLostState", m_taskId, m_url, m_destPath
End Sub

Public Sub SetManager(ByVal manager As CAsyncDownloadManager)
    Set m_manager = manager
End Sub

Public Property Get TaskId() As Long
    TaskId = m_taskId
End Property

' ----------------------------------------------------------------------------
' Запуск загрузки (с поддержкой докачки)
' ----------------------------------------------------------------------------
Public Sub Start()
    On Error GoTo ErrorHandler
    
    ' Установка таймаутов WinHTTP
    m_http.SetTimeouts m_resolveTimeout, m_connectTimeout, m_sendTimeout, m_receiveTimeout
    
    ' Проверка существующего файла для докачки
    Dim existingSize As Currency
    existingSize = GetFileSizeByPath(m_destPath)
    
    m_http.Open "GET", m_url, True
    
    ' Если файл уже частично скачан, запрашиваем остаток
    If existingSize > 0@ Then
        ' Добавляем заголовок Range для докачки
        m_http.SetRequestHeader "Range", "bytes=" & existingSize & "-"
        m_startOffset = existingSize
        ' При использовании Range сервер вернет Content-Range вместо Content-Length
    Else
        m_startOffset = 0@
        ' Очищаем кэш, чтобы избежать старых данных
        m_http.SetRequestHeader "Cache-Control", "no-cache"
        m_http.SetRequestHeader "Pragma", "no-cache"
    End If
    
    ' Инициализация буферизированной записи
    ' Важно: режим Append реализуется внутри CFileWriter через Seek
    Set m_fileWriter = New CFileWriter
    If Not m_fileWriter.Init(m_destPath, m_bufferSize) Then
        FailWithError "Не удалось создать/открыть файл: " & m_destPath
        Exit Sub
    End If
    
    m_http.Send
    
    ' Уведомление о старте (если это первая попытка, а не рестарт)
    If m_restartCount = 0 And Not m_callback Is Nothing Then
        m_callback.OnStart m_taskId, m_url, m_destPath
    End If
    
    m_lastActivityTime = Timer
    Exit Sub

ErrorHandler:
    FailWithError "Ошибка запуска: " & Err.Description
End Sub

' ----------------------------------------------------------------------------
' Отмена загрузки
' ----------------------------------------------------------------------------
Public Sub Cancel()
    If Not m_http Is Nothing Then
        On Error Resume Next
        m_http.Abort
        On Error GoTo 0
    End If
    If Not m_lossCallback Is Nothing Then m_lossCallback.Cancel
    If Not m_fileWriter Is Nothing Then m_fileWriter.CloseFile
    m_isCompleted = True
End Sub

' ----------------------------------------------------------------------------
' Проверка: файл уже скачан полностью? (Используется менеджером)
' ----------------------------------------------------------------------------
Public Function CheckAndCompleteIfFileExists() As Boolean
    If m_isCompleted Then Exit Function
    
    ' Получаем полный ожидаемый размер (если он уже известен из заголовков)
    ' Если еще неизвестен (загрузка не началась), проверяем просто наличие файла > 0
    ' Но надежнее проверять только если мы знаем Content-Length
    
    If m_expectedSize > 0@ Then
        Dim actualSize As Currency
        actualSize = GetFileSizeByPath(m_destPath)
        
        If actualSize >= m_expectedSize Then
            ' Файл совпадает или больше ожидаемого (на всякий случай обрезаем или считаем успехом)
            m_isCompleted = True
            If Not m_fileWriter Is Nothing Then m_fileWriter.CloseFile
            
            If Not m_callback Is Nothing Then
                m_callback.OnComplete m_taskId, m_url, m_destPath, True, ""
            End If
            If Not m_manager Is Nothing Then
                m_manager.OnWorkerComplete m_taskId, True
            End If
            RaiseEvent WorkerComplete(m_taskId, True)
            CheckAndCompleteIfFileExists = True
        End If
    End If
End Function

' ----------------------------------------------------------------------------
' Принудительный рестарт задачи (с докачкой)
' ----------------------------------------------------------------------------
Public Sub Restart()
    If m_restartCount >= m_maxRestarts Then
        FailWithError "Превышено количество попыток рестарта (" & m_maxRestarts & ")"
        Exit Sub
    End If
    
    m_restartCount = m_restartCount + 1
    
    ' Закрываем текущий файл и поток, но НЕ удаляем файл!
    ' Мы хотим продолжить запись в него.
    If Not m_fileWriter Is Nothing Then
        m_fileWriter.CloseFile
        Set m_fileWriter = Nothing
    End If
    
    ' Сброс флагов сессии
    m_receivedBytes = 0@
    m_isCompleted = False
    m_suspectedStuck = False
    m_suspicionStartTime = 0
    m_speedHistoryCount = 0
    
    ' Пересоздаем HTTP объект
    Set m_http = New WinHttp.WinHttpRequest
    
    If Not m_callback Is Nothing Then
        ' Можно добавить логирование рестарта
        ' m_callback.OnStatusUpdate ... (если есть такой метод)
    End If
    
    ' Запускаем заново. Метод Start() сам увидит размер файла и добавит Range.
    Start
End Sub

' ----------------------------------------------------------------------------
' Вспомогательные методы
' ----------------------------------------------------------------------------
Private Sub FailWithError(ByVal errorMsg As String)
    m_isCompleted = True
    If Not m_fileWriter Is Nothing Then m_fileWriter.CloseFile
    
    If Not m_callback Is Nothing Then
        m_callback.OnComplete m_taskId, m_url, m_destPath, False, errorMsg
    End If
    If Not m_manager Is Nothing Then
        m_manager.OnWorkerComplete m_taskId, False
    End If
    RaiseEvent WorkerComplete(m_taskId, False)
End Sub

' ----------------------------------------------------------------------------
' Расчет динамического таймаута
' ----------------------------------------------------------------------------
Private Function CalculateDynamicTimeout(currentSpeed As Double) As Double
    Dim timeout As Double
    timeout = m_baseIdleTimeout
    
    If currentSpeed > 0 Then
        Dim timePerBuffer As Double
        timePerBuffer = m_bufferSize / currentSpeed
        
        ' Динамический расчет: время на 3 буфера + половина базового запаса
        Dim calculated As Double
        calculated = (timePerBuffer * 3) + (m_baseIdleTimeout * 0.5)
        calculated = calculated * m_safetyFactor
        
        ' Не может быть меньше базового
        If calculated > m_baseIdleTimeout Then
            timeout = calculated
        Else
            timeout = m_baseIdleTimeout
        End If
        
        ' Ограничитель сверху (1 час)
        If timeout > 3600 Then timeout = 3600
    End If
    
    CalculateDynamicTimeout = timeout
End Function

Private Sub RecordSpeedSample(bytes As Currency, timeElapsed As Currency)
    If timeElapsed <= 0 Then Exit Sub
    Dim speed As Double
    speed = bytes / timeElapsed
    
    m_speedHistoryIndex = m_speedHistoryIndex + 1
    If m_speedHistoryIndex > UBound(m_speedHistory) Then m_speedHistoryIndex = 1
    
    m_speedHistory(m_speedHistoryIndex) = speed
    If m_speedHistoryCount < UBound(m_speedHistory) Then m_speedHistoryCount = m_speedHistoryCount + 1
End Sub

Private Function GetAverageSpeed() As Double
    If m_speedHistoryCount = 0 Then
        GetAverageSpeed = 0
        Exit Function
    End If
    
    Dim sum As Double
    Dim i As Integer
    For i = 1 To m_speedHistoryCount
        sum = sum + m_speedHistory(i)
    Next i
    GetAverageSpeed = sum / m_speedHistoryCount
End Function

' ----------------------------------------------------------------------------
' Проверка на зависание (IsStuck)
' ----------------------------------------------------------------------------
Public Function IsStuck() As Boolean
    If Not m_isCompleted And m_http Is Nothing Then
        IsStuck = False ' Еще не запущен
        Exit Function
    End If
    
    Dim currentTime As Double
    currentTime = Timer
    If currentTime < m_lastActivityTime Then currentTime = currentTime + 86400 ' Переход через полночь
    
    Dim idleTime As Double
    idleTime = currentTime - m_lastActivityTime
    
    Dim avgSpeed As Double
    avgSpeed = GetAverageSpeed()
    
    ' 1. Проверка низкой скорости (Двухфазная)
    If avgSpeed > 0 And avgSpeed < m_minSpeedThreshold Then
        If Not m_suspectedStuck Then
            m_suspectedStuck = True
            m_suspicionStartTime = currentTime
        Else
            Dim suspicionDuration As Double
            suspicionDuration = currentTime - m_suspicionStartTime
            
            ' Динамическое ожидание при низкой скорости
            Dim dynamicWaitTimeout As Double
            If avgSpeed > 0 Then
                dynamicWaitTimeout = (m_bufferSize / avgSpeed) * m_safetyFactor
                If dynamicWaitTimeout < 10 Then dynamicWaitTimeout = 10
                If dynamicWaitTimeout > 300 Then dynamicWaitTimeout = 300
            Else
                dynamicWaitTimeout = 60
            End If
            
            If suspicionDuration > dynamicWaitTimeout Then
                IsStuck = True
                Exit Function
            End If
        End If
    ElseIf avgSpeed = 0 And idleTime > 5 Then
        ' Нет данных и нет активности
        If Not m_suspectedStuck Then
            m_suspectedStuck = True
            m_suspicionStartTime = currentTime
        Else
            If (currentTime - m_suspicionStartTime) > 30 Then
                IsStuck = True
                Exit Function
            End If
        End If
    Else
        ' Скорость в норме
        If m_suspectedStuck Then
            m_suspectedStuck = False
            m_suspicionStartTime = 0
        End If
    End If
    
    ' 2. Проверка абсолютного простоя (Динамический таймаут)
    Dim dynamicTimeout As Double
    dynamicTimeout = CalculateDynamicTimeout(avgSpeed)
    
    If idleTime > dynamicTimeout Then
        IsStuck = True
        Exit Function
    End If
    
    IsStuck = False
End Function

' ----------------------------------------------------------------------------
' СОБЫТИЯ WINHTTP
' ----------------------------------------------------------------------------

' Событие: Получены данные
Private Sub m_http_OnResponseDataAvailable(arrData() As Byte)
    If m_isCompleted Then Exit Sub
    
    ' Запись в буферизированный файл
    m_fileWriter.WriteBytes arrData
    
    Dim chunkSize As Currency
    chunkSize = UBound(arrData) - LBound(arrData) + 1
    m_receivedBytes = m_receivedBytes + chunkSize
    
    ' Обновление времени активности
    m_lastActivityTime = Timer
    
    ' Сброс флага зависания при получении данных
    If m_suspectedStuck Then
        m_suspectedStuck = False
        m_suspicionStartTime = 0
    End If
    
    ' Запись скорости в историю (эмуляция интервала 0.5 сек для простоты,
    ' в идеале нужно замерять реальное время между вызовами)
    RecordSpeedSample chunkSize, 0.5
    
    ' Определение общего размера файла (Content-Length или Content-Range)
    If m_totalBytes = 0@ Then
        ' Пытаемся получить Content-Range (если была докачка)
        ' Формат: bytes 100-200/1000 -> общий размер 1000
        Dim rangeHeader As String
        On Error Resume Next
        rangeHeader = m_http.GetResponseHeader("Content-Range")
        On Error GoTo 0
        
        If rangeHeader <> "" Then
            ' Парсим "bytes start-end/total"
            Dim parts() As String
            parts = Split(rangeHeader, "/")
            If UBound(parts) >= 1 Then
                On Error Resume Next
                m_totalBytes = CDec(parts(1))
                On Error GoTo 0
            End If
            m_expectedSize = m_totalBytes
        Else
            ' Если Content-Range нет, берем Content-Length (обычный GET)
            ' И добавляем смещение, если это докачка (хотя при обычном GET смещения быть не должно)
            Dim lenHeader As String
            On Error Resume Next
            lenHeader = m_http.GetResponseHeader("Content-Length")
            On Error GoTo 0
            
            If lenHeader <> "" Then
                On Error Resume Next
                m_totalBytes = CDec(lenHeader) + m_startOffset
                On Error GoTo 0
                m_expectedSize = m_totalBytes
            End If
        End If
    End If
    
    ' Расчет процента
    Dim percent As Integer
    If m_totalBytes > 0@ Then
        percent = CInt((m_startOffset + m_receivedBytes) / m_totalBytes * 100)
        If percent > 100 Then percent = 100
    Else
        percent = -1
    End If
    
    ' Уведомление о прогрессе (адаптивное)
    If Not m_callback Is Nothing Then
        Dim threshold As Long
        ' Частота обновления зависит от размера файла
        If m_totalBytes < 1048576 Then
            threshold = 100       ' < 1MB
        ElseIf m_totalBytes < 10485760 Then
            threshold = 300   ' < 10MB
        Else:
            threshold = 800   ' > 10MB
        End If
        
        m_chunkCounter = m_chunkCounter + 1
        If m_chunkCounter Mod threshold = 0 Or percent = 100 Then
            m_callback.OnProgress m_taskId, m_startOffset + m_receivedBytes, m_totalBytes, percent
        End If
    End If
End Sub

' Событие: Загрузка завершена
Private Sub m_http_OnResponseFinished()
    If m_isCompleted Then Exit Sub
    m_isCompleted = True
    
    ' Закрываем файл (сброс буфера)
    If Not m_fileWriter Is Nothing Then
        m_fileWriter.CloseFile
        Set m_fileWriter = Nothing
    End If
    
    Dim success As Boolean
    Dim errorMsg As String
    
    ' Проверка статуса HTTP
    ' 200 OK (полный файл) или 206 Partial Content (докачка) - оба успешны
    If m_http.status = 200 Or m_http.status = 206 Then
        Dim actualSize As Currency
        actualSize = GetFileSizeByPath(m_destPath)
        
        ' Если мы знаем ожидаемый размер, проверяем совпадение
        If m_expectedSize > 0@ Then
            If actualSize = m_expectedSize Then
                success = True
            Else
                ' Допускаем небольшую погрешность или ситуацию, когда сервер не отдал длину
                ' Но строго говоря, если длина известна, она должна совпадать
                ' Если actualSize > expectedSize (редко), тоже можно считать успехом, но лучше проверить
                If actualSize > m_expectedSize Then
                     success = True ' Файл больше, чем ожидалось (сервер изменил контент?), примем.
                Else
                     success = False
                     errorMsg = "Несоответствие размера: ожидалось " & Format(m_expectedSize, "#,##0") & _
                                " байт, получено " & Format(actualSize, "#,##0")
                End If
            End If
        Else
            ' Если размер неизвестен, считаем успехом, если файл > 0 и нет ошибок HTTP
            If actualSize > 0 Then success = True Else success = False
            If Not success Then errorMsg = "Файл пуст или не создан."
        End If
    Else
        success = False
        errorMsg = "HTTP " & m_http.status & " " & m_http.StatusText
    End If
    
    If Not m_callback Is Nothing Then
        m_callback.OnComplete m_taskId, m_url, m_destPath, success, errorMsg
    End If
    If Not m_manager Is Nothing Then
        m_manager.OnWorkerComplete m_taskId, success
    End If
    RaiseEvent WorkerComplete(m_taskId, success)
End Sub

' Событие: Ошибка WinHTTP
Private Sub m_http_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    If m_isCompleted Then Exit Sub
    m_isCompleted = True
    If Not m_fileWriter Is Nothing Then
        m_fileWriter.CloseFile
        Set m_fileWriter = Nothing
    End If
    
    Dim errorMsg As String
    errorMsg = "WinHTTP ошибка " & ErrorNumber & ": " & ErrorDescription
    
    If Not m_callback Is Nothing Then
        m_callback.OnComplete m_taskId, m_url, m_destPath, False, errorMsg
    End If
    If Not m_manager Is Nothing Then
        m_manager.OnWorkerComplete m_taskId, False
    End If
    RaiseEvent WorkerComplete(m_taskId, False)
End Sub

' Деструктор
Private Sub Class_Terminate()
    If Not m_fileWriter Is Nothing Then m_fileWriter.CloseFile
    If Not m_lossCallback Is Nothing Then m_lossCallback.Cancel
End Sub



