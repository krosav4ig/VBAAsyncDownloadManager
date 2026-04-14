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
' Модуль: CAsyncDownloadWorker.cls
' Назначение: Асинхронная загрузка с буферизацией, таймаутами и проверкой файла
' ============================================================================
Option Explicit

Private WithEvents m_http As WinHttp.WinHttpRequest
Attribute m_http.VB_VarHelpID = -1
Private m_taskId As Long
Private m_url As String
Private m_destPath As String
Private m_callback As IDownloadCallback
Private m_manager As CAsyncDownloadManager
Private m_totalBytes As Currency
Private m_receivedBytes As Currency
Private m_fileWriter As CFileWriter
Private m_lossCallback As StateLossCallback
Private m_startTime As Double
Private m_lastReportedPercent As Integer
Private m_chunkCounter As Long
Private m_isCompleted As Boolean          ' Флаг для предотвращения двойного завершения
Private m_expectedSize As Currency        ' Content-Length (0 если неизвестен)
Private m_bufferSize As Long              ' Размер буфера записи

' Таймауты WinHTTP
Private m_resolveTimeout As Long
Private m_connectTimeout As Long
Private m_sendTimeout As Long
Private m_receiveTimeout As Long

Private Const KB As Currency = 1024
Private Const MB As Currency = KB * 1024      ' 1,048,576
Private Const GB As Currency = MB * 1024      ' 1,073,741,824

Public Event WorkerComplete(ByVal TaskId As Long, ByVal success As Boolean)

' Инициализация
Public Sub Init(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String, _
                ByVal callback As IDownloadCallback, ByVal bufferSize As Long, _
                ByVal resolveTimeout As Long, ByVal connectTimeout As Long, _
                ByVal sendTimeout As Long, ByVal receiveTimeout As Long)
    m_taskId = TaskId
    m_url = url
    m_destPath = destPath
    Set m_callback = callback
    m_bufferSize = bufferSize
    m_resolveTimeout = resolveTimeout
    m_connectTimeout = connectTimeout
    m_sendTimeout = sendTimeout
    m_receiveTimeout = receiveTimeout
    Set m_http = New WinHttp.WinHttpRequest
    m_totalBytes = 0@
    m_receivedBytes = 0@
    m_startTime = Timer
    m_lastReportedPercent = -1
    m_chunkCounter = 0
    m_isCompleted = False
    m_expectedSize = 0@
    
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

' Запуск загрузки
Public Sub Start()
    ' Установка таймаутов
    m_http.SetTimeouts m_resolveTimeout, m_connectTimeout, m_sendTimeout, m_receiveTimeout
    ' Инициализация буферизированной записи
    Set m_fileWriter = New CFileWriter
    If Not m_fileWriter.Init(m_destPath, m_bufferSize) Then
        FailWithError "Не удалось создать файл: " & m_destPath
        Exit Sub
    End If
    
    m_http.Open "GET", m_url, True
    m_http.Send
    
    If Not m_callback Is Nothing Then
        m_callback.OnStart m_taskId, m_url, m_destPath
    End If
End Sub

' Отмена загрузки
Public Sub Cancel()
    If Not m_http Is Nothing Then
        On Error Resume Next
        m_http.abort
        On Error GoTo 0
    End If
    If Not m_lossCallback Is Nothing Then m_lossCallback.Cancel
    If Not m_fileWriter Is Nothing Then m_fileWriter.CloseFile
    m_isCompleted = True
End Sub

' Принудительное успешное завершение (используется при проверке уже существующего файла)
Public Function CheckAndCompleteIfFileExists() As Boolean
    If m_isCompleted Then Exit Function
    If m_expectedSize = 0@ Then Exit Function   ' Неизвестен ожидаемый размер – не проверяем
    
    Dim actualSize As Currency
    actualSize = GetFileSizeByPath(m_destPath)
    If actualSize > 0@ And actualSize = m_expectedSize Then
        ' Файл уже полностью скачан – завершаем задачу успешно
        m_isCompleted = True
        If Not m_callback Is Nothing Then
            m_callback.OnComplete m_taskId, m_url, m_destPath, True, ""
        End If
        If Not m_manager Is Nothing Then
            m_manager.OnWorkerComplete m_taskId, True
        End If
        RaiseEvent WorkerComplete(m_taskId, True)
        CheckAndCompleteIfFileExists = True
    Else
        CheckAndCompleteIfFileExists = False
    End If
End Function

' Вспомогательная функция при ошибке
Private Sub FailWithError(ByVal errorMsg As String)
    m_isCompleted = True
    If Not m_callback Is Nothing Then
        m_callback.OnComplete m_taskId, m_url, m_destPath, False, errorMsg
    End If
    If Not m_manager Is Nothing Then
        m_manager.OnWorkerComplete m_taskId, False
    End If
    RaiseEvent WorkerComplete(m_taskId, False)
    If Not m_fileWriter Is Nothing Then m_fileWriter.CloseFile
End Sub

' Получение Content-Length (безопасно)
Private Function GetContentLength() As Currency
    Dim headerValue As String
    headerValue = m_http.GetResponseHeader("Content-Length")
    If headerValue = "" Then
        GetContentLength = 0@
        Exit Function
    End If
    On Error Resume Next
    GetContentLength = CDec(headerValue)
    If Err.Number <> 0 Then GetContentLength = 0@
    On Error GoTo 0
End Function

' Событие: получены данные
Private Sub m_http_OnResponseDataAvailable(Data() As Byte)
    If m_isCompleted Then Exit Sub
    m_fileWriter.WriteBytes Data
    m_receivedBytes = m_receivedBytes + (UBound(Data) - LBound(Data) + 1)
    
    If m_totalBytes = 0@ Then
        m_totalBytes = GetContentLength
        m_expectedSize = m_totalBytes
    End If
    
    Dim percent As Integer
    If m_totalBytes > 0@ Then
        percent = CInt(CDec(m_receivedBytes) / CDec(m_totalBytes) * 100)
        If percent > 100 Then percent = 100
    Else
        percent = -1
    End If
    
    ' Уведомление о прогрессе с адаптивной частотой
    If Not m_callback Is Nothing Then
        Dim threshold As Long
        If m_totalBytes > 0 Then
            If m_totalBytes < 1 * MB Then          ' менее 1 МБ
                threshold = 100
            ElseIf m_totalBytes < 10 * MB Then     ' от 1 до 10 МБ
                threshold = 300
            Else
                threshold = 800                    ' более 10 МБ
            End If
        Else
            threshold = 200   ' если неизвестный размер
        End If
        m_chunkCounter = m_chunkCounter + 1
        If m_chunkCounter Mod threshold = 0 Then
            m_callback.OnProgress m_taskId, m_receivedBytes, m_totalBytes, percent
        End If
    End If
End Sub

' Событие: загрузка завершена
Private Sub m_http_OnResponseFinished()
    If m_isCompleted Then Exit Sub
    m_isCompleted = True
    ' Закрываем файл и сбрасываем буфер
    m_fileWriter.CloseFile
    
    Dim success As Boolean
    Dim errorMsg As String
    
    If m_http.status = 200 Then
        ' Проверяем размер файла
        Dim actualSize As Currency
        actualSize = m_fileWriter.GetFileSize
        If m_expectedSize > 0@ And actualSize <> m_expectedSize Then
            success = False
            errorMsg = "Несоответствие размера: ожидалось " & Format(m_expectedSize, "#,##0") & " байт, получено " & Format(actualSize, "#,##0")
        Else
            success = True
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

' Событие: ошибка WinHTTP
Private Sub m_http_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    If m_isCompleted Then Exit Sub
    m_isCompleted = True
    m_fileWriter.CloseFile
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

