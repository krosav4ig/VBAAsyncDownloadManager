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
' Назначение: Асинхронная загрузка одного файла через WinHTTP с событиями
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
Private m_stream As Object          ' ADODB.Stream
Private m_lossCallback As StateLossCallback   ' защита от потери состояния
Private m_startTime As Double
Private m_lastReportedPercent As Integer   ' Для оптимизации событий прогресса
Private m_chunkCounter As Long              ' Счётчик фрагментов для пропуска
Private Const MAX_LONG As Long = 2147483647   ' 2^31 - 1
Private Const KB As Currency = 1024
Private Const MB As Currency = KB * KB      ' 1,048,576
Private Const GB As Currency = MB * KB      ' 1,073,741,824


Public Event WorkerComplete(ByVal taskId As Long, ByVal success As Boolean)

' Инициализация рабочего процесса
Public Sub Init(ByVal taskId As Long, ByVal url As String, ByVal destPath As String, _
                ByVal callback As IDownloadCallback)
    m_taskId = taskId
    m_url = url
    m_destPath = destPath
    Set m_callback = callback
    Set m_http = New WinHttp.WinHttpRequest
    m_totalBytes = 0
    m_receivedBytes = 0
    m_startTime = Timer
    m_lastReportedPercent = -1
    m_chunkCounter = 0

    
    ' Создаём защитный колбэк, который вызовет макрос даже при сбросе VBA
    Set m_lossCallback = New StateLossCallback
    m_lossCallback.InitByMacroName "OnWorkerLostState", m_taskId, m_url, m_destPath
End Sub

' Устанавливает ссылку на менеджер (для уведомлений о завершении)
Public Sub SetManager(ByVal manager As CAsyncDownloadManager)
    Set m_manager = manager
End Sub

' Запуск асинхронной загрузки
Public Sub Start()
    ' Создаём бинарный поток для сохранения данных
    Set m_stream = CreateObject("ADODB.Stream")
    m_stream.Type = 1   ' adTypeBinary
    m_stream.Open
    
    ' Асинхронный GET-запрос
    m_http.Open "GET", m_url, True
    m_http.Send
    
    ' Уведомляем callback о начале загрузки
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
    If Not m_lossCallback Is Nothing Then
        m_lossCallback.Cancel
    End If
End Sub

' Освобождение ресурсов
Private Sub Terminate()
    If Not m_stream Is Nothing Then
        On Error Resume Next
        m_stream.Close
        Set m_stream = Nothing
        On Error GoTo 0
    End If
    If Not m_lossCallback Is Nothing Then
        On Error Resume Next
        m_lossCallback.Cancel
        On Error GoTo 0
    End If
End Sub

' Безопасное получение Content-Length (без переполнения)
Private Function GetContentLength() As Currency
    Dim headerValue As String
    headerValue = m_http.GetResponseHeader("Content-Length")
    
    If headerValue = "" Then
        GetContentLength = 0@
        Exit Function
    End If
    
    ' Используем CDec для преобразования больших чисел
    On Error Resume Next
    GetContentLength = CDec(headerValue)
    If Err.Number <> 0 Then
        GetContentLength = 0@
        Err.Clear
    End If
    On Error GoTo 0
End Function

' Событие: получена порция данных
Private Sub m_http_OnResponseDataAvailable(Data() As Byte)
    ' Записываем данные в поток
    m_stream.Write Data
    m_receivedBytes = m_receivedBytes + (UBound(Data) - LBound(Data) + 1)
    
    ' Пытаемся получить общий размер из заголовка Content-Length (если ещё не знаем)
    If m_totalBytes = 0 Then
        m_totalBytes = GetContentLength
    End If
    
     ' Вычисляем процент (безопасно, используя Currency)
    Dim percent As Integer
    If m_totalBytes > 0@ Then
        ' Используем CDec для промежуточных вычислений
        Dim ratio As Variant
        ratio = CDec(m_receivedBytes) / CDec(m_totalBytes)
        percent = CInt(ratio * 100)
        If percent > 100 Then percent = 100
    Else
        percent = -1
    End If
    
    ' БЕЗОПАСНОЕ ОБНОВЛЕНИЕ СЧЁТЧИКА (циклический сброс)
    If m_chunkCounter >= MAX_LONG Then
        m_chunkCounter = 0
    Else
        m_chunkCounter = m_chunkCounter + 1
    End If
    
    ' Уведомляем callback о прогрессе
    If Not m_callback Is Nothing Then
        ' Для маленьких файлов обновляем чаще, для больших - реже
        Dim updateThreshold As Integer
        If m_totalBytes < 1 * MB Then  ' Менее 1MB
            updateThreshold = 100   ' каждые 100 фрагментов
        ElseIf m_totalBytes < 10 * MB Then  ' Менее 100MB
            updateThreshold = 300  ' каждые 300 фрагментов
        Else
            updateThreshold = 800  ' каждые 800 фрагментов
        End If
        
        If m_chunkCounter Mod updateThreshold = 0 Then
            m_callback.OnProgress m_taskId, m_receivedBytes, m_totalBytes, percent
        End If
    End If
End Sub

' Событие: загрузка завершена (все данные получены)
Private Sub m_http_OnResponseFinished()
    Dim success As Boolean
    Dim errorMsg As String
    Dim fso As Object
    
    ' Проверяем HTTP статус
    If m_http.status = 200 Then
        On Error Resume Next
        ' Создаём папку назначения, если её нет
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim folder As String
        folder = fso.GetParentFolderName(m_destPath)
        If folder <> "" And Not fso.FolderExists(folder) Then
            fso.CreateFolder folder
        End If
        ' Сохраняем поток в файл
        m_stream.SaveToFile m_destPath, 2   ' adSaveCreateOverWrite
        If Err.Number = 0 Then
            success = True
        Else
            success = False
            errorMsg = "Ошибка сохранения файла: " & Err.Description
        End If
        On Error GoTo 0
    Else
        success = False
        errorMsg = "HTTP " & m_http.status & " " & m_http.StatusText
    End If
    
    ' Освобождаем ресурсы
    Terminate
    
    ' Уведомляем callback о завершении
    If Not m_callback Is Nothing Then
        m_callback.OnComplete m_taskId, m_url, m_destPath, success, errorMsg
    End If
    
    ' Уведомляем менеджера, чтобы запустить следующую задачу
    If Not m_manager Is Nothing Then
        m_manager.OnWorkerComplete m_taskId, success
    End If
    RaiseEvent WorkerComplete(m_taskId, success)
End Sub

' Событие: ошибка WinHTTP
Private Sub m_http_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    Terminate
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

' Деструктор класса
Private Sub Class_Terminate()
    Terminate
End Sub


