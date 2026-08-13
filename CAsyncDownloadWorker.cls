VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAsyncDownloadWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private m_receivedBytes As Currency
Private m_startOffset As Currency
Private m_expectedSize As Currency

Private m_resolveTimeout As Long
Private m_connectTimeout As Long
Private m_sendTimeout As Long
Private m_receiveTimeout As Long
Private m_bufferSize As Long

Private m_lastActivityTime As Date
Private m_maxRestarts As Integer
Private m_restartCount As Integer

Private m_isCompleted As Boolean
Private m_isCancelled As Boolean
Private m_isHeadChecked As Boolean
Private m_lastReportedPercent As Integer
Private m_postProcessMacroName As String

Private Const LL_CRITICAL As Long = 1
Private Const LL_ERROR As Long = 2
Private Const LL_WARNING As Long = 3
Private Const LL_INFO As Long = 4
Private Const LL_DEBUG As Long = 5

Public Event WorkerComplete(ByVal TaskId As Long, ByVal success As Boolean)

Public Sub Init(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String, _
                ByVal callback As IDownloadCallback, ByVal bufferSize As Long, _
                ByVal resolveTimeout As Long, ByVal connectTimeout As Long, _
                ByVal sendTimeout As Long, ByVal receiveTimeout As Long, _
                Optional ByVal maxRetries As Integer = 3, _
                Optional ByVal postProcessMacroName As String = "")
    
    m_taskId = TaskId
    m_url = url
    m_destPath = destPath
    Set m_callback = callback
    m_bufferSize = bufferSize
    m_resolveTimeout = resolveTimeout
    m_connectTimeout = connectTimeout
    m_sendTimeout = sendTimeout
    m_receiveTimeout = receiveTimeout
    m_maxRestarts = maxRetries
    m_restartCount = 0
    m_isHeadChecked = False
    m_lastReportedPercent = -1
    m_postProcessMacroName = postProcessMacroName
    
    Set m_http = New WinHttp.WinHttpRequest
    Set m_lossCallback = New StateLossCallback
    m_lossCallback.InitByMacroName "OnWorkerLostState", m_taskId, m_url, m_destPath
    ResetTracking
End Sub

Public Sub SetManager(ByVal manager As CAsyncDownloadManager)
    Set m_manager = manager
End Sub

Public Property Get TaskId() As Long
    TaskId = m_taskId
End Property

Public Property Get IsActive() As Boolean
    IsActive = (Not m_isCompleted) And (Not m_isCancelled)
End Property

Public Sub Start()
    On Error GoTo ErrorHandler
    
    If m_isCancelled Then Exit Sub
    
    ResetTracking
    m_isCompleted = False
    m_receivedBytes = 0@
    m_expectedSize = 0@
    
    m_startOffset = GetFileSizeByPath(m_destPath)
    If m_startOffset < 0 Then m_startOffset = 0@
    
    If Not m_isHeadChecked Then
        GetRemoteFileSize
        m_isHeadChecked = True
    End If
    
    If CheckAndCompleteIfFileExists() Then Exit Sub
    
    StartDownload
    Exit Sub
ErrorHandler:
    LogMessage LL_ERROR, "Ошибка запуска: " & Err.Description, Err.Number
    FailWithError "Ошибка запуска: " & Err.Description
End Sub

Private Sub GetRemoteFileSize()
    On Error GoTo ErrorHandler
    
    Dim headHttp As WinHttp.WinHttpRequest
    Set headHttp = New WinHttp.WinHttpRequest
    
    headHttp.SetTimeouts m_resolveTimeout, m_connectTimeout, m_sendTimeout, m_receiveTimeout
    headHttp.Open "HEAD", m_url, False
    headHttp.SetRequestHeader "Cache-Control", "no-cache"
    headHttp.Send
    
    If headHttp.status = 200 Or headHttp.status = 206 Then
        Dim contentLength As String
        contentLength = ""
        On Error Resume Next
        contentLength = headHttp.GetResponseHeader("Content-Length")
        On Error GoTo 0
        
        If contentLength <> "" And IsNumeric(contentLength) Then
            m_expectedSize = CDec(contentLength)
        End If
    End If
    Exit Sub
ErrorHandler:
    LogMessage LL_DEBUG, "HEAD запрос не удался: " & Err.Description
End Sub

Private Sub StartDownload()
    On Error GoTo ErrorHandler
    
    m_http.SetTimeouts m_resolveTimeout, m_connectTimeout, m_sendTimeout, m_receiveTimeout
    m_http.Open "GET", m_url, True
    
    If m_startOffset > 0 Then
        m_http.SetRequestHeader "Range", "bytes=" & Format$(m_startOffset, "0") & "-"
    End If
    
    m_http.SetRequestHeader "Cache-Control", "no-cache"
    m_http.SetRequestHeader "Pragma", "no-cache"
    m_http.SetRequestHeader "User-Agent", "Excel-Downloader/2.0"
    
    Set m_fileWriter = New CFileWriter
    If Not m_fileWriter.Init(m_destPath, m_bufferSize, m_startOffset > 0) Then
        FailWithError "Не удалось открыть файл: " & m_destPath
        Exit Sub
    End If
    
    m_http.Send
    m_lastActivityTime = Now
    
    If Not m_callback Is Nothing Then
        m_callback.OnStart m_taskId, m_url, m_destPath
    End If
    Exit Sub
ErrorHandler:
    LogMessage LL_ERROR, "Ошибка: " & Err.Description, Err.Number
    FailWithError "Ошибка запуска: " & Err.Description
End Sub

Public Function CheckAndCompleteIfFileExists() As Boolean
    If m_isCompleted Then
        CheckAndCompleteIfFileExists = True
        Exit Function
    End If
    
    If m_expectedSize <= 0 Or m_startOffset <= 0 Then
        CheckAndCompleteIfFileExists = False
        Exit Function
    End If
    
    If m_startOffset >= m_expectedSize Then
        CompleteWithSuccess
        CheckAndCompleteIfFileExists = True
    Else
        CheckAndCompleteIfFileExists = False
    End If
End Function

Private Sub CompleteWithSuccess()
    If m_isCompleted Then Exit Sub
    m_isCompleted = True
    
    On Error Resume Next
    If Not m_fileWriter Is Nothing Then
        m_fileWriter.CloseFile
        Set m_fileWriter = Nothing
    End If
    On Error GoTo 0
    
    If Not m_callback Is Nothing Then
        m_callback.OnComplete m_taskId, m_url, m_destPath, True, ""
    End If
    
    If Not m_manager Is Nothing Then
        m_manager.OnDownloadComplete m_taskId, m_destPath, m_postProcessMacroName
    End If
    
    RaiseEvent WorkerComplete(m_taskId, True)
    m_lossCallback.Cancel
    Set m_lossCallback = Nothing
End Sub

Private Sub FailWithError(ByVal errorMsg As String)
    If m_isCompleted Then Exit Sub
    m_isCompleted = True
    LogMessage LL_ERROR, errorMsg
    
    On Error Resume Next
    If Not m_fileWriter Is Nothing Then
        m_fileWriter.CloseFile
        Set m_fileWriter = Nothing
    End If
    On Error GoTo 0
    
    If Not m_callback Is Nothing Then
        m_callback.OnComplete m_taskId, m_url, m_destPath, False, errorMsg
    End If
    
    If Not m_manager Is Nothing Then
        m_manager.FinalizeTask m_taskId, False
    End If
    
    RaiseEvent WorkerComplete(m_taskId, False)
    m_lossCallback.Cancel
    Set m_lossCallback = Nothing
    
End Sub

Public Sub Cancel()
    m_isCancelled = True
    On Error Resume Next
    If Not m_http Is Nothing Then m_http.Abort
    If Not m_fileWriter Is Nothing Then m_fileWriter.CloseFile
    If Not m_lossCallback Is Nothing Then m_lossCallback.Cancel
    On Error GoTo 0
End Sub

Public Sub Restart()
    If m_restartCount >= m_maxRestarts Then
        FailWithError "Превышено количество рестартов (" & m_maxRestarts & ")"
        Exit Sub
    End If
    
    m_restartCount = m_restartCount + 1
    LogMessage LL_WARNING, "Рестарт #" & m_restartCount
    
    On Error Resume Next
    If Not m_fileWriter Is Nothing Then
        m_fileWriter.Flush
        m_fileWriter.CloseFile
        Set m_fileWriter = Nothing
    End If
    If Not m_http Is Nothing Then
        m_http.Abort
        Set m_http = Nothing
    End If
    On Error GoTo 0
    
    Set m_http = New WinHttp.WinHttpRequest
    m_isCompleted = False
    m_receivedBytes = 0@
    m_lastReportedPercent = -1
    m_isHeadChecked = False
    
    Start
End Sub

Public Function IsStuck() As Boolean
    If m_isCompleted Or m_isCancelled Then
        IsStuck = False
        Exit Function
    End If
    
    If m_lastActivityTime = 0 Then
        IsStuck = False
        Exit Function
    End If
    
    Dim idleSeconds As Double
    idleSeconds = (Now - m_lastActivityTime) * 86400#
    If idleSeconds < 0 Then idleSeconds = idleSeconds + 86400#
    
    Dim timeoutSec As Double
    timeoutSec = (m_receiveTimeout / 1000#) + 60
    
    If idleSeconds > timeoutSec Then
        IsStuck = True
    Else
        IsStuck = False
    End If
End Function

Private Sub LogMessage(ByVal level As Long, ByVal message As String, _
                       Optional ByVal errorCode As Long = 0)
    If Not m_callback Is Nothing Then
        m_callback.OnLog level, m_taskId, message, errorCode
    End If
End Sub

Private Sub ResetTracking()
    m_lastActivityTime = Now
    m_lastReportedPercent = -1
End Sub


Private Sub m_http_OnResponseDataAvailable(arrData() As Byte)
    If m_isCompleted Or m_isCancelled Then Exit Sub
    On Error GoTo ErrorHandler
    
    If Not m_fileWriter Is Nothing Then
        m_fileWriter.WriteBytes arrData
    End If
    
    Dim chunkSize As Currency
    chunkSize = UBound(arrData) - LBound(arrData) + 1
    m_receivedBytes = m_receivedBytes + chunkSize
    m_lastActivityTime = Now
    
    If m_expectedSize = 0 Then ParseContentSizeFromResponse
    
    If Not m_callback Is Nothing Then
        Dim percent As Integer
        If m_expectedSize > 0 Then
            percent = CInt((m_startOffset + m_receivedBytes) * 100@ / m_expectedSize)
            If percent > 100 Then percent = 100
        Else
            percent = -1
        End If
        
        If percent >= 0 And percent <> m_lastReportedPercent And (percent Mod 5 = 0 Or percent = 100) Then
            m_callback.OnProgress m_taskId, m_startOffset + m_receivedBytes, m_expectedSize, percent
            m_lastReportedPercent = percent
        End If
    End If
    Exit Sub
ErrorHandler:
    LogMessage LL_ERROR, "Ошибка в OnResponseDataAvailable: " & Err.Description, Err.Number
End Sub

Private Sub ParseContentSizeFromResponse()
    On Error Resume Next
    
    Dim rangeHeader As String
    rangeHeader = m_http.GetResponseHeader("Content-Range")
    
    If rangeHeader <> "" Then
        Dim slashPos As Long
        slashPos = InStrRev(rangeHeader, "/")
        If slashPos > 0 Then
            Dim sizeStr As String
            sizeStr = Mid$(rangeHeader, slashPos + 1)
            If IsNumeric(sizeStr) Then m_expectedSize = CDec(sizeStr)
        End If
    Else
        Dim lenHeader As String
        lenHeader = m_http.GetResponseHeader("Content-Length")
        If lenHeader <> "" And IsNumeric(lenHeader) Then
            m_expectedSize = CDec(lenHeader) + m_startOffset
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub m_http_OnResponseFinished()
    If m_isCompleted Or m_isCancelled Then Exit Sub
    
    If Not m_fileWriter Is Nothing Then m_fileWriter.Flush
    
    Dim success As Boolean
    Dim errorMsg As String
    
    Select Case m_http.status
        Case 200, 206, 416: success = True
        Case 404: success = False: errorMsg = "Файл не найден (404)"
        Case 403: success = False: errorMsg = "Доступ запрещён (403)"
        Case Else: success = False: errorMsg = "HTTP " & m_http.status & ": " & m_http.StatusText
    End Select
    
    If success Then
        If m_expectedSize > 0 Then
            Dim actualSize As Currency
            actualSize = GetFileSizeByPath(m_destPath)
            If actualSize >= m_expectedSize - 1024@ Then
                CompleteWithSuccess
            ElseIf m_restartCount < m_maxRestarts Then
                Restart
            Else
                FailWithError "Несоответствие размера"
            End If
        Else
            If m_receivedBytes > 0 Or GetFileSizeByPath(m_destPath) > 0 Then
                CompleteWithSuccess
            Else
                FailWithError "Не получено данных"
            End If
        End If
    Else
        FailWithError errorMsg
    End If
End Sub

Private Sub m_http_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    If m_isCompleted Or m_isCancelled Then Exit Sub
    
    LogMessage LL_ERROR, "WinHTTP ошибка " & ErrorNumber & ": " & ErrorDescription
    
    On Error Resume Next
    If Not m_fileWriter Is Nothing Then
        m_fileWriter.Flush
        m_fileWriter.CloseFile
        Set m_fileWriter = Nothing
    End If
    On Error GoTo 0
    
    If m_restartCount < m_maxRestarts Then
        Restart
    Else
        FailWithError "WinHTTP ошибка после " & m_restartCount & " рестартов"
    End If
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Cancel
    Set m_lossCallback = Nothing
    Set m_http = Nothing
    Set m_callback = Nothing
    Set m_manager = Nothing
    On Error GoTo 0
End Sub

