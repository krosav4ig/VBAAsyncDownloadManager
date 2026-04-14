VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDownloadCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' Модуль: IDownloadCallback.cls
' Назначение: Интерфейс для callback-уведомлений (старт, прогресс, завершение)
' ============================================================================
Option Explicit

' Уведомление о начале загрузки (вызывается один раз при старте)
' taskId   - идентификатор задачи
' url      - исходный URL
' destPath - путь сохранения
Public Sub OnStart(ByVal taskId As Long, ByVal url As String, ByVal destPath As String)
End Sub

' Уведомление о прогрессе загрузки
' taskId        - идентификатор задачи
' bytesReceived - получено байт (Currency для больших файлов)
' totalBytes    - всего байт (0, если неизвестно)
' percent       - процент выполнения (от 0 до 100, -1 если неизвестно)
Public Sub OnProgress(ByVal taskId As Long, ByVal bytesReceived As Currency, _
                      ByVal totalBytes As Currency, ByVal percent As Integer)
End Sub

' Уведомление о завершении загрузки
' taskId        - идентификатор задачи
' url           - исходный URL
' destPath      - путь сохранения
' success       - успех (True/False)
' errorMsg      - текст ошибки (пусто, если success=True)
Public Sub OnComplete(ByVal taskId As Long, ByVal url As String, ByVal destPath As String, _
                      ByVal success As Boolean, ByVal errorMsg As String)
End Sub

