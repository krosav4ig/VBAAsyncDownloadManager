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
' Интерфейс: IDownloadCallback
' Назначение: Колбэки для отслеживания загрузки
' ============================================================================
Option Explicit

' Уровни логирования
Public Enum logLevel
    LL_CRITICAL = 1  ' Только критические ошибки
    LL_ERROR = 2     ' Ошибки
    LL_WARNING = 3   ' Предупреждения
    LL_INFO = 4      ' Информационные сообщения
    LL_DEBUG = 5     ' Детальная отладка
End Enum

' Начало загрузки
Public Sub OnStart(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String)
End Sub

' Прогресс загрузки
Public Sub OnProgress(ByVal TaskId As Long, ByVal bytesReceived As Currency, _
                      ByVal totalBytes As Currency, ByVal percent As Integer)
End Sub

' Завершение загрузки
Public Sub OnComplete(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String, _
                      ByVal success As Boolean, ByVal errorMsg As String)
End Sub

' Логирование (общий метод)
Public Sub OnLog(ByVal level As logLevel, ByVal TaskId As Long, ByVal message As String, _
                 Optional ByVal errorCode As Long = 0)
End Sub

