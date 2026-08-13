VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDownloadCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Enum logLevel
    LL_CRITICAL = 1
    LL_ERROR = 2
    LL_WARNING = 3
    LL_INFO = 4
    LL_DEBUG = 5
End Enum

Public Sub OnStart(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String)
End Sub

Public Sub OnProgress(ByVal TaskId As Long, ByVal bytesReceived As Currency, _
                      ByVal totalBytes As Currency, ByVal percent As Integer)
End Sub

Public Sub OnComplete(ByVal TaskId As Long, ByVal url As String, ByVal destPath As String, _
                      ByVal success As Boolean, ByVal errorMsg As String)
End Sub

Public Sub OnLog(ByVal level As logLevel, ByVal TaskId As Long, ByVal message As String, _
                 Optional ByVal errorCode As Long = 0)
End Sub

