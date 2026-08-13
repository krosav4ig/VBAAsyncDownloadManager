VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CManagerEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private WithEvents m_manager As CAsyncDownloadManager
Attribute m_manager.VB_VarHelpID = -1

Private m_logger As CDownloadLogger

Public Sub Init(ByVal manager As CAsyncDownloadManager, ByVal logger As CDownloadLogger)
    Set m_manager = manager
    Set m_logger = logger
End Sub

Private Sub m_manager_AllTasksCompleted()
    If Not m_logger Is Nothing Then
        m_logger.LogMessage LL_INFO, 0, "[EVENT] Все задачи завершены. Успешно: " & _
            m_manager.CompletedCount & ", Ошибок: " & m_manager.FailedCount
    End If
End Sub

Private Sub m_manager_TaskCompleted(ByVal TaskId As Long, ByVal success As Boolean)
    If Not m_logger Is Nothing Then
        If success Then
            m_logger.LogMessage LL_DEBUG, TaskId, "[EVENT] Задача успешно завершена"
        Else
            m_logger.LogMessage LL_WARNING, TaskId, "[EVENT] Задача завершена с ошибкой"
        End If
    End If
End Sub

Private Sub m_manager_ProgressUpdated()
    If Not m_logger Is Nothing Then
        Dim postProcCount As Long
        postProcCount = 0
        If Not g_pendingAsyncMarkers Is Nothing Then postProcCount = g_pendingAsyncMarkers.count
        
        m_logger.LogMessage LL_INFO, 0, "[EVENT] Активно: " & m_manager.ActiveCount & _
            " | Ожидает: " & m_manager.PendingCount & _
            " | Готово: " & m_manager.CompletedCount & _
            " | Ошибок: " & m_manager.FailedCount & _
            " | Постобр: " & postProcCount
    End If
End Sub

Private Sub Class_Terminate()
    Set m_manager = Nothing
    Set m_logger = Nothing
End Sub
