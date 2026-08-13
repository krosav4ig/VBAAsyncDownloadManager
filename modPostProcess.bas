Attribute VB_Name = "modPostProcess"
Option Explicit

' Константы для маркеров
Public Const DONE_EXTENSION As String = ".done"
Public Const ERR_EXTENSION As String = ".err"
Public Const TMP_EXTENSION As String = ".tmp"

#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

' Глобальный контекст для асинхронной постобработки
Public g_currentPostProcessTaskId As Long
Public g_pendingAsyncMarkers As Collection

Public Sub InitPostProcess()
    Set g_pendingAsyncMarkers = New Collection
    g_currentPostProcessTaskId = 0
End Sub

' ============================================================================
' Главная функция: асинхронная конвертация XLSX в CSV
' Возвращает Long: 1 = асинхронно запущено, 0 = ошибка/не тот формат
' ============================================================================
Public Function ConvertXlsxToCsv_Async(ByVal filePath As String) As Long
    On Error GoTo ErrorHandler
    
    ' Проверка расширения
    Dim ext As String
    ext = LCase(Right$(filePath, 5))
    If ext <> ".xlsx" And LCase(Right$(filePath, 4)) <> ".xls" Then
        ConvertXlsxToCsv_Async = 0
        Exit Function
    End If
    
    ' Проверка существования файла
    If Dir(filePath) = "" Then
        ConvertXlsxToCsv_Async = 0
        Exit Function
    End If
    
    ' Формируем пути
    Dim csvPath As String
    csvPath = Left$(filePath, InStrRev(filePath, ".") - 1) & ".csv"
    
    Dim markerPath As String
    markerPath = filePath & DONE_EXTENSION
    
    Dim errPath As String
    errPath = filePath & ERR_EXTENSION
    
    ' Генерируем VBScript
    Dim vbsPath As String
    vbsPath = Environ$("TEMP") & "\convert_xlsx_" & GetTickCount() & "_" & g_currentPostProcessTaskId & ".vbs"
    
    Dim vbsCode As String
    vbsCode = GenerateConvertVbs(filePath, csvPath, markerPath, errPath)
    
    ' Записываем VBScript во временный файл
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.CreateTextFile(vbsPath, True, True) ' True = Unicode
    ts.Write vbsCode
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    ' Запускаем VBScript асинхронно (0 = скрытое окно, False = не ждать)
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run """" & vbsPath & """", 0, False
    Set wsh = Nothing
    
    ' Регистрируем маркер для отслеживания
    If g_pendingAsyncMarkers Is Nothing Then Set g_pendingAsyncMarkers = New Collection
    
    Dim taskInfo As Variant
    taskInfo = Array(markerPath, g_currentPostProcessTaskId)
    
    On Error Resume Next
    g_pendingAsyncMarkers.Add taskInfo, markerPath
    On Error GoTo 0
    
    Debug.Print "[PostProcess] Асинхронная конвертация запущена: " & filePath
    ConvertXlsxToCsv_Async = 1
    Exit Function
    
ErrorHandler:
    Debug.Print "[PostProcess] Ошибка запуска конвертации: " & Err.Description
    ConvertXlsxToCsv_Async = 0
End Function

' ============================================================================
' Генерация кода VBScript для конвертации
' ============================================================================
Private Function GenerateConvertVbs(ByVal xlsxPath As String, _
                                     ByVal csvPath As String, _
                                     ByVal markerPath As String, _
                                     ByVal errPath As String) As String
    Dim code As String
    
    code = "On Error Resume Next" & vbCrLf
    code = code & "Dim xlApp, wb, fso, ts, errHappened" & vbCrLf
    code = code & "errHappened = False" & vbCrLf
    code = code & "Set xlApp = createobject(\"\"excel.application\"\")" & vbCrLf
    code = code & "xlApp.Visible = False" & vbCrLf
    code = code & "xlApp.DisplayAlerts = False" & vbCrLf
    code = code & "xlApp.AutomationSecurity = 3" & vbCrLf
    code = code & vbCrLf
    code = code & "Set wb = xlApp.Workbooks.Open(\"\"" & xlsxPath & "\"\", 0, True)" & vbCrLf
    code = code & "If Err.Number = 0 Then" & vbCrLf
    code = code & "  wb.SaveAs \"\"" & csvPath & "\"\", 6,,,,,,,,,,true" & vbCrLf
    code = code & "  If Err.Number <> 0 Then errHappened = True" & vbCrLf
    code = code & "  wb.Close False" & vbCrLf
    code = code & "  If Not errHappened Then" & vbCrLf
    code = code & "    Set fso = CreateObject(\"\"Scripting.FileSystemObject\"\")" & vbCrLf
    code = code & "    fso.DeleteFile \"\"" & xlsxPath & "\"\", True" & vbCrLf
    code = code & "    If Err.Number <> 0 Then errHappened = True" & vbCrLf
    code = code & "  End If" & vbCrLf
    code = code & "Else" & vbCrLf
    code = code & "  errHappened = True" & vbCrLf
    code = code & "  If Not wb Is Nothing Then wb.Close False" & vbCrLf
    code = code & "End If" & vbCrLf
    code = code & vbCrLf
    code = code & "Set wb = Nothing" & vbCrLf
    code = code & "Set xlApp = Nothing" & vbCrLf
    code = code & vbCrLf
    code = code & "Set fso = CreateObject(\"\"Scripting.FileSystemObject\"\")" & vbCrLf
    code = code & "If errHappened Then" & vbCrLf
    code = code & "  Set ts = fso.CreateTextFile(\"\"" & errPath & "\"\", True)" & vbCrLf
    code = code & "Else" & vbCrLf
    code = code & "  Set ts = fso.CreateTextFile(\"\"" & markerPath & "\"\", True)" & vbCrLf
    code = code & "End If" & vbCrLf
    code = code & "ts.Close" & vbCrLf
    GenerateConvertVbs = code
End Function

' ============================================================================
' Проверка маркеров (вызывается из TimerCheck)
' ============================================================================
Public Sub CheckAsyncMarkers()

    If g_pendingAsyncMarkers Is Nothing Then Exit Sub
    If g_pendingAsyncMarkers.count = 0 Then Exit Sub
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim i As Long: i = 1
    
    Do While i <= g_pendingAsyncMarkers.count
        Dim item As Variant
        item = g_pendingAsyncMarkers.item(i)
        
        Dim markerPath As String: markerPath = item(0)
        Dim taskId As Long: taskId = item(1)
        
        ' Формируем путь к файлу ошибки на основе базового пути
        Dim basePath As String
        basePath = Left(markerPath, Len(markerPath) - Len(DONE_EXTENSION))
        Dim errPath As String
        errPath = basePath & ERR_EXTENSION
        
        ' Проверяем наличие маркера успеха (.done) или ошибки (.err)
        Dim successMarker As Boolean
        Dim errorMarker As Boolean
        successMarker = fso.FileExists(markerPath)
        errorMarker = fso.FileExists(errPath)
        
        If successMarker Or errorMarker Then
            ' Конвертация завершена (успешно или с ошибкой)
            On Error Resume Next
            If successMarker Then fso.DeleteFile markerPath, True
            If errorMarker Then fso.DeleteFile errPath, True
            On Error GoTo 0
            
            If successMarker Then
                Debug.Print "[PostProcess] Конвертация успешна для TaskId=" & taskId
                If Not g_manager Is Nothing Then
                    g_manager.FinalizeTask taskId, True
                End If
            Else
                Debug.Print "[PostProcess] Конвертация с ошибкой для TaskId=" & taskId & ", файл: " & errPath
                If Not g_manager Is Nothing Then
                    g_manager.FinalizeTask taskId, False
                End If
            End If
            
            g_pendingAsyncMarkers.Remove i
        Else
            i = i + 1
        End If
    Loop
    
    Set fso = Nothing
End Sub

' ============================================================================
' Очистка очереди (при остановке)
' ============================================================================
Public Sub ClearAsyncMarkers()
    If g_pendingAsyncMarkers Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim i As Long
    For i = 1 To g_pendingAsyncMarkers.count
        Dim item As Variant
        item = g_pendingAsyncMarkers.item(i)
        Dim markerPath As String: markerPath = item(0)
        
        Dim errPath As String
        errPath = Left(markerPath, Len(markerPath) - Len(DONE_EXTENSION)) & ERR_EXTENSION
        
        If fso.FileExists(markerPath) Then
            fso.DeleteFile markerPath, True
        End If
        If fso.FileExists(errPath) Then
            fso.DeleteFile errPath, True
        End If
    Next i
    
    Set fso = Nothing
    g_pendingAsyncMarkers.Clear
    On Error GoTo 0
End Sub

