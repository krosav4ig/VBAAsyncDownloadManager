Attribute VB_Name = "modPostProcess"
Option Explicit

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
    markerPath = filePath & ".done"
    
    ' Генерируем VBScript
    Dim vbsPath As String
    vbsPath = Environ$("TEMP") & "\convert_xlsx_" & GetTickCount() & "_" & g_currentPostProcessTaskId & ".vbs"
    
    Dim vbsCode As String
    vbsCode = GenerateConvertVbs(filePath, csvPath, markerPath)
    
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
                                     ByVal markerPath As String) As String
    Dim code As String
    
    code = "On Error Resume Next" & vbCrLf
    code = code & "Dim xlApp, wb, fso, ts" & vbCrLf
    code = code & "Set xlApp = createobject(""excel.application"")" & vbCrLf
    'code = code & "Set xlApp = GetObject(""" & Environ("tmp") & "\fileconverter.xlsm"").Parent" & vbCrLf
    code = code & "xlApp.Visible = False" & vbCrLf
    code = code & "xlApp.DisplayAlerts = False" & vbCrLf
    code = code & "xlApp.AutomationSecurity = 3" & vbCrLf
    code = code & vbCrLf
    code = code & "Set wb = xlApp.Workbooks.Open(""" & xlsxPath & """, 0, True)" & vbCrLf
    code = code & "If Err.Number = 0 Then" & vbCrLf
    code = code & "  wb.SaveAs """ & csvPath & """, 6,,,,,,,,,,true" & vbCrLf
    code = code & "  wb.Close False" & vbCrLf
    code = code & "  Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    code = code & "  fso.DeleteFile """ & xlsxPath & """, True" & vbCrLf
    code = code & "Else" & vbCrLf
    code = code & "  If Not wb Is Nothing Then wb.Close False" & vbCrLf
    code = code & "End If" & vbCrLf
    code = code & vbCrLf
    code = code & "Set wb = Nothing" & vbCrLf
    code = code & "Set xlApp = Nothing" & vbCrLf
    code = code & vbCrLf
    code = code & "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    code = code & "Set ts = fso.CreateTextFile(""" & markerPath & """, True)" & vbCrLf
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
        Dim TaskId As Long: TaskId = item(1)
        
        If fso.FileExists(markerPath) Then
            ' Маркер найден - конвертация завершена
            On Error Resume Next
            fso.DeleteFile markerPath, True
            On Error GoTo 0
            
            Debug.Print "[PostProcess] Конвертация завершена для TaskId=" & TaskId
            
            If Not g_manager Is Nothing Then
                g_manager.FinalizeTask TaskId, True
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
        If fso.FileExists(markerPath) Then
            fso.DeleteFile markerPath, True
        End If
    Next i
    
    Set fso = Nothing
    g_pendingAsyncMarkers.Clear
    On Error GoTo 0
End Sub

