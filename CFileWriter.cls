VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' Класс: CFileWriter
' Назначение: Буферизированная запись в файл с поддержкой больших файлов
' ============================================================================
Option Explicit

Private Const MAX_BUFFER_SIZE As Long = 1048576  ' 1MB максимум

Private m_hFile As LongPtr
Private m_buffer() As Byte
Private m_bufferSize As Long
Private m_bufferPos As Long
Private m_isOpen As Boolean
Private m_filePath As String
Private m_writeErrorCount As Long
Private m_totalBytesWritten As Currency

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        Destination As Any, Source As Any, ByVal Length As Long)
#End If

' ============================================================================
' Инициализация
' ============================================================================
Public Function Init(ByVal destPath As String, ByVal bufferSize As Long, _
                     Optional ByVal appendMode As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    If m_isOpen Then CloseFile
    
    ' Ограничение размера буфера
    If bufferSize < 4096 Then bufferSize = 4096
    If bufferSize > MAX_BUFFER_SIZE Then bufferSize = MAX_BUFFER_SIZE
    
    m_bufferSize = bufferSize
    m_filePath = destPath
    m_writeErrorCount = 0
    m_totalBytesWritten = 0@
    
    ' Создание директории при необходимости
    If Not CreateDirectoryRecursive(GetParentPath(destPath)) Then
        Init = False
        Exit Function
    End If
    
    ' Открытие файла
    m_hFile = OpenFileForWrite(destPath, appendMode)
    
    If m_hFile = 0 Then
        Init = False
        Exit Function
    End If
    
    ' Выделение буфера
    ReDim m_buffer(0 To m_bufferSize - 1)
    m_bufferPos = 0
    m_isOpen = True
    Init = True
    Exit Function
    
ErrorHandler:
    Init = False
End Function

' ============================================================================
' Запись данных
' ============================================================================
Public Sub WriteBytes(Data() As Byte)
    If Not m_isOpen Then Exit Sub
    If m_writeErrorCount > 3 Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim dataLen As Long
    dataLen = UBound(Data) - LBound(Data) + 1
    If dataLen <= 0 Then Exit Sub
    
    Dim offset As Long: offset = 0
    
    Do While offset < dataLen
        Dim chunk As Long
        chunk = dataLen - offset
        If chunk > (m_bufferSize - m_bufferPos) Then chunk = m_bufferSize - m_bufferPos
        
        CopyMemory m_buffer(m_bufferPos), Data(offset), chunk
        
        m_bufferPos = m_bufferPos + chunk
        offset = offset + chunk
        m_totalBytesWritten = m_totalBytesWritten + chunk
        
        If m_bufferPos >= m_bufferSize Then
            If Not FlushInternal() Then
                m_writeErrorCount = m_writeErrorCount + 1
                Exit Sub
            End If
        End If
    Loop
    
    Exit Sub
    
ErrorHandler:
    m_writeErrorCount = m_writeErrorCount + 1
End Sub

' ============================================================================
' Внутренний сброс буфера
' ============================================================================
Private Function FlushInternal() As Boolean
    If Not m_isOpen Then
        FlushInternal = False
        Exit Function
    End If
    
    If m_bufferPos > 0 Then
        If Not WriteToFile(m_hFile, m_buffer, m_bufferPos) Then
            FlushInternal = False
            Exit Function
        End If
        m_bufferPos = 0
    End If
    
    FlushInternal = True
End Function

' ============================================================================
' Публичный сброс буфера
' ============================================================================
Public Sub Flush()
    If Not m_isOpen Then Exit Sub
    FlushInternal
End Sub

' ============================================================================
' Закрытие файла
' ============================================================================
Public Sub CloseFile()
    If Not m_isOpen Then Exit Sub
    
    FlushInternal
    CloseFileHandle m_hFile
    m_isOpen = False
    m_hFile = 0
    Erase m_buffer
End Sub

' ============================================================================
' Получение текущего размера (файл + буфер)
' ============================================================================
Public Function GetCurrentSize() As Currency
    If m_isOpen Then
        GetCurrentSize = m_totalBytesWritten
    Else
        GetCurrentSize = GetFileSizeByPath(m_filePath)
        If GetCurrentSize < 0 Then GetCurrentSize = 0
    End If
End Function

' ============================================================================
' Проверка целостности
' ============================================================================
Public Function IsHealthy() As Boolean
    IsHealthy = (m_writeErrorCount <= 3)
End Function

' ============================================================================
' Вспомогательные функции
' ============================================================================
Private Function GetParentPath(ByVal filePath As String) As String
    Dim pos As Integer
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        GetParentPath = Left(filePath, pos - 1)
    Else
        GetParentPath = ""
    End If
End Function

Private Function CreateDirectoryRecursive(ByVal dirPath As String) As Boolean
    If dirPath = "" Then
        CreateDirectoryRecursive = True
        Exit Function
    End If
    
    On Error Resume Next
    MkDir dirPath
    If Err.Number = 75 Then ' Path already exists
        Err.Clear
        CreateDirectoryRecursive = True
    ElseIf Err.Number = 76 Then ' Path not found - create parent
        Err.Clear
        CreateDirectoryRecursive = CreateDirectoryRecursive(GetParentPath(dirPath))
        If CreateDirectoryRecursive Then
            MkDir dirPath
            CreateDirectoryRecursive = (Err.Number = 0)
        End If
    Else
        CreateDirectoryRecursive = (Err.Number = 0)
    End If
    On Error GoTo 0
End Function

' ============================================================================
Private Sub Class_Terminate()
    CloseFile
End Sub

