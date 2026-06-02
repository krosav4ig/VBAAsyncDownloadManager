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
' Класс: CFileWriter.cls
' Назначение: Буферизированная запись байтов в файл с задаваемым размером буфера.
'             Использует WinAPI для высокой производительности и поддержки больших файлов (>2GB).
' ============================================================================
Option Explicit

Private m_hFile As LongPtr
Private m_buffer() As Byte
Private m_bufferSize As Long
Private m_bufferPos As Long
Private m_isOpen As Boolean
Private m_filePath As String

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    #End If
#End If

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

' Инициализация: создаёт/перезаписывает файл, выделяет буфер
' Если файл уже существует (при рестарте), он открывается для дозаписи
Public Function Init(ByVal destPath As String, ByVal bufferSize As Long, Optional ByVal appendMode As Boolean = False) As Boolean
    If m_isOpen Then CloseFile
    
    ' Если режим дозаписи (append), мы не очищаем файл, а открываем существующий
    ' Логика создания/открытия вынесена в modFileAPI (OpenFileForWrite)
    m_hFile = OpenFileForWrite(destPath, appendMode)
    
    If m_hFile = 0 Or m_hFile = -1 Then
        Init = False
        Exit Function
    End If
    
    ReDim m_buffer(0 To bufferSize - 1)
    m_bufferSize = bufferSize
    m_bufferPos = 0
    m_filePath = destPath
    m_isOpen = True
    Init = True
End Function

' Добавление порции данных (буферизация)
' Данные копируются во внутренний буфер и сбрасываются на диск при заполнении
Public Sub WriteBytes(Data() As Byte)
    If Not m_isOpen Then Exit Sub
    
    Dim dataLen As Long
    dataLen = UBound(Data) - LBound(Data) + 1
    Dim offset As Long: offset = 0
    
    Do While offset < dataLen
        Dim chunk As Long
        chunk = dataLen - offset
        If chunk > (m_bufferSize - m_bufferPos) Then chunk = m_bufferSize - m_bufferPos
        
        ' Копирование памяти (быстрее побайтового цикла)
        CopyMemory m_buffer(m_bufferPos), Data(offset), chunk
        
        m_bufferPos = m_bufferPos + chunk
        offset = offset + chunk
        
        If m_bufferPos = m_bufferSize Then Flush
    Loop
End Sub

' Принудительная запись буфера на диск
Public Sub Flush()
    If Not m_isOpen Then Exit Sub
    If m_bufferPos > 0 Then
        If Not WriteToFile(m_hFile, m_buffer, m_bufferPos) Then
            Err.Raise 1001, "CFileWriter", "Ошибка записи в файл"
        End If
        m_bufferPos = 0
    End If
End Sub

' Закрытие файла (вызывать после завершения всех данных)
Public Sub CloseFile()
    If Not m_isOpen Then Exit Sub
    Flush
    CloseFileHandle m_hFile
    m_isOpen = False
    m_hFile = 0
End Sub

' Получение текущего размера файла (Currency)
Public Function GetFileSize() As Currency
    If Not m_isOpen Then
        GetFileSize = GetFileSizeByPath(m_filePath)
    Else
        Dim liSize As Currency
        If GetFileSizeEx(m_hFile, liSize) <> 0 Then
            GetFileSize = liSize * 10000
        Else
            GetFileSize = 0@
        End If
    End If
End Function

' Деструктор
Private Sub Class_Terminate()
    If m_isOpen Then CloseFile
End Sub
