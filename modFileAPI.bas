Attribute VB_Name = "modFileAPI"
' ============================================================================
' Модуль: modFileAPI.bas
' Назначение: Работа с файлами через WinAPI с поддержкой больших файлов (>2GB)
' ============================================================================
Option Explicit

' Константы
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const OPEN_EXISTING As Long = 3
Private Const OPEN_ALWAYS As Long = 4
Private Const CREATE_ALWAYS As Long = 2
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_BEGIN As Long = 0
Private Const FILE_END As Long = 2

#If VBA7 Then
    ' 64-bit declarations
    Private Declare PtrSafe Function CreateFileW Lib "kernel32" ( _
        ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function WriteFile Lib "kernel32" ( _
        ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
        lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
    
    Private Declare PtrSafe Function ReadFile Lib "kernel32" ( _
        ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
        lpNumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
    
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As LongPtr) As Long
    
    Private Declare PtrSafe Function GetFileSizeEx Lib "kernel32" ( _
        ByVal hFile As LongPtr, lpFileSize As Currency) As Long
    
    Private Declare PtrSafe Function SetFilePointerEx Lib "kernel32" ( _
        ByVal hFile As LongPtr, ByVal liDistanceToMove As Currency, _
        lpNewFilePointer As Currency, ByVal dwMoveMethod As Long) As Long
    
    Private Declare PtrSafe Function FlushFileBuffers Lib "kernel32" ( _
        ByVal hFile As LongPtr) As Long
    
    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
#Else
    ' 32-bit declarations
    Private Declare Function CreateFileW Lib "kernel32" ( _
        ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long
    
    Private Declare Function WriteFile Lib "kernel32" ( _
        ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
        lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
    
    Private Declare Function ReadFile Lib "kernel32" ( _
        ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
        lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
    
    Private Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
    
    Private Declare Function GetFileSizeEx Lib "kernel32" ( _
        ByVal hFile As Long, lpFileSize As Currency) As Long
    
    Private Declare Function SetFilePointerEx Lib "kernel32" ( _
        ByVal hFile As Long, ByVal liDistanceToMove As Currency, _
        lpNewFilePointer As Currency, ByVal dwMoveMethod As Long) As Long
    
    Private Declare Function FlushFileBuffers Lib "kernel32" ( _
        ByVal hFile As Long) As Long
    
    Private Declare Function GetLastError Lib "kernel32" () As Long
#End If

' ============================================================================
' Получение размера файла (возвращает байты как Currency, поддерживает >2GB)
' ============================================================================
Public Function GetFileSizeByPath(ByVal filePath As String) As Currency
    On Error GoTo ErrorHandler
    
    If filePath = "" Then
        GetFileSizeByPath = 0@
        Exit Function
    End If
    
    #If VBA7 Then
        Dim hFile As LongPtr
    #Else
        Dim hFile As Long
    #End If
    
    ' Открываем файл только для чтения
    hFile = CreateFileW(StrPtr(filePath), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                        0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    
    If hFile = INVALID_HANDLE_VALUE Then
        ' Файл не существует или недоступен
        GetFileSizeByPath = -1@
        Exit Function
    End If
    
    Dim fileSize As Currency
    If GetFileSizeEx(hFile, fileSize) <> 0 Then
        ' GetFileSizeEx возвращает размер в 64-битном формате,
        ' Currency хранит как масштабированное целое (4 знака)
        GetFileSizeByPath = fileSize * 10000
    Else
        GetFileSizeByPath = -1@
    End If
    
    CloseHandle hFile
    Exit Function
    
ErrorHandler:
    GetFileSizeByPath = -1@
End Function

' ============================================================================
' Проверка существования и размера файла
' ============================================================================
Public Function FileExistsAndGetSize(ByVal filePath As String, ByRef outSize As Currency) As Boolean
    outSize = GetFileSizeByPath(filePath)
    FileExistsAndGetSize = (outSize >= 0@)
End Function

' ============================================================================
' Открытие файла для записи (с поддержкой дозаписи)
' ============================================================================
#If VBA7 Then
Public Function OpenFileForWrite(ByVal filePath As String, Optional ByVal appendMode As Boolean = False) As LongPtr
#Else
Public Function OpenFileForWrite(ByVal filePath As String, Optional ByVal appendMode As Boolean = False) As Long
#End If
    Dim hFile As LongPtr
    Dim createDisposition As Long
    
    If appendMode Then
        ' Режим дозаписи - открываем существующий или создаём новый
        createDisposition = OPEN_ALWAYS
    Else
        ' Режим перезаписи
        createDisposition = CREATE_ALWAYS
    End If
    
    hFile = CreateFileW(StrPtr(filePath), GENERIC_WRITE, FILE_SHARE_READ, _
                        0, createDisposition, FILE_ATTRIBUTE_NORMAL, 0)
    
    If hFile = INVALID_HANDLE_VALUE Then
        OpenFileForWrite = 0
        Exit Function
    End If
    
    ' Если режим дозаписи - перемещаем указатель в конец
    If appendMode Then
        Dim fileSize As Currency
        If GetFileSizeEx(hFile, fileSize) <> 0 Then
            Dim newPos As Currency
            SetFilePointerEx hFile, fileSize, newPos, FILE_BEGIN
        End If
    End If
    
    OpenFileForWrite = hFile
End Function

' ============================================================================
' Запись данных в файл
' ============================================================================
#If VBA7 Then
Public Function WriteToFile(ByVal hFile As LongPtr, buffer() As Byte, ByVal bytesToWrite As Long) As Boolean
#Else
Public Function WriteToFile(ByVal hFile As Long, buffer() As Byte, ByVal bytesToWrite As Long) As Boolean
#End If
    If bytesToWrite <= 0 Then
        WriteToFile = True
        Exit Function
    End If
    
    Dim written As Long
    Dim result As Long
    
    result = WriteFile(hFile, buffer(0), bytesToWrite, written, 0)
    
    WriteToFile = (result <> 0 And written = bytesToWrite)
    
    If Not WriteToFile Then
        Debug.Print "WriteFile error: " & GetLastError() & ", attempted: " & bytesToWrite & ", written: " & written
    End If
End Function

' ============================================================================
#If VBA7 Then
Public Sub CloseFileHandle(ByVal hFile As LongPtr)
#Else
Public Sub CloseFileHandle(ByVal hFile As Long)
#End If
    If hFile <> 0 And hFile <> INVALID_HANDLE_VALUE Then
        FlushFileBuffers hFile
        CloseHandle hFile
    End If
End Sub

