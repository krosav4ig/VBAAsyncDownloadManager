Attribute VB_Name = "modFileAPI"
' ============================================================================
' Модуль: modFileAPI.bas (стандартный)
' Назначение: WinAPI для бинарной записи, получения размера файла, таймеров
' ============================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function CreateFileW Lib "kernel32" ( _
        ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As LongPtr) As LongPtr
        
    Private Declare PtrSafe Function WriteFile Lib "kernel32" ( _
        ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
        lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
        
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As LongPtr) As Long
        
    Public Declare PtrSafe Function GetFileSizeEx Lib "kernel32" ( _
        ByVal hFile As LongPtr, lpFileSize As Currency) As Long
        
    Private Declare PtrSafe Function SetFilePointerEx Lib "kernel32" ( _
        ByVal hFile As LongPtr, liDistanceToMove As Currency, lpNewFilePointer As Currency, _
        ByVal dwMoveMethod As Long) As Long
        
    Private Declare PtrSafe Function FlushFileBuffers Lib "kernel32" ( _
        ByVal hFile As LongPtr) As Long
#Else
    Private Declare Function CreateFileW Lib "kernel32" ( _
        ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long
        
    Private Declare Function WriteFile Lib "kernel32" ( _
        ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
        lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
        
    Private Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
        
    Public Declare Function GetFileSizeEx Lib "kernel32" ( _
        ByVal hFile As Long, lpFileSize As Currency) As Long
        
    Private Declare Function SetFilePointerEx Lib "kernel32" ( _
        ByVal hFile As Long, liDistanceToMove As Currency, lpNewFilePointer As Currency, _
        ByVal dwMoveMethod As Long) As Long
        
    Private Declare Function FlushFileBuffers Lib "kernel32" ( _
        ByVal hFile As Long) As Long
#End If

Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const OPEN_ALWAYS As Long = 4
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const INVALID_HANDLE_VALUE As LongPtr = -1
Private Const FILE_BEGIN As Long = 0

' Вспомогательная функция для получения размера файла по пути (Currency)
Public Function GetFileSizeByPath(ByVal filePath As String) As Currency
    On Error Resume Next
    #If VBA7 Then
        Dim hFile As LongPtr
    #Else
        Dim hFile As Long
    #End If
    Dim size As Currency
    hFile = CreateFileW(StrPtr(filePath), 0, FILE_SHARE_READ, ByVal 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        If GetFileSizeEx(hFile, size) <> 0 Then
            GetFileSizeByPath = size
        Else
            GetFileSizeByPath = 0@
        End If
        CloseHandle hFile
    Else
        GetFileSizeByPath = 0@
    End If
End Function

' Открытие файла для бинарной записи (возвращает хэндл)
#If VBA7 Then
Public Function OpenFileForWrite(ByVal filePath As String) As LongPtr
#Else
Public Function OpenFileForWrite(ByVal filePath As String) As Long
#End If
    OpenFileForWrite = CreateFileW(StrPtr(filePath), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
End Function

' Запись буфера в файл (возвращает True при успехе)
#If VBA7 Then
Public Function WriteToFile(ByVal hFile As LongPtr, buffer() As Byte, ByVal bytesToWrite As Long) As Boolean
#Else
Public Function WriteToFile(ByVal hFile As Long, buffer() As Byte, ByVal bytesToWrite As Long) As Boolean
#End If
    Dim written As Long
    If bytesToWrite = 0 Then WriteToFile = True: Exit Function
    Dim res As Long
    res = WriteFile(hFile, buffer(0), bytesToWrite, written, ByVal 0)
    WriteToFile = (res <> 0 And written = bytesToWrite)
End Function

' Сброс буферов и закрытие файла
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

