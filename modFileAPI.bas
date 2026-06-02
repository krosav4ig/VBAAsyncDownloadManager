Attribute VB_Name = "modFileAPI"
' ============================================================================
' Модуль: modFileAPI.bas (стандартный)
' Назначение: Обертки над WinAPI для работы с файлами (открытие, запись, размер).
'             Поддержка 64-бит и файлов размером более 2 ГБ (тип Currency).
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

' Получение размера файла по пути
Public Function GetFileSizeByPath(ByVal filePath As String) As Currency
    On Error Resume Next
    #If VBA7 Then
        Dim hFile As LongPtr
    #Else
        Dim hFile As Long
    #End If
    Dim size As Currency
    
    ' Открываем файл только для чтения атрибутов
    hFile = CreateFileW(StrPtr(filePath), 0, FILE_SHARE_READ, ByVal 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    
    If hFile <> INVALID_HANDLE_VALUE Then
        If GetFileSizeEx(hFile, size) <> 0 Then
            GetFileSizeByPath = size * 10000
        Else
            GetFileSizeByPath = 0@
        End If
        CloseHandle hFile
    Else
        GetFileSizeByPath = 0@
    End If
End Function

' Открытие файла для записи
#If VBA7 Then
Public Function OpenFileForWrite(ByVal filePath As String, Optional ByVal appendMode As Boolean = False) As LongPtr
#Else
Public Function OpenFileForWrite(ByVal filePath As String, Optional ByVal appendMode As Boolean = False) As Long
#End If
    Dim dwCreationDisposition As Long
    If appendMode Then
        ' Если файл существует, открываем его (OPEN_EXISTING), иначе создаем (OPEN_ALWAYS)
        ' Но для простоты используем OPEN_ALWAYS всегда, а позицию ставим вручную, если нужно
        dwCreationDisposition = OPEN_ALWAYS
    Else
        ' Перезапись (CREATE_ALWAYS) - но OPEN_ALWAYS с обнулением тоже подойдет, если мы сами контролируем
        ' Для надежной перезаписи лучше CREATE_ALWAYS (2), но тогда файл удалится перед созданием.
        ' Оставим OPEN_ALWAYS (4) и будем писать с начала, если не append.
        dwCreationDisposition = 2 ' CREATE_ALWAYS для чистой перезаписи
    End If
    
    Dim hFile As LongPtr
    hFile = CreateFileW(StrPtr(filePath), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0, dwCreationDisposition, FILE_ATTRIBUTE_NORMAL, 0)
    
    If hFile = INVALID_HANDLE_VALUE Then
        OpenFileForWrite = 0
        Exit Function
    End If
    
    ' Если режим дозаписи, перемещаем указатель в конец
    If appendMode Then
        Dim fileSize As Currency
        fileSize = 0@
        ' Получаем размер и ставим указатель
        ' Прямой вызов API для перемещения указателя
        If GetFileSizeEx(hFile, fileSize) <> 0 Then
             Dim dummy As Currency
             SetFilePointerEx hFile, fileSize, dummy, FILE_BEGIN
        End If
    End If
    
    OpenFileForWrite = hFile
End Function

' Запись буфера в файл
#If VBA7 Then
Public Function WriteToFile(ByVal hFile As LongPtr, buffer() As Byte, ByVal bytesToWrite As Long) As Boolean
#Else
Public Function WriteToFile(ByVal hFile As Long, buffer() As Byte, ByVal bytesToWrite As Long) As Boolean
#End If
    Dim written As Long
    If bytesToWrite = 0 Then
        WriteToFile = True
        Exit Function
    End If
    
    Dim res As Long
    res = WriteFile(hFile, buffer(0), bytesToWrite, written, ByVal 0)
    
    WriteToFile = (res <> 0 And written = bytesToWrite)
End Function

' Закрытие хэндла
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

