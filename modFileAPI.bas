Attribute VB_Name = "modFileAPI"

Option Explicit

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

#If VBA7 Then
    Private Declare PtrSafe Function CreateFileW Lib "kernel32" ( _
        ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function WriteFile Lib "kernel32" ( _
        ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
        lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
    
    Private Declare PtrSafe Function GetFileSizeEx Lib "kernel32" ( _
        ByVal hFile As LongPtr, lpFileSize As Currency) As Long
    
    Private Declare PtrSafe Function SetFilePointerEx Lib "kernel32" ( _
        ByVal hFile As LongPtr, ByVal liDistanceToMove As Currency, _
        lpNewFilePointer As Currency, ByVal dwMoveMethod As Long) As Long
    
    Private Declare PtrSafe Function FlushFileBuffers Lib "kernel32" ( _
        ByVal hFile As LongPtr) As Long
    
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As LongPtr) As Long
    
    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
#Else
    Private Declare Function CreateFileW Lib "kernel32" ( _
        ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long
    
    Private Declare Function WriteFile Lib "kernel32" ( _
        ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
        lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
    
    Private Declare Function GetFileSizeEx Lib "kernel32" ( _
        ByVal hFile As Long, lpFileSize As Currency) As Long
    
    Private Declare Function SetFilePointerEx Lib "kernel32" ( _
        ByVal hFile As Long, ByVal liDistanceToMove As Currency, _
        lpNewFilePointer As Currency, ByVal dwMoveMethod As Long) As Long
    
    Private Declare Function FlushFileBuffers Lib "kernel32" ( _
        ByVal hFile As Long) As Long
    
    Private Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long
    
    Private Declare Function GetLastError Lib "kernel32" () As Long
#End If

Public Function GetFileSizeByPath(ByVal filePath As String) As Currency
    On Error GoTo ErrorHandler
    If filePath = "" Then GetFileSizeByPath = 0@: Exit Function
    
#If VBA7 Then
    Dim hFile As LongPtr
#Else
    Dim hFile As Long
#End If
    
    hFile = CreateFileW(StrPtr(filePath), GENERIC_READ, _
                        FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, _
                        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    
    If hFile = INVALID_HANDLE_VALUE Then
        GetFileSizeByPath = -1@
        Exit Function
    End If
    
    Dim fileSize As Currency
    If GetFileSizeEx(hFile, fileSize) <> 0 Then
        GetFileSizeByPath = fileSize * 10000
    Else
        GetFileSizeByPath = -1@
    End If
    
    CloseHandle hFile
    Exit Function
ErrorHandler:
    GetFileSizeByPath = -1@
End Function

#If VBA7 Then
Public Function OpenFileForWrite(ByVal filePath As String, _
                                  Optional ByVal appendMode As Boolean = False) As LongPtr
#Else
Public Function OpenFileForWrite(ByVal filePath As String, _
                                  Optional ByVal appendMode As Boolean = False) As Long
#End If
#If VBA7 Then
    Dim hFile As LongPtr
#Else
    Dim hFile As Long
#End If
    
    Dim createDisposition As Long
    createDisposition = IIf(appendMode, OPEN_ALWAYS, CREATE_ALWAYS)
    
    hFile = CreateFileW(StrPtr(filePath), GENERIC_WRITE, FILE_SHARE_READ, _
                        0&, createDisposition, FILE_ATTRIBUTE_NORMAL, 0&)
    
    If hFile = INVALID_HANDLE_VALUE Then
        OpenFileForWrite = 0
        Exit Function
    End If
    
    If appendMode Then
        Dim fileSize As Currency, newPos As Currency
        If GetFileSizeEx(hFile, fileSize) <> 0 Then
            SetFilePointerEx hFile, fileSize, newPos, FILE_BEGIN
        End If
    End If
    
    OpenFileForWrite = hFile
End Function

#If VBA7 Then
Public Function WriteToFile(ByVal hFile As LongPtr, buffer() As Byte, _
                             ByVal bytesToWrite As Long) As Boolean
#Else
Public Function WriteToFile(ByVal hFile As Long, buffer() As Byte, _
                             ByVal bytesToWrite As Long) As Boolean
#End If
    If bytesToWrite <= 0 Then WriteToFile = True: Exit Function
    
    Dim written As Long, result As Long
    result = WriteFile(hFile, buffer(LBound(buffer)), bytesToWrite, written, 0&)
    WriteToFile = (result <> 0 And written = bytesToWrite)
End Function

#If VBA7 Then
Public Sub CloseFileHandle(ByVal hFile As LongPtr)
#Else
Public Sub CloseFileHandle(ByVal hFile As Long)
#End If
    If hFile <> 0 And hFile <> INVALID_HANDLE_VALUE Then
        On Error Resume Next
        FlushFileBuffers hFile
        On Error GoTo 0
        CloseHandle hFile
    End If
End Sub

