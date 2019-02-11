VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsByteBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByRef destination As Any, _
                                       ByRef source As Any, _
                                       ByVal length As Long)

Private data()                      As Byte

Private CurrentPos                  As Long
Private lastPos                     As Long
Private hFile                       As Long

Private Const INVALID_FILE_HANDLER  As Integer = -1
Private Const DEFAULT_MAX_SIZE_FILE As Long = 65535
Private Const BYTE_SIZE             As Byte = 1
Private Const BOOL_SIZE             As Byte = 2
Private Const INTEGER_SIZE          As Byte = 2
Private Const LONG_SIZE             As Byte = 4
Private Const SINGLE_SIZE           As Byte = 4
Private Const DOUBLE_SIZE           As Byte = 8
Private Const STRING_LENGTH_SIZE    As Byte = 2

Public Sub initializeReader(ByRef arrayByte() As Byte)
    
    On Error GoTo initializeReader_Err
    
    lastPos = UBound(arrayByte)
    ReDim data(lastPos) As Byte

    Call CopyMemory(data(0), arrayByte(0), lastPos + 1)
    CurrentPos = 0

    
    Exit Sub

initializeReader_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "initializeReader"
    End If
Resume Next
    
End Sub

Public Sub initializeWriter(ByVal fileHandler As Integer)
    
    On Error GoTo initializeWriter_Err
    
    ReDim data(DEFAULT_MAX_SIZE_FILE * 20) As Byte

    hFile = fileHandler
    CurrentPos = 0
    lastPos = -1

    
    Exit Sub

initializeWriter_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "initializeWriter"
    End If
Resume Next
    
End Sub

Public Sub getBytes(ByRef destination() As Byte, Optional ByVal length As Long = -1)
    
    On Error GoTo getBytes_Err
    

    If length >= 0 Then
        Call CopyMemory(destination(0), data(CurrentPos), length)
    Else
        Call CopyMemory(destination(0), data(0), lastPos + 1)

    End If

    
    Exit Sub

getBytes_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "getBytes"
    End If
Resume Next
    
End Sub

Public Function getByte() As Byte
    
    On Error GoTo getByte_Err
    
    getByte = data(CurrentPos)
    CurrentPos = CurrentPos + BYTE_SIZE

    
    Exit Function

getByte_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "getByte"
    End If
Resume Next
    
End Function

Public Function getBoolean() As Boolean
    
    On Error GoTo getBoolean_Err
    
    Call CopyMemory(getBoolean, data(CurrentPos), BOOL_SIZE)
    CurrentPos = CurrentPos + BOOL_SIZE

    
    Exit Function

getBoolean_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "getBoolean"
    End If
Resume Next
    
End Function

Public Function getInteger() As Integer
    
    On Error GoTo getInteger_Err
    
    Call CopyMemory(getInteger, data(CurrentPos), INTEGER_SIZE)
    CurrentPos = CurrentPos + INTEGER_SIZE

    
    Exit Function

getInteger_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "getInteger"
    End If
Resume Next
    
End Function

Public Function getLong() As Long
    
    On Error GoTo getLong_Err
    
    Call CopyMemory(getLong, data(CurrentPos), LONG_SIZE)
    CurrentPos = CurrentPos + LONG_SIZE

    
    Exit Function

getLong_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "getLong"
    End If
Resume Next
    
End Function

Public Function getSingle() As Single
    
    On Error GoTo getSingle_Err
    
    Call CopyMemory(getSingle, data(CurrentPos), SINGLE_SIZE)
    CurrentPos = CurrentPos + SINGLE_SIZE

    
    Exit Function

getSingle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "getSingle"
    End If
Resume Next
    
End Function

Public Function getDouble() As Double
    
    On Error GoTo getDouble_Err
    
    Call CopyMemory(getDouble, data(CurrentPos), DOUBLE_SIZE)
    CurrentPos = CurrentPos + DOUBLE_SIZE

    
    Exit Function

getDouble_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "getDouble"
    End If
Resume Next
    
End Function

Public Function getString(Optional ByVal length As Integer = -1) As String
    
    On Error GoTo getString_Err
    
    Dim ret() As Byte

    If length < 0 Then
        length = getInteger()
        getString = getString(length)
    Else

        If length > 0 Then
            ReDim ret(length - 1) As Byte
       
            Call CopyMemory(ret(0), data(CurrentPos), length)
       
            getString = StrConv(ret, vbUnicode)
            CurrentPos = CurrentPos + length

        End If

    End If

    
    Exit Function

getString_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "getString"
    End If
Resume Next
    
End Function

Public Sub putByte(ByVal value As Byte)
    
    On Error GoTo putByte_Err
    
    data(lastPos + 1) = value
    lastPos = lastPos + BYTE_SIZE

    
    Exit Sub

putByte_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "putByte"
    End If
Resume Next
    
End Sub

Public Sub putBoolean(ByVal value As Boolean)
    
    On Error GoTo putBoolean_Err
    
    Call CopyMemory(data(lastPos + 1), value, BOOL_SIZE)
    lastPos = lastPos + BOOL_SIZE

    
    Exit Sub

putBoolean_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "putBoolean"
    End If
Resume Next
    
End Sub

Public Sub putInteger(ByVal value As Integer)
    
    On Error GoTo putInteger_Err
    
    Call CopyMemory(data(lastPos + 1), value, INTEGER_SIZE)
    lastPos = lastPos + INTEGER_SIZE

    
    Exit Sub

putInteger_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "putInteger"
    End If
Resume Next
    
End Sub

Public Sub putLong(ByVal value As Long)
    
    On Error GoTo putLong_Err
    
    Call CopyMemory(data(lastPos + 1), value, LONG_SIZE)
    lastPos = lastPos + LONG_SIZE

    
    Exit Sub

putLong_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "putLong"
    End If
Resume Next
    
End Sub

Public Sub putSingle(ByVal value As Single)
    
    On Error GoTo putSingle_Err
    
    Call CopyMemory(data(lastPos + 1), value, SINGLE_SIZE)
    lastPos = lastPos + SINGLE_SIZE

    
    Exit Sub

putSingle_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "putSingle"
    End If
Resume Next
    
End Sub

Public Sub putDouble(ByVal value As Double)
    
    On Error GoTo putDouble_Err
    
    Call CopyMemory(data(lastPos + 1), value, DOUBLE_SIZE)
    lastPos = lastPos + DOUBLE_SIZE

    
    Exit Sub

putDouble_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "putDouble"
    End If
Resume Next
    
End Sub

Public Sub putString(ByRef str As String, Optional ByVal withLength As Boolean = True)
    
    On Error GoTo putString_Err
    
    Dim length As Long

    length = Len(str)

    If withLength Then
        Call putInteger(length)
        Call putString(str, False)
    Else

        If length > 0 Then
            Call CopyMemory(data(lastPos + 1), ByVal StrPtr(StrConv(str, vbFromUnicode)), length)
   
            lastPos = lastPos + length

        End If

    End If

    
    Exit Sub

putString_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "putString"
    End If
Resume Next
    
End Sub

Public Sub getVoid(ByVal length As Integer)
    
    On Error GoTo getVoid_Err
    
    CurrentPos = CurrentPos + length

    
    Exit Sub

getVoid_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "getVoid"
    End If
Resume Next
    
End Sub

Public Sub putVoid(ByVal length As Integer)
    
    On Error GoTo putVoid_Err
    
    lastPos = lastPos + length

    
    Exit Sub

putVoid_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "putVoid"
    End If
Resume Next
    
End Sub

Public Sub clearData()
    
    On Error GoTo clearData_Err
    
    ReDim data(DEFAULT_MAX_SIZE_FILE) As Byte

    CurrentPos = 0
    lastPos = -1
    hFile = -1

    
    Exit Sub

clearData_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "clearData"
    End If
Resume Next
    
End Sub

Public Function getLastPos() As Long
    
    On Error GoTo getLastPos_Err
    
    getLastPos = lastPos

    
    Exit Function

getLastPos_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "getLastPos"
    End If
Resume Next
    
End Function

Public Function getCurrentPos() As Long
    
    On Error GoTo getCurrentPos_Err
    
    getCurrentPos = CurrentPos

    
    Exit Function

getCurrentPos_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "getCurrentPos"
    End If
Resume Next
    
End Function

Public Function EOF() As Boolean
    
    On Error GoTo EOF_Err
    
    EOF = (CurrentPos > UBound(data))

    
    Exit Function

EOF_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "EOF"
    End If
Resume Next
    
End Function

Public Sub saveBuffer()
    
    On Error GoTo saveBuffer_Err
    
    Dim buf() As Byte

    If hFile > 0 Then
        ReDim buf(lastPos) As Byte
   
        Call CopyMemory(buf(0), data(0), lastPos + 1)
        Put hFile, , buf

    End If

    
    Exit Sub

saveBuffer_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "saveBuffer"
    End If
Resume Next
    
End Sub

Private Sub Class_Initialize()
    
    On Error GoTo Class_Initialize_Err
    
    hFile = INVALID_FILE_HANDLER

    
    Exit Sub

Class_Initialize_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "Class_Initialize"
    End If
Resume Next
    
End Sub

Private Sub Class_Terminate()
    
    On Error GoTo Class_Terminate_Err
    
    Erase data()

    
    Exit Sub

Class_Terminate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsByteBuffer" & "->" & "Class_Terminate"
    End If
Resume Next
    
End Sub

