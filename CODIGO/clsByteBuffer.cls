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

