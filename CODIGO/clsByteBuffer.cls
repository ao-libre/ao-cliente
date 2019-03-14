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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                                    ByRef destination As Any, _
                                    ByRef source As Any, _
                                    ByVal length As Long)

Private data() As Byte

Private CurrentPos As Long
Private lastPos As Long
Private hFile As Long

Private Const INVALID_FILE_HANDLER As Integer = -1
Private Const DEFAULT_MAX_SIZE_FILE As Long = 65535
Private Const BYTE_SIZE As Byte = 1
Private Const BOOL_SIZE As Byte = 2
Private Const INTEGER_SIZE As Byte = 2
Private Const LONG_SIZE As Byte = 4
Private Const SINGLE_SIZE As Byte = 4
Private Const DOUBLE_SIZE As Byte = 8

Public Sub initializeReader(ByRef arrayByte() As Byte)
lastPos = UBound(arrayByte)
ReDim data(lastPos) As Byte

Call CopyMemory(data(0), arrayByte(0), lastPos + 1)
CurrentPos = 0
End Sub

Public Sub initializeWriter(ByVal fileHandler As Integer)
    ReDim data(DEFAULT_MAX_SIZE_FILE * 20) As Byte

    hFile = fileHandler
    CurrentPos = 0
    lastPos = -1
End Sub

Public Sub getBytes(ByRef destination() As Byte, Optional ByVal length As Long = -1)
If length >= 0 Then
    Call CopyMemory(destination(0), data(CurrentPos), length)
Else
    Call CopyMemory(destination(0), data(0), lastPos + 1)
End If
End Sub

Public Function getByte() As Byte
getByte = data(CurrentPos)
CurrentPos = CurrentPos + BYTE_SIZE
End Function

Public Function getBoolean() As Boolean
Call CopyMemory(getBoolean, data(CurrentPos), BOOL_SIZE)
CurrentPos = CurrentPos + BOOL_SIZE
End Function

Public Function getInteger() As Integer
Call CopyMemory(getInteger, data(CurrentPos), INTEGER_SIZE)
CurrentPos = CurrentPos + INTEGER_SIZE
End Function

Public Function getLong() As Long
Call CopyMemory(getLong, data(CurrentPos), LONG_SIZE)
CurrentPos = CurrentPos + LONG_SIZE
End Function

Public Function getSingle() As Single
Call CopyMemory(getSingle, data(CurrentPos), SINGLE_SIZE)
CurrentPos = CurrentPos + SINGLE_SIZE
End Function

Public Function getDouble() As Double
Call CopyMemory(getDouble, data(CurrentPos), DOUBLE_SIZE)
CurrentPos = CurrentPos + DOUBLE_SIZE
End Function

Public Function getString(Optional ByVal length As Integer = -1) As String
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

End Function

Public Sub putByte(ByVal value As Byte)
data(lastPos + 1) = value
lastPos = lastPos + BYTE_SIZE
End Sub

Public Sub putBoolean(ByVal value As Boolean)
Call CopyMemory(data(lastPos + 1), value, BOOL_SIZE)
lastPos = lastPos + BOOL_SIZE
End Sub

Public Sub putInteger(ByVal value As Integer)
Call CopyMemory(data(lastPos + 1), value, INTEGER_SIZE)
lastPos = lastPos + INTEGER_SIZE
End Sub

Public Sub putLong(ByVal value As Long)
Call CopyMemory(data(lastPos + 1), value, LONG_SIZE)
lastPos = lastPos + LONG_SIZE
End Sub

Public Sub putSingle(ByVal value As Single)
Call CopyMemory(data(lastPos + 1), value, SINGLE_SIZE)
lastPos = lastPos + SINGLE_SIZE
End Sub

Public Sub putDouble(ByVal value As Double)
Call CopyMemory(data(lastPos + 1), value, DOUBLE_SIZE)
lastPos = lastPos + DOUBLE_SIZE
End Sub

Public Sub putString(ByRef str As String, Optional ByVal withLength As Boolean = True)
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
End Sub

Public Sub getVoid(ByVal length As Integer)
CurrentPos = CurrentPos + length
End Sub

Public Sub putVoid(ByVal length As Integer)
lastPos = lastPos + length
End Sub

Public Sub clearData()
ReDim data(DEFAULT_MAX_SIZE_FILE) As Byte

CurrentPos = 0
lastPos = -1
hFile = -1
End Sub

Public Function getLastPos() As Long
getLastPos = lastPos
End Function

Public Function getCurrentPos() As Long
getCurrentPos = CurrentPos
End Function

Public Function EOF() As Boolean
EOF = (CurrentPos > UBound(data))
End Function

Public Sub saveBuffer()
    Dim buf() As Byte

    If hFile > 0 Then
        ReDim buf(lastPos) As Byte
   
        Call CopyMemory(buf(0), data(0), lastPos + 1)
        Put hFile, , buf
    End If
End Sub

Private Sub Class_Initialize()
hFile = INVALID_FILE_HANDLER
End Sub

Private Sub Class_Terminate()
On Error Resume Next

    Erase data()
End Sub


