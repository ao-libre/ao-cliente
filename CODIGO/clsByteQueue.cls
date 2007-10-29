VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsByteQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsByteQueue.cls - FIFO list of bytes.
' Creates and manipulates byte arrays to be sent and received by both client and server
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
' FIFO list of bytes
' Used to create and manipulate the byte arrays to be sent and received by both client and server
'
' @author Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.1.0
' @date 20060427

'**************************************************************************
' - HISTORY
'       v1.0.0  -   Initial release ( 2006/04/27 - Juan Martín Sotuyo Dodero )
'       v1.1.0  -   Added Single and Double support ( 2007/10/28 - Juan Martín Sotuyo Dodero )
'**************************************************************************

Option Explicit
Option Base 0       'It's the default, but we make it explicit just in case...

''
' The error number thrown when there is not enough data in
' the buffer to read the specified data type.
' It's 9 (subscript out of range) + the object error constant
Private Const NOT_ENOUGH_DATA As Long = vbObjectError + 9

''
' The error number thrown when there is not enough space in
' the buffer to write.
Private Const NOT_ENOUGH_SPACE As Long = vbObjectError + 10


''
' Default size of a data buffer (10 Kbs)
'
' @see Class_Initialize
Private Const DATA_BUFFER As Long = 10240

''
' The byte data
Dim data() As Byte

''
' How big the data array is
Dim queueCapacity As Long

''
' How far into the data array have we written
Dim queueLength As Long

''
' CopyMemory is the fastest way to copy memory blocks, so we abuse of it
'
' @param destination Where the data will be copied.
' @param source The data to be copied.
' @param length Number of bytes to be copied.

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)

''
' Initializes the queue with the default queueCapacity
'
' @see DATA_BUFFER

Private Sub Class_Initialize()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Initializes the queue with the default queueCapacity
'***************************************************
    ReDim data(DATA_BUFFER - 1) As Byte
    
    queueCapacity = DATA_BUFFER
End Sub

''
' Clean up and release resources

Private Sub Class_Terminate()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Clean up
'***************************************************
    Erase data
End Sub

''
' Copies another ByteQueue's data into this object.
'
' @param source The ByteQueue whose buffer will eb copied.
' @remarks  This method will resize the ByteQueue's buffer to match
'           the source. All previous data on this object will be lost.

Public Sub CopyBuffer(ByRef source As clsByteQueue)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'A Visual Basic equivalent of a Copy Contructor
'***************************************************
    If source.length = 0 Then
        'Clear the list and exit
        Call RemoveData(length)
        Exit Sub
    End If
    
    ' Set capacity and resize array - make sure all data is lost
    queueCapacity = source.Capacity
    
    ReDim data(queueCapacity - 1) As Byte
    
    ' Read buffer
    Dim buf() As Byte
    ReDim buf(source.length - 1) As Byte
    
    Call source.PeekBlock(buf, source.length)
    
    queueLength = 0
    
    ' Write buffer
    Call WriteBlock(buf, source.length)
End Sub

''
' Returns the smaller of val1 and val2
'
' @param val1 First value to compare
' @param val2 Second Value to compare
' @return   The smaller of val1 and val2
' @remarks  This method is faster than Iif() and cleaner, therefore it's used instead of it

Private Function min(ByVal val1 As Long, ByVal val2 As Long) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'It's faster than iif and I like it better
'***************************************************
    If val1 < val2 Then
        min = val1
    Else
        min = val2
    End If
End Function

''
' Writes a byte array at the end of the byte queue if there is enough space.
' Otherwise it throws NOT_ENOUGH_DATA.
'
' @param buf Byte array containing the data to be copied. MUST have 0 as the first index.
' @param datalength Total number of elements in the array
' @return   The actual number of bytes copied
' @remarks  buf MUST be Base 0
' @see RemoveData
' @see ReadData
' @see NOT_ENOUGH_DATA

Private Function WriteData(ByRef buf() As Byte, ByVal dataLength As Long) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'If the queueCapacity allows it copyes a byte buffer to the queue, if not it throws NOT_ENOUGH_DATA
'***************************************************
    'Check if there is enough free space
    If queueCapacity - queueLength - dataLength < 0 Then
        Call Err.raise(NOT_ENOUGH_SPACE)
        Exit Function
    End If
    
    'Copy data from buffer
    Call CopyMemory(data(queueLength), buf(0), dataLength)
    
    'Update length of data
    queueLength = queueLength + dataLength
    WriteData = dataLength
End Function

''
' Reads a byte array from the beginning of the byte queue if there is enough data available.
' Otherwise it throws NOT_ENOUGH_DATA.
'
' @param buf Byte array where to copy the data. MUST have 0 as the first index and already be sized properly.
' @param datalength Total number of elements in the array
' @return   The actual number of bytes copied
' @remarks  buf MUST be Base 0 and be already resized to be able to contain the requested bytes.
' This method performs no checks of such things as being a private method it's supposed that the consistency of the module is to be kept.
' If there is not enough data available it will read all available data.
' @see WriteData
' @see RemoveData
' @see NOT_ENOUGH_DATA

Private Function ReadData(ByRef buf() As Byte, ByVal dataLength As Long) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'If enough memory is available, it copies the requested number of bytes to the buffer
'***************************************************
    'Check if we can read the number of bytes requested
    If dataLength > queueLength Then
        Call Err.raise(NOT_ENOUGH_DATA)
        Exit Function
    End If
    
    'Copy data to buffer
    Call CopyMemory(buf(0), data(0), dataLength)
    ReadData = dataLength
End Function

''
' Removes a given number of bytes from the beginning of the byte queue.
' If there is less data available than the requested amount it removes all data.
'
' @param datalength Total number of bytes to remove
' @return   The actual number of bytes removed
' @see WriteData
' @see ReadData

Private Function RemoveData(ByVal dataLength As Long) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Moves forward the queue overwriting the first dataLength bytes
'***************************************************
    'Figure out how many bytes we can remove
    RemoveData = min(dataLength, queueLength)
    
    'Remove data - prevent rt9 when cleaning a full queue
    If RemoveData <> queueCapacity Then _
        Call CopyMemory(data(0), data(RemoveData), queueLength - RemoveData)
    
    'Update length
    queueLength = queueLength - RemoveData
End Function

''
' Writes a single byte at the end of the queue
'
' @param value The value to be written
' @return   The number of bytes written
' @see PeekByte
' @see ReadByte

Public Function WriteByte(ByVal value As Byte) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Writes a byte to the queue
'***************************************************
    Dim buf(0) As Byte
    
    buf(0) = value
    
    WriteByte = WriteData(buf, 1)
End Function

''
' Writes an integer at the end of the queue
'
' @param value The value to be written
' @return   The number of bytes written
' @see PeekInteger
' @see ReadInteger

Public Function WriteInteger(ByVal value As Integer) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Writes an integer to the queue
'***************************************************
    Dim buf(1) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), value, 2)
    
    WriteInteger = WriteData(buf, 2)
End Function

''
' Writes a long at the end of the queue
'
' @param value The value to be written
' @return   The number of bytes written
' @see PeekLong
' @see ReadLong

Public Function WriteLong(ByVal value As Long) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Writes a long to the queue
'***************************************************
    Dim buf(3) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), value, 4)
    
    WriteLong = WriteData(buf, 4)
End Function

''
' Writes a single at the end of the queue
'
' @param value The value to be written
' @return   The number of bytes written
' @see PeekSingle
' @see ReadSingle

Public Function WriteSingle(ByVal value As Single) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/28/07
'Writes a single to the queue
'***************************************************
    Dim buf(3) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), value, 4)
    
    WriteSingle = WriteData(buf, 4)
End Function

''
' Writes a double at the end of the queue
'
' @param value The value to be written
' @return   The number of bytes written
' @see PeekDouble
' @see ReadDouble

Public Function WriteDouble(ByVal value As Double) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/28/07
'Writes a double to the queue
'***************************************************
    Dim buf(7) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), value, 8)
    
    WriteDouble = WriteData(buf, 8)
End Function

''
' Writes a boolean value at the end of the queue
'
' @param value The value to be written
' @return   The number of bytes written
' @see PeekBoolean
' @see ReadBoolean

Public Function WriteBoolean(ByVal value As Boolean) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Writes a byte to the queue
'***************************************************
    Dim buf(0) As Byte
    
    If value Then buf(0) = 1
    
    WriteBoolean = WriteData(buf, 1)
End Function

''
' Writes a fixed length ASCII string at the end of the queue
'
' @param value The string to be written
' @return   The number of bytes written
' @see PeekASCIIStringFixed
' @see ReadASCIIStringFixed

Public Function WriteASCIIStringFixed(ByVal value As String) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Writes a fixed length ASCII string to the queue
'***************************************************
    Dim buf() As Byte
    ReDim buf(Len(value) - 1) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), ByVal StrPtr(StrConv(value, vbFromUnicode)), Len(value))
    
    WriteASCIIStringFixed = WriteData(buf, Len(value))
End Function

''
' Writes a fixed length unicode string at the end of the queue
'
' @param value The string to be written
' @return   The number of bytes written
' @see PeekUnicodeStringFixed
' @see ReadUnicodeStringFixed

Public Function WriteUnicodeStringFixed(ByVal value As String) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Writes a fixed length UNICODE string to the queue
'***************************************************
    Dim buf() As Byte
    ReDim buf(LenB(value)) As Byte
    
    'Copy data to temp buffer
    Call CopyMemory(buf(0), ByVal StrPtr(value), LenB(value))
    
    WriteUnicodeStringFixed = WriteData(buf, LenB(value))
End Function

''
' Writes a variable length ASCII string at the end of the queue
'
' @param value The string to be written
' @return   The number of bytes written
' @see PeekASCIIString
' @see ReadASCIIString

Public Function WriteASCIIString(ByVal value As String) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Writes a variable length ASCII string to the queue
'***************************************************
    Dim buf() As Byte
    ReDim buf(Len(value) + 1) As Byte
    
    'Copy length to temp buffer
    Call CopyMemory(buf(0), CInt(Len(value)), 2)
    
    If Len(value) > 0 Then
        'Copy data to temp buffer
        Call CopyMemory(buf(2), ByVal StrPtr(StrConv(value, vbFromUnicode)), Len(value))
    End If
    
    WriteASCIIString = WriteData(buf, Len(value) + 2)
End Function

''
' Writes a variable length unicode string at the end of the queue
'
' @param value The string to be written
' @return   The number of bytes written
' @see PeekUnicodeString
' @see ReadUnicodeString

Public Function WriteUnicodeString(ByVal value As String) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Writes a variable length UNICODE string to the queue
'***************************************************
    Dim buf() As Byte
    ReDim buf(LenB(value) + 1) As Byte
    
    'Copy length to temp buffer
    Call CopyMemory(buf(0), CInt(Len(value)), 2)
    
    If Len(value) > 0 Then
        'Copy data to temp buffer
        Call CopyMemory(buf(2), ByVal StrPtr(value), LenB(value))
    End If
    
    WriteUnicodeString = WriteData(buf, LenB(value) + 2)
End Function

''
' Writes a byte array at the end of the queue
'
' @param value The byte array to be written. MUST be Base 0.
' @param length The number of elements to copy from the byte array. If less than 0 it will copy the whole array.
' @return   The number of bytes written
' @remarks  value() MUST be Base 0.
' @see PeekBlock
' @see ReadBlock

Public Function WriteBlock(ByRef value() As Byte, Optional ByVal length As Long = -1) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Writes a byte array to the queue
'***************************************************
    'Prevent from copying memory outside the array
    If length > UBound(value()) + 1 Or length < 0 Then length = UBound(value()) + 1
    
    WriteBlock = WriteData(value, length)
End Function

''
' Reads a single byte from the begining of the queue and removes it
'
' @return   The read value
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekByte
' @see WriteByte

Public Function ReadByte() As Byte
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a byte from the queue and removes it
'***************************************************
    Dim buf(0) As Byte
    
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 1))
    
    ReadByte = buf(0)
End Function

''
' Reads an integer from the begining of the queue and removes it
'
' @return   The read value
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekInteger
' @see WriteInteger

Public Function ReadInteger() As Integer
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads an integer from the queue and removes it
'***************************************************
    Dim buf(1) As Byte
    
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 2))
    
    'Copy data to temp buffer
    Call CopyMemory(ReadInteger, buf(0), 2)
End Function

''
' Reads a long from the begining of the queue and removes it
'
' @return   The read value
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekLong
' @see WriteLong

Public Function ReadLong() As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a long from the queue and removes it
'***************************************************
    Dim buf(3) As Byte
    
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 4))
    
    'Copy data to temp buffer
    Call CopyMemory(ReadLong, buf(0), 4)
End Function

''
' Reads a single from the begining of the queue and removes it
'
' @return   The read value
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekSingle
' @see WriteSingle

Public Function ReadSingle() As Single
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/28/07
'Reads a single from the queue and removes it
'***************************************************
    Dim buf(3) As Byte
    
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 4))
    
    'Copy data to temp buffer
    Call CopyMemory(ReadSingle, buf(0), 4)
End Function

''
' Reads a double from the begining of the queue and removes it
'
' @return   The read value
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekDouble
' @see WriteDouble

Public Function ReadDouble() As Double
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/28/07
'Reads a double from the queue and removes it
'***************************************************
    Dim buf(7) As Byte
    
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 8))
    
    'Copy data to temp buffer
    Call CopyMemory(ReadDouble, buf(0), 8)
End Function

''
' Reads a Boolean from the begining of the queue and removes it
'
' @return   The read value
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekBoolean
' @see WriteBoolean

Public Function ReadBoolean() As Boolean
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a Boolean from the queue and removes it
'***************************************************
    Dim buf(0) As Byte
    
    'Read the data and remove it
    Call RemoveData(ReadData(buf, 1))
    
    If buf(0) = 1 Then ReadBoolean = True
End Function

''
' Reads a fixed length ASCII string from the begining of the queue and removes it
'
' @param length The length of the string to be read
' @return   The read string
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' If there is not enough data to read the complete string then nothing is removed and an empty string is returned
' @see PeekASCIIStringFixed
' @see WriteUnicodeStringFixed

Public Function ReadASCIIStringFixed(ByVal length As Long) As String
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a fixed length ASCII string from the queue and removes it
'***************************************************
    If length <= 0 Then Exit Function
    
    If queueLength >= length Then
        Dim buf() As Byte
        ReDim buf(length - 1) As Byte
        
        'Read the data and remove it
        Call RemoveData(ReadData(buf, length))
        
        ReadASCIIStringFixed = StrConv(buf, vbUnicode)
    Else
        Call Err.raise(NOT_ENOUGH_DATA)
    End If
End Function

''
' Reads a fixed length unicode string from the begining of the queue and removes it
'
' @param length The length of the string to be read.
' @return   The read string if enough data is available, an empty string otherwise.
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way.
' If there is not enough data to read the complete string then nothing is removed and an empty string is returned
' @see PeekUnicodeStringFixed
' @see WriteUnicodeStringFixed

Public Function ReadUnicodeStringFixed(ByVal length As Long) As String
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a fixed length UNICODE string from the queue and removes it
'***************************************************
    If length <= 0 Then Exit Function
    
    If queueLength >= length * 2 Then
        Dim buf() As Byte
        ReDim buf(length * 2 - 1) As Byte
        
        'Read the data and remove it
        Call RemoveData(ReadData(buf, length * 2))
        
        ReadUnicodeStringFixed = buf
    Else
        Call Err.raise(NOT_ENOUGH_DATA)
    End If
End Function

''
' Reads a variable length ASCII string from the begining of the queue and removes it
'
' @return   The read string
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' If there is not enough data to read the complete string then nothing is removed and an empty string is returned
' @see PeekASCIIString
' @see WriteASCIIString

Public Function ReadASCIIString() As String
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a variable length ASCII string from the queue and removes it
'***************************************************
    Dim buf(1) As Byte
    Dim length As Integer
    
    'Make sure we can read a valid length
    If queueLength > 1 Then
        'Read the length
        Call ReadData(buf, 2)
        Call CopyMemory(length, buf(0), 2)
        
        'Make sure there are enough bytes
        If queueLength >= length + 2 Then
            'Remove the length
            Call RemoveData(2)
            
            If length > 0 Then
                Dim buf2() As Byte
                ReDim buf2(length - 1) As Byte
                
                
                'Read the data and remove it
                Call RemoveData(ReadData(buf2, length))
                
                ReadASCIIString = StrConv(buf2, vbUnicode)
            End If
        Else
            Call Err.raise(NOT_ENOUGH_DATA)
        End If
    Else
        Call Err.raise(NOT_ENOUGH_DATA)
    End If
End Function

''
' Reads a variable length unicode string from the begining of the queue and removes it
'
' @return   The read string if enough data is available, an empty string otherwise.
' @remarks  Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' If there is not enough data to read the complete string then nothing is removed and an empty string is returned
' @see PeekUnicodeString
' @see WriteUnicodeString

Public Function ReadUnicodeString() As String
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a variable length UNICODE string from the queue and removes it
'***************************************************
    Dim buf(1) As Byte
    Dim length As Integer
    
    'Make sure we can read a valid length
    If queueLength > 1 Then
        'Read the length
        Call ReadData(buf, 2)
        Call CopyMemory(length, buf(0), 2)
        
        'Make sure there are enough bytes
        If queueLength >= length * 2 + 2 Then
            'Remove the length
            Call RemoveData(2)
            
            Dim buf2() As Byte
            ReDim buf2(length * 2 - 1) As Byte
            
            'Read the data and remove it
            Call RemoveData(ReadData(buf2, length * 2))
            
            ReadUnicodeString = buf2
        Else
            Call Err.raise(NOT_ENOUGH_DATA)
        End If
    Else
        Call Err.raise(NOT_ENOUGH_DATA)
    End If
End Function

''
' Reads a byte array from the begining of the queue and removes it
'
' @param block Byte array which will contain the read data. MUST be Base 0 and previously resized to contain the requested amount of bytes.
' @param dataLength Number of bytes to retrieve from the queue.
' @return   The number of read bytes.
' @remarks  The block() array MUST be Base 0 and previously resized to be able to contain the requested bytes.
' Read methods removes the data from the queue.
' Data removed can't be recovered by the queue in any way
' @see PeekBlock
' @see WriteBlock

Public Function ReadBlock(ByRef block() As Byte, ByVal dataLength As Long) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a byte array from the queue and removes it
'***************************************************
    'Read the data and remove it
    If dataLength > 0 Then _
        ReadBlock = RemoveData(ReadData(block(), dataLength))
End Function

''
' Reads a single byte from the begining of the queue but DOES NOT remove it.
'
' @return   The read value.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadByte
' @see WriteByte

Public Function PeekByte() As Byte
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a byte from the queue but doesn't removes it
'***************************************************
    Dim buf(0) As Byte
    
    'Read the data and remove it
    Call ReadData(buf, 1)
    
    PeekByte = buf(0)
End Function

''
' Reads an integer from the begining of the queue but DOES NOT remove it.
'
' @return   The read value.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadInteger
' @see WriteInteger

Public Function PeekInteger() As Integer
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads an integer from the queue but doesn't removes it
'***************************************************
    Dim buf(1) As Byte
    
    'Read the data and remove it
    Call ReadData(buf, 2)
    
    'Copy data to temp buffer
    Call CopyMemory(PeekInteger, buf(0), 2)
End Function

''
' Reads a long from the begining of the queue but DOES NOT remove it.
'
' @return   The read value.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadLong
' @see WriteLong

Public Function PeekLong() As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a long from the queue but doesn't removes it
'***************************************************
    Dim buf(3) As Byte
    
    'Read the data and remove it
    Call ReadData(buf, 4)
    
    'Copy data to temp buffer
    Call CopyMemory(PeekLong, buf(0), 4)
End Function

''
' Reads a single from the begining of the queue but DOES NOT remove it.
'
' @return   The read value.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadSingle
' @see WriteSingle

Public Function PeekSingle() As Single
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/28/07
'Reads a single from the queue but doesn't removes it
'***************************************************
    Dim buf(3) As Byte
    
    'Read the data and remove it
    Call ReadData(buf, 4)
    
    'Copy data to temp buffer
    Call CopyMemory(PeekSingle, buf(0), 4)
End Function

''
' Reads a double from the begining of the queue but DOES NOT remove it.
'
' @return   The read value.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadDouble
' @see WriteDouble

Public Function PeekDouble() As Double
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/28/07
'Reads a double from the queue but doesn't removes it
'***************************************************
    Dim buf(7) As Byte
    
    'Read the data and remove it
    Call ReadData(buf, 8)
    
    'Copy data to temp buffer
    Call CopyMemory(PeekDouble, buf(0), 8)
End Function

''
' Reads a Bollean from the begining of the queue but DOES NOT remove it.
'
' @return   The read value.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadBoolean
' @see WriteBoolean

Public Function PeekBoolean() As Boolean
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a Boolean from the queue but doesn't removes it
'***************************************************
    Dim buf(0) As Byte
    
    'Read the data and remove it
    Call ReadData(buf, 1)
    
    If buf(0) = 1 Then PeekBoolean = True
End Function

''
' Reads a fixed length ASCII string from the begining of the queue but DOES NOT remove it.
'
' @param length The length of the string to be read
' @return   The read string if enough data is available, an empty string otherwise.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' If there is not enough data to read the complete string then an empty string is returned
' @see ReadASCIIStringFixed
' @see WriteASCIIStringFixed

Public Function PeekASCIIStringFixed(ByVal length As Long) As String
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a fixed length ASCII string from the queue but doesn't removes it
'***************************************************
    If length <= 0 Then Exit Function
    
    If queueLength >= length Then
        Dim buf() As Byte
        ReDim buf(length - 1) As Byte
        
        'Read the data and remove it
        Call ReadData(buf, length)
        
        PeekASCIIStringFixed = StrConv(buf, vbUnicode)
    Else
        Call Err.raise(NOT_ENOUGH_DATA)
    End If
End Function

''
' Reads a fixed length unicode string from the begining of the queue but DOES NOT remove it.
'
' @param length The length of the string to be read
' @return   The read string if enough data is available, an empty string otherwise.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' If there is not enough data to read the complete string then an empty string is returned
' @see ReadUnicodeStringFixed
' @see WriteUnicodeStringFixed

Public Function PeekUnicodeStringFixed(ByVal length As Long) As String
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a fixed length UNICODE string from the queue but doesn't removes it
'***************************************************
    If length <= 0 Then Exit Function
    
    If queueLength >= length * 2 Then
        Dim buf() As Byte
        ReDim buf(length * 2 - 1) As Byte
        
        'Read the data and remove it
        Call ReadData(buf, length * 2)
        
        PeekUnicodeStringFixed = buf
    Else
        Call Err.raise(NOT_ENOUGH_DATA)
    End If
End Function

''
' Reads a variable length ASCII string from the begining of the queue but DOES NOT remove it.
'
' @return   The read string if enough data is available, an empty string otherwise.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' If there is not enough data to read the complete string then an empty string is returned
' @see ReadASCIIString
' @see WriteASCIIString

Public Function PeekASCIIString() As String
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a variable length ASCII string from the queue but doesn't removes it
'***************************************************
    Dim buf(1) As Byte
    Dim length As Integer
    
    'Make sure we can read a valid length
    If queueLength > 1 Then
        'Read the length
        Call ReadData(buf, 2)
        Call CopyMemory(length, buf(0), 2)
        
        'Make sure there are enough bytes
        If queueLength >= length + 2 Then
            Dim buf2() As Byte
            ReDim buf2(length + 1) As Byte
            
            'Read the data (we have to read the length again)
            Call ReadData(buf2, length + 2)
            
            If length > 0 Then
                'Get rid of the length
                Dim buf3() As Byte
                ReDim buf3(length - 1) As Byte
                Call CopyMemory(buf3(0), buf2(2), length)
                
                PeekASCIIString = StrConv(buf3, vbUnicode)
            End If
        Else
            Call Err.raise(NOT_ENOUGH_DATA)
        End If
    Else
        Call Err.raise(NOT_ENOUGH_DATA)
    End If
End Function

''
' Reads a variable length unicode string from the begining of the queue but DOES NOT remove it.
'
' @return   The read string if enough data is available, an empty string otherwise.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' If there is not enough data to read the complete string then an empty string is returned
' @see ReadUnicodeString
' @see WriteUnicodeString

Public Function PeekUnicodeString() As String
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a variable length UNICODE string from the queue but doesn't removes it
'***************************************************
    Dim buf(1) As Byte
    Dim length As Integer
    
    'Make sure we can read a valid length
    If queueLength > 1 Then
        'Read the length
        Call ReadData(buf, 2)
        Call CopyMemory(length, buf(0), 2)
        
        'Make sure there are enough bytes
        If queueLength >= length * 2 + 2 Then
            Dim buf2() As Byte
            ReDim buf2(length * 2 + 1) As Byte
            
            'Read the data (we need to read the length again)
            Call ReadData(buf2, length * 2 + 2)
            
            'Get rid of the length bytes
            Dim buf3() As Byte
            ReDim buf3(length * 2 - 1) As Byte
            Call CopyMemory(buf3(0), buf2(2), length * 2)
            
            PeekUnicodeString = buf3
        Else
            Call Err.raise(NOT_ENOUGH_DATA)
        End If
    Else
        Call Err.raise(NOT_ENOUGH_DATA)
    End If
End Function

''
' Reads a byte array from the begining of the queue but DOES NOT remove it.
'
' @param block() Byte array that will contain the read data. MUST be Base 0 and previously resized to contain the requested amount of bytes.
' @param dataLength Number of bytes to be read
' @return   The actual number of read bytes.
' @remarks  Peek methods, unlike Read methods, don't remove the data from the queue.
' @see ReadBlock
' @see WriteBlock

Public Function PeekBlock(ByRef block() As Byte, ByVal dataLength As Long) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Reads a byte array from the queue but doesn't removes it
'***************************************************
    'Read the data
    If dataLength > 0 Then _
        PeekBlock = ReadData(block(), dataLength)
End Function

''
' Retrieves the current capacity of the queue.
'
' @return   The current capacity of the queue.

Public Property Get Capacity() As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Retrieves the current capacity of the queue
'***************************************************
    Capacity = queueCapacity
End Property

''
' Sets the capacity of the queue.
'
' @param value The new capacity of the queue.
' @remarks If the new capacity is smaller than the current Length, all exceeding data is lost.
' @see Length

Public Property Let Capacity(ByVal value As Long)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Sets the current capacity of the queue.
'All data in the queue exceeding the new capacity is lost
'***************************************************
    'Upate capacity
    queueCapacity = value
    
    'All extra data is lost
    If length > value Then queueLength = value
    
    'Resize the queue
    ReDim Preserve data(queueCapacity - 1) As Byte
End Property

''
' Retrieves the length of the total data in the queue.
'
' @return   The length of the total data in the queue.

Public Property Get length() As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Retrieves the current number of bytes in the queue
'***************************************************
    length = queueLength
End Property

''
' Retrieves the NOT_ENOUGH_DATA error code.
'
' @return   NOT_ENOUGH_DATA.

Public Property Get NotEnoughDataErrCode() As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Retrieves the NOT_ENOUGH_DATA error code
'***************************************************
    NotEnoughDataErrCode = NOT_ENOUGH_DATA
End Property

''
' Retrieves the NOT_ENOUGH_SPACE error code.
'
' @return   NOT_ENOUGH_SPACE.

Public Property Get NotEnoughSpaceErrCode() As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/27/06
'Retrieves the NOT_ENOUGH_SPACE error code
'***************************************************
    NotEnoughSpaceErrCode = NOT_ENOUGH_SPACE
End Property

