VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This file is part of Argentum Online, licensed under the APACHE License.
'
' Copyright (c) 2017 Argentum-Dev <http://www.argentum-dev.com>
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
' http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
Option Explicit

Private Declare Sub GetMem2 _
    Lib "msvbvm60" (ByRef Pointer As Any, ByRef Value As Any)
Private Declare Sub GetMem4 _
    Lib "msvbvm60" (ByRef Pointer As Any, ByRef Value As Any)
Private Declare Sub GetMem8 _
    Lib "msvbvm60" (ByRef Pointer As Any, ByRef Value As Any)
Private Declare Sub PutMem2 _
    Lib "msvbvm60" (ByRef Pointer As Any, ByVal Value As Integer)
Private Declare Sub PutMem4 _
    Lib "msvbvm60" (ByRef Pointer As Any, ByVal Value As Any)
Private Declare Sub PutMem8 _
    Lib "msvbvm60" (ByRef Pointer As Any, ByRef Value As Any)
Private Declare Sub RtlMoveMemory _
    Lib "kernel32" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

''
' The error number thrown when there is not enough data in  the buffer to read the specified data type.
''
Private Const NOT_ENOUGH_DATA  As Long = vbObjectError + 9

''
' The error number thrown when there is not enough space in the buffer to write.
''
Private Const NOT_ENOUGH_SPACE As Long = vbObjectError + 10

''
' Default capacity of the buffer (8 kbytes)
'
' @see Class_Initialize
''
Private Const DEFAULT_CAPACITY As Long = 8192

''
' The underlying byte array of the buffer
''
Private mData()                As Byte

''
' The current position within the byte array
''
Private mPosition              As Long

''
' The current limit of the byte array
''
Private mLimit                 As Long

''
' The current marked position
''
Private mMark                  As Long

''
' \brief  Constructor of the class.
''
Private Sub Class_Initialize()

    ReDim mData(0 To DEFAULT_CAPACITY - 1) As Byte
    
    mLimit = DEFAULT_CAPACITY
    
End Sub

''
' \brief  Destructor of the class.
''
Private Sub Class_Terminate()

    Erase mData
    
End Sub

''
' \brief  Clear the buffer.
'
' \note   position will be set to zero and length to capacity.
''
Public Sub Clear()
    
    mLimit = UBound(mData) + 1
    
    mPosition = 0: mMark = 0
    
End Sub

''
' \brief  Flip the buffer.
'
' \note   position will be ser to zero and length to position.
''
Public Sub Flip()

    mLimit = mPosition

    mPosition = 0: mMark = 0
    
End Sub

''
' \brief  Mark the buffer.
'
' \see    #Reset()
''
Public Sub Mark()

    mMark = mPosition
    
End Sub

''
' \brief  Reset the buffer.
'
' \see    #Mark()
''
Public Sub Reset()

    mPosition = mMark
    
End Sub

''
' \brief  Skip an arbitrary number of bytes from the buffer
'
' \param[in] Bytes      - The number of bytes to skip from the buffer
''
Public Sub Skip(ByVal Bytes As Long)

    If (mPosition + Bytes > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
            
        mPosition = mPosition + Bytes
        
    End If
    
End Sub

''
' \brief  Read a 8-bit integer from the buffer without increasing its position.
'
' \see    #ReadByte()
''
Public Function PeekByte() As Byte
    
    If (mPosition + &H1 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
            
        PeekByte = mData(mPosition)
        
    End If
    
End Function

''
' \brief  Read a 16-bit integer from the buffer without increasing its position.
'
' \see    #ReadInteger()
''
Public Function PeekInteger() As Byte
    
    If (mPosition + &H2 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
            
        Call GetMem2(mData(mPosition), PeekInteger)

    End If
    
End Function

''
' \brief  Read a 32-bit integer from the buffer without increasing its position.
'
' \see    #ReadLong()
''
Public Function PeekLong() As Byte
    
    If (mPosition + &H4 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
            
        Call GetMem4(mData(mPosition), PeekLong)

    End If
    
End Function

''
' \brief  Peek a boolean from the buffer.
'
' \see    #ReadBoolean()
''
Public Function PeekBoolean() As Boolean
    
    If (mPosition + &H1 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
            
        PeekBoolean = mData(mPosition)

    End If

End Function

''
' \brief  Read a 8-bit integer from the buffer.
'
' \see    #PeekByte()
''
Public Function ReadByte() As Byte
    
    If (mPosition + &H1 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
            
        ReadByte = mData(mPosition)
        
        mPosition = mPosition + &H1
        
    End If

End Function

''
' \brief  Read an arbitrary number of bytes from the buffer.
'
' \param[in] Values     - The array that will contain the bytes
' \param[in] Length     - The number of bytes to read (must be less than \ref Values capacity)
''
Public Function ReadBytes(ByRef Values() As Byte, Optional ByVal Length As Long = &HFFFFFFFF)
    
    If (Length = &HFFFFFFFF) Then
    
        Length = UBound(Values) + 1

    End If
    
    If (mPosition + Length > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
        
        Call RtlMoveMemory(Values(0), mData(mPosition), Length)
                
        mPosition = mPosition + Length
        
    End If

End Function

''
' \brief  Read a boolean from the buffer.
'
' \see    #PeekBoolean()
''
Public Function ReadBoolean() As Boolean
    
    If (mPosition + &H1 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
            
        ReadBoolean = mData(mPosition)
        
        mPosition = mPosition + &H1
        
    End If

End Function

''
' \brief  Read a 16-bit integer from the buffer.
'
' \see    #PeekInteger()
''
Public Function ReadInteger() As Integer
        
    If (mPosition + &H2 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
            
        Call GetMem2(mData(mPosition), ReadInteger)
        
        mPosition = mPosition + &H2
        
    End If

End Function

''
' \brief  Read a 32-bit integer from the buffer.
'
' \see    #PeekLong()
''
Public Function ReadLong() As Long
            
    If (mPosition + &H4 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
            
        Call GetMem4(mData(mPosition), ReadLong)
        
        mPosition = mPosition + &H4
        
    End If

End Function

''
' \brief  Read a 64-bit integer from the buffer.
''
Public Function ReadCurrency() As Currency
                
    If (mPosition + &H8 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
            
        Call GetMem8(mData(mPosition), ReadCurrency)
        
        mPosition = mPosition + &H8
        
    End If

End Function

''
' \brief  Read a 32-bit floating point number from the buffer.
''
Public Function ReadSingle() As Single
                    
    If (mPosition + &H4 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
                
        Call GetMem4(mData(mPosition), ReadSingle)
        
        mPosition = mPosition + &H4
        
    End If
    
End Function

''
' \brief  Read a 64-bit floating point number from the buffer.
''
Public Function ReadDouble() As Double
                    
    If (mPosition + &H8 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_DATA)

    Else
                
        Call GetMem8(mData(mPosition), ReadDouble)
        
        mPosition = mPosition + &H8
    
    End If
    
End Function

''
' \brief  Read a string from the buffer.
'
' \param[in, opt] Length    - The number of characters to read, or retrieve it from the buffer itself
''
Public Function ReadString(Optional ByVal Length As Long = &HFFFFFFFF) As String

    If (Length = &HFFFFFFFF) Then
    
        Length = ReadInteger()
        
    End If
    
    If (Length > 0) Then
        
        If (mPosition + Length > mLimit) Then
            
            Call Err.Raise(NOT_ENOUGH_DATA)
    
        Else
                
            Dim Bytes() As Byte
            
            ReDim Bytes(0 To Length - 1) As Byte
            
            Call ReadBytes(Bytes, Length)
            
            ReadString = StrConv(Bytes, vbUnicode)
            
        End If
    
    End If
    
End Function

''
' \brief  Read an unicode string from the buffer.
'
' \param[in, opt] Length    - The number of characters to read, or retrieve it from the buffer itself
''
Public Function ReadUnicode(Optional ByVal Length As Long = &HFFFFFFFF) As String

    If (Length = &HFFFFFFFF) Then
    
        Length = ReadInteger()
        
    End If
  
    If (Length > 0) Then
               
        Length = Length * &H2
           
        If (mPosition + Length > mLimit) Then
            
            Call Err.Raise(NOT_ENOUGH_DATA)
    
        Else
            
            Dim Bytes() As Byte
            
            ReDim Bytes(0 To Length - 1) As Byte
            
            Call ReadBytes(Bytes, Length)
            
            ReadUnicode = Bytes
            
        End If
    
    End If
    
End Function

''
' \brief  Write a 8-bit integer into the buffer.
'
' \param[in] Value      - The value to write into the buffer
''
Public Sub WriteByte(ByVal Value As Byte)
    
    If (mPosition + &H1 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_SPACE)

    Else
        
        mData(mPosition) = Value
        
        mPosition = mPosition + &H1
        
    End If
    
End Sub

''
' \brief  Write an arbitrary number of bytes into the buffer.
'
' \param[in] Values     - The array that contains the bytes
' \param[in] Length     - The number of bytes to write (must be less or equal than \ref Values capacity)
''
Public Sub WriteBytes(ByRef Values() As Byte, Optional ByVal Length As Long = &HFFFFFFFF)
    
    If (Length = &HFFFFFFFF) Then
    
        Length = UBound(Values) + 1

    End If
    
    If (mPosition + Length > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_SPACE)

    Else
        
        Call RtlMoveMemory(mData(mPosition), Values(0), Length)
                
        mPosition = mPosition + Length
        
    End If

End Sub

''
' \brief  Write a boolean into the buffer.
'
' \param[in] Value      - The value to write into the buffer
''
Public Sub WriteBoolean(ByVal Value As Boolean)
    
    If (mPosition + &H1 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_SPACE)

    Else
      
        mData(mPosition) = Value
        
        mPosition = mPosition + &H1
        
    End If
    
End Sub

''
' \brief  Write a 16-bit integer into the buffer.
'
' \param[in] Value      - The value to write into the buffer
''
Public Sub WriteInteger(ByVal Value As Integer)
    
    If (mPosition + &H2 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_SPACE)

    Else
          
        Call PutMem2(mData(mPosition), Value)
        
        mPosition = mPosition + &H2
    
    End If
    
End Sub

''
' \brief  Write a 32-bit integer into the buffer.
'
' \param[in] Value      - The value to write into the buffer
''
Public Sub WriteLong(ByVal Value As Long)
    
    If (mPosition + &H4 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_SPACE)

    Else
          
        Call PutMem4(mData(mPosition), Value)
        
        mPosition = mPosition + &H4
        
    End If
    
End Sub

''
' \brief  Write a 64-bit integer into the buffer.
'
' \param[in] Value      - The value to write into the buffer
''
Public Sub WriteCurrency(ByVal Value As Currency)
    
    If (mPosition + &H8 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_SPACE)

    Else
              
        Call PutMem8(mData(mPosition), ByVal Value)
        
        mPosition = mPosition + &H8
        
    End If
    
End Sub

''
' \brief  Write a 32-bit floating point number into the buffer.
'
' \param[in] Value      - The value to write into the buffer
''
Public Sub WriteSingle(ByVal Value As Single)
    
    If (mPosition + &H4 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_SPACE)

    Else
                  
        Call PutMem4(mData(mPosition), Value)
        
        mPosition = mPosition + &H4
        
    End If
    
End Sub

''
' \brief  Write a 64-bit floating point number into the buffer.
'
' \param[in] Value      - The value to write into the buffer
''
Public Sub WriteDouble(ByVal Value As Double)
    
    If (mPosition + &H8 > mLimit) Then
        
        Call Err.Raise(NOT_ENOUGH_SPACE)

    Else
              
        Call PutMem8(mData(mPosition), ByVal Value)
        
        mPosition = mPosition + &H8
        
    End If
    
End Sub

''
' \brief  Write a string into the buffer.
'
' \param[in]      Value     - The value to write into the buffer
' \param[in, opt] Length    - The number of characters to write, or use the entire array
''
Public Sub WriteString(ByRef Value As String, Optional ByVal Length As Long = &HFFFFFFFF)

    If (Length = &HFFFFFFFF) Then
    
        Length = Len(Value)
                    
        Call WriteInteger(Length)

    End If
    
    If (Length > 0) Then
        
        If (mPosition + Length > mLimit) Then
            
            Call Err.Raise(NOT_ENOUGH_SPACE)
    
        Else
    
            Call RtlMoveMemory(mData(mPosition), ByVal StrPtr(StrConv(Value, vbFromUnicode)), Length)
            
            mPosition = mPosition + Length
            
        End If

    End If
    
End Sub

''
' \brief  Write an unicode string into the buffer.
'
' \param[in]      Value     - The value to write into the buffer
' \param[in, opt] Length    - The number of characters to write, or use the entire array
''
Public Sub WriteUnicode(ByRef Value As String, Optional ByVal Length As Long = &HFFFFFFFF)

    If (Length = &HFFFFFFFF) Then
    
        Length = Len(Value)
                       
        Call WriteInteger(Length)

    End If
            
    If (Length > 0) Then
            
        Length = Length * &H2
           
        If (mPosition + Length > mLimit) Then
            
            Call Err.Raise(NOT_ENOUGH_SPACE)
    
        Else
    
            Call RtlMoveMemory(mData(mPosition), ByVal StrPtr(Value), Length)
    
            mPosition = mPosition + Length
            
        End If
        
    End If
    
End Sub

''
' \brief  Wrap an arbitrary number of bytes into the buffer.
'
' \param[in] Values     - The array that contains the bytes
' \param[in] Length     - The number of bytes to write (must be less or equal than \ref Values capacity)
''
Public Sub Wrap(ByRef Values() As Byte, Optional ByVal Length As Long = &HFFFFFFFF)

    If (Length = &HFFFFFFFF) Then
        
        Length = UBound(Values) + 1
        
    End If

    If (Capacity() < Length) Then
        
        ReDim mData(0 To Length - 1) As Byte
        
    End If
        
    Call Clear
    Call WriteBytes(Values, Length)
    Call Flip
    
End Sub

''
' \brief  Get the capacity (in bytes) of the buffer.
''
Public Property Get Capacity() As Long

    Capacity = UBound(mData) + 1
    
End Property

''
' \brief  Changes the capacity of the buffer.
'
' \param[in] Value      - The new capacity of the buffer
''
Public Property Let Capacity(ByVal Value As Long)
    
    ReDim Preserve mData(0 To Value - 1) As Byte
    
End Property

''
' \brief  Get the limit (in bytes) of the buffer.
''
Public Property Get Limit() As Long

    Limit = mLimit
    
End Property

''
' \brief  Changes the limit of the buffer.
'
' \param[in] Value      - The new limit of the buffer
''
Public Property Let Limit(ByVal Value As Long)

    mLimit = IIf(Value > UBound(mData) + 1, UBound(mData) + 1, Value)
    
End Property

''
' \brief  Get the position (in bytes) of the buffer.
''
Public Property Get Position() As Long

    Position = mPosition
    
End Property

''
' \brief  Changes the position of the buffer.
'
' \param[in] Value      - The new position of the buffer
''
Public Property Let Position(ByVal Value As Long)

    mPosition = IIf(Value > mLimit, mLimit, Value)
    
End Property

''
' \brief  Get the remaining bytes of the buffer.
''
Public Property Get Remaining() As Long

    Remaining = mLimit - mPosition
    
End Property

''
' \brief  Retrieves the NOT_ENOUGH_DATA error code.
'
' \return NOT_ENOUGH_DATA.
''
Public Property Get NotEnoughDataErrCode() As Long

    NotEnoughDataErrCode = NOT_ENOUGH_DATA
    
End Property

''
' \brief  Retrieves the NOT_ENOUGH_SPACE error code.
'
' \return NOT_ENOUGH_SPACE.
''
Public Property Get NotEnoughSpaceErrCode() As Long

    NotEnoughSpaceErrCode = NOT_ENOUGH_SPACE
    
End Property