VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================
' Name:     vbAccelerator cStringBuilder
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     1 January 2002
'
' Copyright © 2002 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' VB can be slow to append strings together because of the continual
' reallocation of string size.  This class pre-allocates a string in
' blocks and hence removes the performance restriction.
'
' Quicker insert and remove is also possible since string space does
' not have to be reallocated.
'
' Example:
' Adding "http://vbaccelerator.com/" 10,000 times to a string:
' Standard VB:   34s
' This Class:    0.35s
'
' ======================================================================================

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (pDst As Any, _
                                       pSrc As Any, _
                                       ByVal ByteLen As Long)
      
Private m_sString    As String
Private m_iChunkSize As Long
Private m_iPos       As Long
Private m_iLen       As Long

Public Property Get ChunkSize() As Long
    ' Return the unicode character chunk size:
    
    On Error GoTo ChunkSize_Err
    
    ChunkSize = m_iChunkSize \ 2

    
    Exit Property

ChunkSize_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "ChunkSize"
    End If
Resume Next
    
End Property


Public Property Get toString() As String
    
    On Error GoTo toString_Err
    

    ' The internal string:
    If m_iPos > 0 Then
        toString = Left$(m_sString, m_iPos \ 2)

    End If

    
    Exit Property

toString_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "toString"
    End If
Resume Next
    
End Property


Public Sub Append(ByRef sThis As String)
    
    On Error GoTo Append_Err
    
    Dim lLen        As Long
    Dim lLenPlusPos As Long
 
    ' Append an item to the string:
    lLen = LenB(sThis)
    lLenPlusPos = lLen + m_iPos

    If lLenPlusPos > m_iLen Then
        Dim lTemp As Long
      
        lTemp = m_iLen

        Do While lTemp < lLenPlusPos
            lTemp = lTemp + m_iChunkSize
        Loop
      
        m_sString = m_sString & Space$((lTemp - m_iLen) \ 2)
        m_iLen = lTemp

    End If
   
    CopyMemory ByVal UnsignedAdd(StrPtr(m_sString), m_iPos), ByVal StrPtr(sThis), lLen
    m_iPos = m_iPos + lLen

    
    Exit Sub

Append_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "Append"
    End If
Resume Next
    
End Sub


Public Sub Insert(ByVal iIndex As Long, ByRef sThis As String)
    
    On Error GoTo Insert_Err
    
    Dim lLen  As Long
    Dim lPos  As Long
    Dim lSize As Long
   
    ' is iIndex within bounds?
    If (iIndex * 2 > m_iPos) Then
        Err.Raise 9
    Else
   
        lLen = LenB(sThis)

        If (m_iPos + lLen) > m_iLen Then
            m_sString = m_sString & Space$(m_iChunkSize \ 2)
            m_iLen = m_iLen + m_iChunkSize

        End If
      
        ' Move existing characters from current position
        lPos = UnsignedAdd(StrPtr(m_sString), iIndex * 2)
        lSize = m_iPos - iIndex * 2
      
        ' moving from iIndex to iIndex + lLen
        CopyMemory ByVal UnsignedAdd(lPos, lLen), ByVal lPos, lSize
      
        ' Insert new characters:
        CopyMemory ByVal lPos, ByVal StrPtr(sThis), lLen
      
        m_iPos = m_iPos + lLen

    End If

    
    Exit Sub

Insert_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "Insert"
    End If
Resume Next
    
End Sub

Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
    ' This function is useful when doing pointer arithmetic,
    ' but note it only works for positive values of Incr
    
    On Error GoTo UnsignedAdd_Err
    

    If Start And &H80000000 Then 'Start < 0
        UnsignedAdd = Start + Incr
    ElseIf (Start Or &H80000000) < -Incr Then
        UnsignedAdd = Start + Incr
    Else
        UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)

    End If
   
    
    Exit Function

UnsignedAdd_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "UnsignedAdd"
    End If
Resume Next
    
End Function

Private Sub Class_Initialize()
    ' The default allocation: 8192 characters.
    
    On Error GoTo Class_Initialize_Err
    
    m_iChunkSize = 16384

    
    Exit Sub

Class_Initialize_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "Class_Initialize"
    End If
Resume Next
    
End Sub

