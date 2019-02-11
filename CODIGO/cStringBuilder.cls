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

Public Property Get length() As Long
    
    On Error GoTo length_Err
    
    length = m_iPos \ 2

    
    Exit Property

length_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "length"
    End If
Resume Next
    
End Property

Public Property Get Capacity() As Long
    
    On Error GoTo Capacity_Err
    
    Capacity = m_iLen \ 2

    
    Exit Property

Capacity_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "Capacity"
    End If
Resume Next
    
End Property

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

Public Property Let ChunkSize(ByVal iChunkSize As Long)
    ' Set the chunksize.  We multiply by 2 because internally
    ' we are considering bytes:
    
    On Error GoTo ChunkSize_Err
    
    m_iChunkSize = iChunkSize * 2

    
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

Public Property Let TheString(ByRef sThis As String)
    
    On Error GoTo TheString_Err
    
    Dim lLen As Long
   
    ' Setting the string:
    lLen = LenB(sThis)

    If lLen = 0 Then
        'Clear
        m_sString = vbNullString
        m_iPos = 0
        m_iLen = 0
    Else

        If m_iLen < lLen Then
            ' Need to expand string to accommodate:
            Do
                m_sString = m_sString & Space$(m_iChunkSize \ 2)
                m_iLen = m_iLen + m_iChunkSize
            Loop While m_iLen < lLen

        End If

        CopyMemory ByVal StrPtr(m_sString), ByVal StrPtr(sThis), lLen
        m_iPos = lLen

    End If
   
    
    Exit Property

TheString_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "TheString"
    End If
Resume Next
    
End Property

Public Sub Clear()
    
    On Error GoTo Clear_Err
    
    m_sString = vbNullString
    m_iPos = 0
    m_iLen = 0

    
    Exit Sub

Clear_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "Clear"
    End If
Resume Next
    
End Sub

Public Sub AppendNL(ByRef sThis As String)
    
    On Error GoTo AppendNL_Err
    
    Append sThis
    Append vbNewLine

    
    Exit Sub

AppendNL_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "AppendNL"
    End If
Resume Next
    
End Sub

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

Public Sub AppendByVal(ByVal sThis As String)
    
    On Error GoTo AppendByVal_Err
    
    Append sThis

    
    Exit Sub

AppendByVal_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "AppendByVal"
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

Public Sub InsertByVal(ByVal iIndex As Long, ByVal sThis As String)
    
    On Error GoTo InsertByVal_Err
    
    Insert iIndex, sThis

    
    Exit Sub

InsertByVal_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "InsertByVal"
    End If
Resume Next
    
End Sub

Public Sub Remove(ByVal iIndex As Long, ByVal lLen As Long)
    
    On Error GoTo Remove_Err
    
    Dim lSrc  As Long
    Dim lDst  As Long
    Dim lSize As Long

    ' is iIndex within bounds?
    If (iIndex * 2 > m_iPos) Then
        Err.Raise 9
    Else

        ' is there sufficient length?
        If ((iIndex + lLen) * 2 > m_iPos) Then
            Err.Raise 9
        Else
            ' Need to copy characters from iIndex*2 to m_iPos back by lLen chars:
            lSrc = UnsignedAdd(StrPtr(m_sString), (iIndex + lLen) * 2)
            lDst = UnsignedAdd(StrPtr(m_sString), iIndex * 2)
            lSize = (m_iPos - (iIndex + lLen) * 2)
            CopyMemory ByVal lDst, ByVal lSrc, lSize
            m_iPos = m_iPos - lLen * 2

        End If

    End If

    
    Exit Sub

Remove_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "Remove"
    End If
Resume Next
    
End Sub

Public Function Find(ByVal sToFind As String, _
                     Optional ByVal lStartIndex As Long = 1, _
                     Optional ByVal compare As VbCompareMethod = vbTextCompare) As Long
    
    On Error GoTo Find_Err
    
   
    Dim lInstr As Long

    If (lStartIndex > 0) Then
        lInstr = InStr(lStartIndex, m_sString, sToFind, compare)
    Else
        lInstr = InStr(m_sString, sToFind, compare)

    End If

    If (lInstr < m_iPos \ 2) Then
        Find = lInstr

    End If

    
    Exit Function

Find_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "Find"
    End If
Resume Next
    
End Function

Public Sub HeapMinimize()
    
    On Error GoTo HeapMinimize_Err
    
    Dim iLen As Long
   
    ' Reduce the string size so only the minimal chunks
    ' are allocated:
    If (m_iLen - m_iPos) > m_iChunkSize Then
        iLen = m_iLen

        Do While (iLen - m_iPos) > m_iChunkSize
            iLen = iLen - m_iChunkSize
        Loop
        m_sString = Left$(m_sString, iLen \ 2)
        m_iLen = iLen

    End If
   
    
    Exit Sub

HeapMinimize_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "cStringBuilder" & "->" & "HeapMinimize"
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

