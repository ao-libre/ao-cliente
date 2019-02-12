VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsIniReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Type ChildNode

    key As String
    value As String

End Type

Private Type MainNode

    Name As String
    values() As ChildNode
    numValues As Integer

End Type

''
'Containts all Main sections of the loaded INI file
'
Private fileData() As MainNode

''
'Stores the total number of main sections in the loaded INI file
'
Private MainNodes  As Long

''
'Default constructor. Does nothing.
'

Private Sub Class_Initialize()

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    '
    '**************************************************************
End Sub

''
'Destroy every array and deallocates al memory.
'

Private Sub Class_Terminate()
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    '
    '**************************************************************
    
    On Error GoTo Class_Terminate_Err
    
    Dim i As Long
    Dim Nodes As Long
    
    'Clean up
    If MainNodes Then
        
        Nodes = MainNodes - 1
        
        For i = 1 To Nodes
            Erase fileData(i).values
        Next i
        
        Erase fileData

    End If

    
    Exit Sub

Class_Terminate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "ClsIniReader" & "->" & "Class_Terminate"
    End If
Resume Next
    
End Sub

''
'Loads a INI file so it's values can be read. Must be called before being able to use GetValue.
'
'   @param    file Complete path of the INI file to be loaded.
'   @see      GetValue
'

Public Sub Initialize(ByVal File As String)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    'Opens the requested file and loads it's data into memory
    '**************************************************************
    
    On Error GoTo Initialize_Err
    
    Dim handle As Integer
    Dim Text   As String
    Dim Pos    As Long
    
    'Get a free handle and start reading line by line until the end
    handle = FreeFile
    
    Open File For Input As handle
    
    Do Until EOF(handle)
        Line Input #handle, Text
        
        'Is it null??
        If Len(Text) Then

            'If it starts with '['   it is a main node or nothing (GetPrivateProfileStringA works this way), otherwise it's a value
            If Left$(Text, 1) = "[" Then
                'If it has an ending ']'   it's a main node, otherwise it's nothing
                Pos = InStr(2, Text, "]")

                If Pos Then
                    'Add a main node
                    ReDim Preserve fileData(MainNodes) As MainNode
                    
                    fileData(MainNodes).Name = UCase$(Trim$(mid$(Text, 2, Pos - 2)))
                    
                    MainNodes = MainNodes + 1

                End If

            Else
                'So it's a value. Check if it has a '=', otherwise it's nothing
                Pos = InStr(2, Text, "=")

                If Pos Then

                    'Is it under any main node??
                    If MainNodes Then

                        With fileData(MainNodes - 1)
                            'Add it to the main node's value
                            ReDim Preserve .values(.numValues) As ChildNode
                            
                            .values(.numValues).value = Right$(Text, Len(Text) - Pos)
                            .values(.numValues).key = UCase$(Left$(Text, Pos - 1))
                            
                            .numValues = .numValues + 1

                        End With

                    End If

                End If

            End If

        End If

    Loop
    
    Close handle
    
    Dim i As Long
    Dim Nodes As Long
    
    If MainNodes Then
        Nodes = MainNodes - 1
        
        'Sort main nodes to allow binary search
        Call SortMainNodes(0, Nodes)
        
        'Sort values of each node to allow binary search
        For i = 0 To Nodes

            If fileData(i).numValues Then Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
        Next i

    End If

    
    Exit Sub

Initialize_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "ClsIniReader" & "->" & "Initialize"
    End If
Resume Next
    
End Sub

''
'Sorts all child nodes within the given MainNode alphabetically by their keys. Uses quicksort.
'
'   @param    Node The MainNode whose values are to be sorted.
'   @param    first The first index to consider when sorting.
'   @param    last The last index to be considered when sorting.
'

Private Sub SortChildNodes(ByRef Node As MainNode, _
                           ByVal first As Integer, _
                           ByVal last As Integer)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    'Sorts the list of values in a given MainNode using quicksort,
    'this allows the use of Binary Search for faster searches
    '**************************************************************
    
    On Error GoTo SortChildNodes_Err
    
    Dim Min  As Integer      'First item in the list
    Dim Max  As Integer      'Last item in the list
    Dim comp As String      'Item used to compare
    Dim Temp As ChildNode
    
    Min = first
    Max = last
    
    With Node
        comp = .values((Min + Max) \ 2).key
    
        Do While Min <= Max
            Do While .values(Min).key < comp And Min < last
                Min = Min + 1
            Loop

            Do While .values(Max).key > comp And Max > first
                Max = Max - 1
            Loop

            If Min <= Max Then
                Temp = .values(Min)
                .values(Min) = .values(Max)
                .values(Max) = Temp
                Min = Min + 1
                Max = Max - 1

            End If

        Loop

    End With
    
    If first < Max Then SortChildNodes Node, first, Max
    If Min < last Then SortChildNodes Node, Min, last

    
    Exit Sub

SortChildNodes_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "ClsIniReader" & "->" & "SortChildNodes"
    End If
Resume Next
    
End Sub

''
'Sorts all main nodes in the loaded INI file alphabetically by their names. Uses quicksort.
'
'   @param    first The first index to consider when sorting.
'   @param    last The last index to be considered when sorting.
'

Private Sub SortMainNodes(ByVal first As Integer, ByVal last As Integer)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    'Sorts the MainNodes list using quicksort,
    'this allows the use of Binary Search for faster searches
    '**************************************************************
    
    On Error GoTo SortMainNodes_Err
    
    Dim Min  As Integer      'First item in the list
    Dim Max  As Integer      'Last item in the list
    Dim comp As String      'Item used to compare
    Dim Temp As MainNode
    
    Min = first
    Max = last
    
    comp = fileData((Min + Max) \ 2).Name
    
    Do While Min <= Max
        Do While fileData(Min).Name < comp And Min < last
            Min = Min + 1
        Loop

        Do While fileData(Max).Name > comp And Max > first
            Max = Max - 1
        Loop

        If Min <= Max Then
            Temp = fileData(Min)
            fileData(Min) = fileData(Max)
            fileData(Max) = Temp
            Min = Min + 1
            Max = Max - 1

        End If

    Loop
    
    If first < Max Then SortMainNodes first, Max
    If Min < last Then SortMainNodes Min, last

    
    Exit Sub

SortMainNodes_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "ClsIniReader" & "->" & "SortMainNodes"
    End If
Resume Next
    
End Sub

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
'   @param    Main The name of the main section in which we will be searching.
'   @param    key The key of the value we are looking for.
'   @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.
'

Public Function GetValue(ByVal Main As String, ByVal key As String) As String
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    'Returns a value if the key and main node exist, or a nullstring otherwise
    '**************************************************************
    
    On Error GoTo GetValue_Err
    
    Dim i As Long
    Dim j As Long
    
    'Search for the main node
    i = FindMain(UCase$(Main))
    
    If i >= 0 Then
        'If valid, binary search among keys
        j = FindKey(fileData(i), UCase$(key))
        
        'If we found it we return it
        If j >= 0 Then GetValue = fileData(i).values(j).value

    End If

    
    Exit Function

GetValue_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "ClsIniReader" & "->" & "GetValue"
    End If
Resume Next
    
End Function

''
'Searches for a given key within a given main node and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
'   @param    Node The MainNode among whose value entries we will be searching.
'   @param    key The key of the value we are looking for.
'   @returns  The index in which the value with the key we are looking for is stored or the negation of the index in which it should be if not found.
'

Private Function FindKey(ByRef Node As MainNode, ByVal key As String) As Long
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    'Returns the index of the value which key matches the requested one,
    'or the negation of the position were it should be if not found
    '**************************************************************
    
    On Error GoTo FindKey_Err
    
    Dim Min As Long
    Dim Max As Long
    Dim mid As Long
    
    Min = 0
    Max = Node.numValues - 1
    
    Do While Min <= Max
        mid = (Min + Max) \ 2
        
        If Node.values(mid).key < key Then
            Min = mid + 1
        ElseIf Node.values(mid).key > key Then
            Max = mid - 1
        Else
            'We found it
            FindKey = mid
            Exit Function

        End If

    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindKey = Not mid

    
    Exit Function

FindKey_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "ClsIniReader" & "->" & "FindKey"
    End If
Resume Next
    
End Function

''
'Searches for a main section with the given name within the loaded INI file and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
'   @param    name The name of the MainNode we are looking for.
'   @returns  The index in which the main section we are looking for is stored or the negation of the index in which it should be if not found.
'

Private Function FindMain(ByVal Name As String) As Long
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    'Returns the index of the MainNode which name matches the requested one,
    'or the negation of the position were it should be if not found
    '**************************************************************
    
    On Error GoTo FindMain_Err
    
    Dim Min As Long
    Dim Max As Long
    Dim mid As Long
    
    Min = 0
    Max = MainNodes - 1
    
    Do While Min <= Max
        mid = (Min + Max) \ 2
        
        If fileData(mid).Name < Name Then
            Min = mid + 1
        ElseIf fileData(mid).Name > Name Then
            Max = mid - 1
        Else
            'We found it
            FindMain = mid
            Exit Function

        End If

    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindMain = Not mid

    
    Exit Function

FindMain_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "ClsIniReader" & "->" & "FindMain"
    End If
Resume Next
    
End Function

