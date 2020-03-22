VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsIniManager.cls - Loads INI files into memory and applies Binary Search to get values at high speed.
' Use it instead of GetVar when reading several values form the same file at once, otherwise it's not usefull.
' Based in the idea of AlejoLP and his clsLeerInis.
'
' Designed and implemented by Juan Martin Sotuyo Dodero (Maraxus)
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
'Loads a complete INI file into memory and sorts it's data and keys for faster searches.
'It is MUCH faster than GetPrivateProfileStringA if you search for several values within a file, otherwise stick to the API.
'It's particularly usefull when you can keep a whole file in memory such as NPCs.dat'
' Based in the idea of AlejoLP and his clsLeerInis.
'
' @author Juan Martin Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.1.0
' @date 20060501

'01/05/2006 - Juan Martin Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - First Release
'
'01/04/2008 - Juan Martin Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Add: KeyExists method allows to check for valid section keys.
'
'10/10/2010 - Torres Patricio(Pato)
'   - Add:  AddNode method to add a main node.
'           AddKey method to add key into the refered main node.
'           DumpFile method to dump the data into a file.
'
'21/03/2020 - Joaquin Ortiz de Zarate (jopiortiz)
'   - Add:  Respect upper/lower-casing while saving Keys and Values.

Option Explicit
Option Base 0

''
'Structure that contains a value and it's key in a INI file
'
' @param    key String containing the key associated to the value.
' @param    value String containing the value of the INI entry.
' @see      MainNode
'
Private Type ChildNode
    key As String
    Value As String
End Type

''
'Structure that contains all info under a tag in a INI file.
'Such tags are indicated with the "[" and "]" characters.
'
' @param    name String containing the text within the "[" and "]" characters.
'It's the key used when searching for a main section of the INI data.
' @param    values Array of ChildNodes, each containing a value entry along with it's key.
' @param    numValues Number of entrys in the main node.

Private Type MainNode
    name As String
    values() As ChildNode
    numValues As Integer
End Type

''
'Containts all Main sections of the loaded INI file
Private fileData() As MainNode

''
'Stores the total number of main sections in the loaded INI file
Private MainNodes  As Long

''
'Default constructor. Does nothing.

Private Sub Class_Initialize()

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    '
    '**************************************************************
End Sub

''
'Destroy every array and deallocates al memory.
'

Private Sub Class_Terminate()

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    '
    '**************************************************************
    Dim i As Long
    
    'Clean up
    If MainNodes Then

        For i = 1 To MainNodes - 1
            Erase fileData(i).values
        Next i
        
        Erase fileData

    End If
    
    MainNodes = 0

End Sub

''
'Loads a INI file so it's values can be read. Must be called before being able to use GetValue.
'
' @param    file Complete path of the INI file to be loaded.
' @see      GetValue

Public Sub Initialize(ByVal File As String)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero
    'Last Modify Date: 27/07/2006
    'Opens the requested file and loads it's data into memory
    '**************************************************************
    Dim handle As Integer
    Dim Text   As String
    Dim Pos    As Long
    
    'Prevent memory losses if we are attempting to reload a file....
    Call Class_Terminate
    
    'Get a free handle and start reading line by line until the end
    handle = FreeFile
    
    Open File For Input As handle
    
    Do Until EOF(handle)
        Line Input #handle, Text
        
        'Is it null??
        If Len(Text) Then

            'If it starts with '[' it is a main node or nothing (GetPrivateProfileStringA works this way), otherwise it's a value
            If Left$(Text, 1) = "[" Then
                
                'If it has an ending ']' it's a main node, otherwise it's nothing
                Pos = InStr(2, Text, "]")

                If Pos Then
                
                    'Add a main node
                    ReDim Preserve fileData(MainNodes) As MainNode
                    
                    fileData(MainNodes).name = Trim$(mid$(Text, 2, Pos - 2))
                    
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
                            
                            .values(.numValues).Value = Right$(Text, Len(Text) - Pos)
                            .values(.numValues).key = Left$(Text, Pos - 1)
                            
                            .numValues = .numValues + 1

                        End With

                    End If

                End If

            End If

        End If

    Loop
    
    Close handle
    
    Dim i As Long
    
    If MainNodes Then
        'Sort main nodes to allow binary search
        Call SortMainNodes(0, MainNodes - 1)
        
        'Sort values of each node to allow binary search
        For i = 0 To MainNodes - 1
            If fileData(i).numValues Then
                Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
            End If
        Next i

    End If

End Sub

''
'Sorts all child nodes within the given MainNode alphabetically by their keys. Uses quicksort.
'
' @param    Node The MainNode whose values are to be sorted.
' @param    first The first index to consider when sorting.
' @param    last The last index to be considered when sorting.

Private Sub SortChildNodes(ByRef Node As MainNode, _
                           ByVal First As Integer, _
                           ByVal Last As Integer)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    'Sorts the list of values in a given MainNode using quicksort,
    'this allows the use of Binary Search for faster searches
    '**************************************************************
    Dim min  As Integer      'First item in the list
    Dim max  As Integer      'Last item in the list
    Dim comp As String      'Item used to compare
    Dim temp As ChildNode
    
    min = First
    max = Last
    
    With Node
        comp = UCase$(.values((min + max) \ 2).key)
    
        Do While min <= max
            Do While UCase$(.values(min).key) < comp And min < Last
                min = min + 1
            Loop

            Do While UCase$(.values(max).key) > comp And max > First
                max = max - 1
            Loop

            If min <= max Then
                temp = .values(min)
                .values(min) = .values(max)
                .values(max) = temp
                min = min + 1
                max = max - 1

            End If

        Loop

    End With
    
    If First < max Then SortChildNodes Node, First, max
    If min < Last Then SortChildNodes Node, min, Last

End Sub

''
'Sorts all main nodes in the loaded INI file alphabetically by their names. Uses quicksort.
'
' @param    first The first index to consider when sorting.
' @param    last The last index to be considered when sorting.

Private Sub SortMainNodes(ByVal First As Integer, ByVal Last As Integer)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    'Sorts the MainNodes list using quicksort,
    'this allows the use of Binary Search for faster searches
    '**************************************************************
    Dim min  As Integer      'First item in the list
    Dim max  As Integer      'Last item in the list
    Dim comp As String      'Item used to compare
    Dim temp As MainNode
    
    min = First
    max = Last
    
    comp = UCase$(fileData((min + max) \ 2).name)
    
    Do While min <= max
        Do While UCase$(fileData(min).name) < comp And min < Last
            min = min + 1
        Loop

        Do While UCase$(fileData(max).name) > comp And max > First
            max = max - 1
        Loop

        If min <= max Then
            temp = fileData(min)
            fileData(min) = fileData(max)
            fileData(max) = temp
            min = min + 1
            max = max - 1

        End If

    Loop
    
    If First < max Then SortMainNodes First, max
    If min < Last Then SortMainNodes min, Last

End Sub

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Function GetValue(ByVal Main As String, ByVal key As String) As String

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    'Returns a value if the key and main node exist, or a nullstring otherwise
    '**************************************************************
    Dim i As Long
    Dim J As Long
    
    'Search for the main node
    i = FindMain(Main)
    
    If i >= 0 Then
    
        'If valid, binary search among keys
        J = FindKey(i, key)
        
        'If we found it we return it
        If J >= 0 Then GetValue = fileData(i).values(J).Value

    End If

End Function

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.

Public Sub ChangeValue(ByVal Main As String, ByVal key As String, ByVal Value As String)

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 27/05/2009
    'If the key and main node exist, changes the value
    '**************************************************************
    Dim i As Long
    Dim J As Long
    
    'Search for the main node
    i = FindMain(Main)
    
    If i < 0 Then i = AddNode(Main)
    
    'If valid, binary search among keys
    J = FindKey(i, key)
    
    If J < 0 Then
        Call AddKey(i, key, Value)
    
    Else
        
        'If we found it we change it
        fileData(i).values(J).Value = Value

    End If

End Sub

''
'Searches for a given key within a given main node and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
' @param    Node The MainNode among whose value entries we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The index in which the value with the key we are looking for is stored or the negation of the index in which it should be if not found.

Private Function FindKey(ByVal MainPos As Long, ByVal key As String) As Long

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    'Returns the index of the value which key matches the requested one,
    'or the negation of the position were it should be if not found
    '**************************************************************
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    key = UCase$(key)
    
    With fileData(MainPos)
    
        min = 0
        max = .numValues - 1
        
        Do While min <= max
            mid = (min + max) \ 2
            
            If UCase$(.values(mid).key) < key Then
                min = mid + 1
            
            ElseIf UCase$(.values(mid).key) > key Then
                max = mid - 1
            
            Else
                
                'We found it
                FindKey = mid
                
                Exit Function
    
            End If
    
        Loop
        
    End With
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindKey = Not min

End Function

''
'Searches for a main section with the given name within the loaded INI file and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
' @param    name The name of the MainNode we are looking for.
' @returns  The index in which the main section we are looking for is stored or the negation of the index in which it should be if not found.

Private Function FindMain(ByVal name As String) As Long

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero
    'Last Modify Date: 5/01/2006
    'Returns the index of the MainNode which name matches the requested one,
    'or the negation of the position were it should be if not found
    '**************************************************************
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    name = UCase$(name)
    
    min = 0
    max = MainNodes - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If UCase$(fileData(mid).name) < name Then
            min = mid + 1
        
        ElseIf UCase$(fileData(mid).name) > name Then
            max = mid - 1
        
        Else
            'We found it
            FindMain = mid
            Exit Function

        End If

    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindMain = Not min

End Function

''
'Checks wether a given key exists or not.
'
' @param    name    The name of the element whose existance is being checked.
' @returns  True if the key exists, false otherwise.

Public Function KeyExists(ByVal name As String) As Boolean
    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero
    'Last Modify Date: 04/01/2008
    'Returns true of the key exists, false otherwise.
    '**************************************************************
    KeyExists = FindMain(name) >= 0

End Function

Private Function AddNode(ByRef name As String) As Long

    '**************************************************************
    'Author: Torres Patricio (Pato)
    'Last Modify Date: 10/10/2010
    '
    '**************************************************************
    Dim i       As Long
    Dim MainPos As Long

    'Add a main node
    ReDim Preserve fileData(MainNodes) As MainNode

    If MainNodes Then
        
        i = MainNodes - 1
        
        MainPos = Not FindMain(name)
    
        Do While i >= MainPos
            fileData(i + 1) = fileData(i)
            i = i - 1
        Loop
    
        fileData(MainPos).numValues = 0
        
        Erase fileData(MainPos).values

    End If

    fileData(MainPos).name = name

    MainNodes = MainNodes + 1

    AddNode = MainPos

End Function

Private Function AddKey(ByVal MainPos As Long, _
                        ByRef key As String, _
                        Optional ByRef Value As String = vbNullString) As Long

    '**************************************************************
    'Author: Torres Patricio (Pato)
    'Last Modify Date: 10/10/2010
    '
    '**************************************************************
    Dim i       As Long
    Dim KeyPos  As Long

    With fileData(MainPos)
        
        'Add it to the main node's value
        ReDim Preserve .values(.numValues) As ChildNode
    
        If .numValues > 0 Then
            i = .numValues - 1
            KeyPos = Not FindKey(MainPos, key)
        
            Do While i >= KeyPos
                .values(i + 1) = .values(i)
                i = i - 1
            Loop

        End If
    
        .values(KeyPos).key = key
        .values(KeyPos).Value = Value
    
        .numValues = .numValues + 1

    End With

    AddKey = KeyPos

End Function

Public Sub DumpFile(ByRef File As String)

    '**************************************************************
    'Author: Torres Patricio (Pato)
    'Last Modify Date: 10/10/2010
    '
    '**************************************************************
    Dim hFile As Integer
    Dim i     As Long
    Dim J     As Long

    hFile = FreeFile()

    Open File For Output As hFile

    For i = 0 To MainNodes - 1

        With fileData(i)
            Print #hFile, "[" & .name & "]"
            
            For J = 0 To .numValues - 1
                Print #hFile, .values(J).key & "=" & .values(J).Value
            Next J
            
            Print #hFile, vbNullString

        End With

    Next i
    
    Close hFile

End Sub
