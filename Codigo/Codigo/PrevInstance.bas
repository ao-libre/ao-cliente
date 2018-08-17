Attribute VB_Name = "PrevInstance"
'**************************************************************
' PrevInstance.bas - Checks for previous instances of the client running
' by using a named mutex.
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
'Prevents multiple instances of the game running on the same computer.
'
' @author Fredy Horacio Treboux (liquid) @and Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070104

Option Explicit

'Declaration of the Win32 API function for creating /destroying a Mutex, and some types and constants.
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByRef lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const ERROR_ALREADY_EXISTS = 183&

Private mutexHID As Long

''
' Creates a Named Mutex. Private function, since we will use it just to check if a previous instance of the app is running.
'
' @param mutexName The name of the mutex, should be universally unique for the mutex to be created.

Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/04/07
'Last Modified by: Juan Martín Sotuyo Dodero (Maraxus) - Changed Security Atributes to make it work in all OS
'***************************************************
    Dim sa As SECURITY_ATTRIBUTES
    
    With sa
        .bInheritHandle = 0
        .lpSecurityDescriptor = 0
        .nLength = LenB(sa)
    End With
    
    mutexHID = CreateMutex(sa, False, "Global\" & mutexName)
    
    CreateNamedMutex = Not (Err.LastDllError = ERROR_ALREADY_EXISTS) 'check if the mutex already existed
End Function

''
' Checks if there's another instance of the app running, returns True if there is or False otherwise.

Public Function FindPreviousInstance() As Boolean
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/04/07
'
'***************************************************
    'We try to create a mutex, the name could be anything, but must contain no backslashes.
    If CreateNamedMutex("UniqueNameThatActuallyCouldBeAnything") Then
        'There's no other instance running
        FindPreviousInstance = False
    Else
        'There's another instance running
        FindPreviousInstance = True
    End If
End Function

''
' Closes the client, allowing other instances to be open.

Public Sub ReleaseInstance()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/04/07
'
'***************************************************
    Call ReleaseMutex(mutexHID)
    Call CloseHandle(mutexHID)
End Sub
