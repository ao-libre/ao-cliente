Attribute VB_Name = "Application"
'**************************************************************
' Application.bas - General API methods regarding the Application in general.
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


Option Explicit

''
' Retrieves the active window's hWnd for this app.
'
' @return Retrieves the active window's hWnd for this app. If this app is not in the foreground it returns 0.

Private Declare Function GetActiveWindow Lib "user32" () As Long

'Declaration of the Win32 API function for creating/destroying a Mutex, and some types and constants.
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

Private sNotepadTaskId As String

''
' Checks if this is the active (foreground) application or not.
'
' @return   True if any of the app's windows are the foreground window, false otherwise.

Public Function IsAppActive() As Boolean
'***************************************************
'Author: Juan Martin Sotuyo Dodero (maraxus)
'Last Modify Date: 03/03/2007
'Checks if this is the active application or not
'***************************************************
    IsAppActive = (GetActiveWindow <> 0)
End Function


''
'Prevents multiple instances of the game running on the same computer.
'
' @author Fredy Horacio Treboux (liquid) @and Juan Martin Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070104


''
' Creates a Named Mutex. Private function, since we will use it just to check if a previous instance of the app is running.
'
' @param mutexName The name of the mutex, should be universally unique for the mutex to be created.

Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/04/07
'Last Modified by: Juan Martin Sotuyo Dodero (Maraxus) - Changed Security Atributes to make it work in all OS
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
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 01/04/07
'
'***************************************************
    Call ReleaseMutex(mutexHID)
    Call CloseHandle(mutexHID)
End Sub

Public Sub LogError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
'**********************************************************
'Author: Jopi
'Guarda una descripcion detallada del error en Errores.log
'**********************************************************
    Dim file As Integer
    file = FreeFile

    'Hacemos un Left para poder solo obtener la letra del HD
    'Por que por culpa del UAC no guarda los logs en la carpeta del juego...
    Dim ErroresPath As String
    ErroresPath = Left$(App.path, 2) & "\AO-Libre\Errores\"

    If Dir(ErroresPath, vbDirectory) = "" Then
        MkDir ErroresPath
    End If

    'Matamos Notepad para evitar abrir decenas de block de notas.
    Shell ("taskkill /PID " & sNotepadTaskId)
        
    Open ErroresPath & "\Errores.log" For Append As #file
    
        Print #file, "Error: " & Numero
        Print #file, "Descripcion: " & Descripcion
        
        If LenB(Linea) <> 0 Then
            Print #file, "Linea: " & Linea
        End If
        
        Print #file, "Componente: " & Componente
        Print #file, "Fecha y Hora: " & Date$ & "-" & Time$
        Print #file, vbNullString
        
    Close #file
    
    Debug.Print "Error: " & Numero & vbNewLine & _
                "Descripcion: " & Descripcion & vbNewLine & _
                "Componente: " & Componente & vbNewLine & _
                "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine

    sNotepadTaskId = Shell("Notepad " & ErroresPath & "\Errores.log")

    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_ERRORES_LOG_CARPETA").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_ERRORES_LOG_CARPETA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_ERRORES_LOG_CARPETA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_ERRORES_LOG_CARPETA").item("COLOR").item(3), _
                            False, False, True)

End Sub
