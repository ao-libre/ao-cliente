Attribute VB_Name = "mVerProcesos"
Option Explicit
 
Public Const TH32CS_SNAPPROCESS As Long = &H2
Public Const MAX_PATH As Integer = 260
 
Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
 
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias _
"CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
 
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" _
(ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
 
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" _
(ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
 
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

'Esta función Api devuelve un valor  Boolean indicando si la ventana es una ventana visible
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

'Esta función retorna el número de caracteres del caption de la ventana
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

'Esta devuelve el texto. Se le pasa el hwnd de la ventana, un buffer donde se
'almacenará el texto devuelto, y el Lenght de la cadena en el último parámetro
'que obtuvimos con el Api GetWindowTextLength
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'Esta es la función Api que busca las ventanas y retorna su handle o Hwnd
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wFlag As Long) As Long

'Constantes para buscar las ventanas mediante el Api GetWindow
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&
Private CANTv As Byte
 
Public Function ListarProcesosUsuario() As String
On Error Resume Next
     
    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim r As Long
    ListarProcesosUsuario = ""
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot = 0 Then
        ListarProcesosUsuario = "ERROR"
        Exit Function
    End If
    
    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapShot, uProcess)
    Dim DatoP As String
    
    While r <> 0
        If InStr(uProcess.szExeFile, ".exe") <> 0 Then
            DatoP = ReadField(1, uProcess.szExeFile, Asc("."))
            ListarProcesosUsuario = ListarProcesosUsuario & "|" & DatoP
     
        End If

        r = ProcessNext(hSnapShot, uProcess)
    Wend
    Call CloseHandle(hSnapShot)
End Function

Public Function ListarCaptionsUsuario() As String
On Error Resume Next
    Dim buf As Long, handle As Long, titulo As String, lenT As Long, ret As Long

    'Obtenemos el Hwnd de la primera ventana, usando la constante GW_HWNDFIRST
    handle = GetWindow(Screen.ActiveForm.hWnd, GW_HWNDFIRST)

    'Este bucle va a recorrer todas las ventanas.
    'cuando GetWindow devielva un 0, es por que no hay mas
    Do While handle <> 0
    
        'Tenemos que comprobar que la ventana es una de tipo visible
        If IsWindowVisible(handle) Then

            'Obtenemos el número de caracteres de la ventana
            lenT = GetWindowTextLength(handle)

            'si es el número anterior es mayor a 0
            If lenT > 0 Then
                'Creamos un buffer. Este buffer tendrá el tamaño con la variable LenT
                titulo = String$(lenT, 0)
                
                'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
                'y tambien debemos pasarle el Hwnd de dicha ventana
                ret = GetWindowText(handle, titulo, lenT + 1)
                titulo$ = Left$(titulo, ret)

                'La agregamos string
                ListarCaptionsUsuario = titulo & "#" & ListarCaptionsUsuario
                CANTv = CANTv + 1
            End If
        End If

        'Buscamos con GetWindow la próxima ventana usando la constante GW_HWNDNEXT
        handle = GetWindow(handle, GW_HWNDNEXT)
       Loop
End Function

