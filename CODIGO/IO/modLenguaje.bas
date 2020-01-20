Attribute VB_Name = "ModLenguaje"
Option Explicit

'Constantes para el Api GetLocaleInfo
'************************************
Const LOCALE_USER_DEFAULT = &H400
'Const LOCALE_SENGLANGUAGE = &H1001
  
'Declaracion de la funcion Api GetLocaleInfo
Private Declare Function GetLocaleInfo _
                Lib "kernel32" _
                Alias "GetLocaleInfoA" (ByVal Locale As Long, _
                                        ByVal LCType As Long, _
                                        ByVal lpLCData As String, _
                                        ByVal cchData As Long) As Long

Public JsonLanguage As Object
Public Language As String

Public Function FileToString(strFileName As String) As String
    '###################################################################################
    ' Convierte un archivo entero a una cadena de texto para almacenarla en una variable
    '###################################################################################
    Dim IFile As Variant
    
    IFile = FreeFile
    Open strFileName For Input As #IFile
        FileToString = StrConv(InputB(LOF(IFile), IFile), vbUnicode)
    Close #IFile
End Function

Public Function ObtainOperativeSystemLanguage(ByVal lInfo As Long) As String
    '*******************************************
    ' Funcion que obtiene el idioma del sistema
    '*******************************************

    Dim Buffer As String, ret As String

    Buffer = String$(256, 0)
            
    ret = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, Buffer, Len(Buffer))
    
    'Si Ret devuelve 0 es porque fallo la llamada al Api
    If ret > 0 Then
        ObtainOperativeSystemLanguage = Left$(Buffer, ret - 1)
    Else
        ObtainOperativeSystemLanguage = "No se pudo obtener el idioma del sistema."

    End If
    
End Function

Public Sub SetLanguageApplication()
    '************************************************************************************.
    ' Carga el JSON con las traducciones en un objeto para su uso a lo largo del proyecto
    '************************************************************************************

    Dim LangFile As String
    
    Language = GetVar(Game.path(INIT) & "Config.ini", "Parameters", "Language")
    
    ' Si no se especifica el idioma en el archivo de configuracion, se le pregunta si quiere usar castellano
    ' y escribimos el archivo de configuracion con el idioma seleccionado
    If LenB(Language) = 0 Then
        If MsgBox("Iniciar con idioma Castellano? // Start with Spanish, if you want the game in English press No", vbYesNo, "Argentum Online Libre") = vbYes Then
            Language = "spanish"
        Else
            Language = "english"
        End If

        Call WriteVar(App.path & "\INIT\Config.ini", "Parameters", "Language", Language)
        'Language = LCase$(ObtainOperativeSystemLanguage(LOCALE_SENGLANGUAGE))
    End If
    
    LangFile = FileToString(Game.path(Lenguajes) & Language & ".json")
    Set JsonLanguage = JSON.parse(LangFile)
End Sub
