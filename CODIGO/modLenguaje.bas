Attribute VB_Name = "ModLenguaje"
Option Explicit

'Constantes para el Api GetLocaleInfo
'************************************
Const LOCALE_USER_DEFAULT = &H400
Const LOCALE_SENGCOUNTRY = &H1002
Const LOCALE_SENGLANGUAGE = &H1001
Const LOCALE_SNATIVELANGNAME = &H4
Const LOCALE_SNATIVECTRYNAME = &H8
  
'Declaración de la función Api GetLocaleInfo
Private Declare Function GetLocaleInfo _
                Lib "kernel32" _
                Alias "GetLocaleInfoA" (ByVal Locale As Long, _
                                        ByVal LCType As Long, _
                                        ByVal lpLCData As String, _
                                        ByVal cchData As Long) As Long

Public JsonLanguage As Object

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
    ' Función que obtiene el idioma del sistema
    '*******************************************

    Dim Buffer As String, ret As String

    Buffer = String$(256, 0)
            
    ret = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, Buffer, Len(Buffer))
    
    'Si Ret devuelve 0 es porque falló la llamada al Api
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
    Dim Language As String
    
    Language = GetVar(DirInit & "Config.ini", "Parameters", "Language")
    
    ' Si no se especifica el idioma en el archivo de configuracion, se usa el del sistema
    If LenB(Language) = 0 Then
        MsgBox "No se pudo cargar el idioma seleccionado. Se usara el idioma que tienes en el sistema.", vbOKOnly, "Carga de Idiomas"
        Language = "spanish"
        'Language = LCase$(ObtainOperativeSystemLanguage(LOCALE_SENGLANGUAGE))
    End If
    
    LangFile = FileToString(DirLenguages & Language & ".json")
    Set JsonLanguage = JSON.parse(LangFile)
End Sub


