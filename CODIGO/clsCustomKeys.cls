VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsCustomKeys.cls - Allows the User to Customize Keys.
'
' Creation Date: 08/03/07
' Created and Implemented by Rapsodius
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
'Customizes User Keys.
'
'@author Rapsodius
'@version 1.1.0
'@date 20080413

'03/08/2007 - Rapsodius
'   - First Release
'
'13/04/2008 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Removed lots of unwanted buggy features keeping it clean. Revised documentation and coding standards.


Option Explicit

''
'Number of Customizable Keys
Private Const NUM_CUSTOM_KEYS As Byte = 27

''
'Number of Key Configurations
Private Const NUM_CUSTOM_CONFIGS As Byte = 9

''
'Error Raised in case the key file doesn't exist
Private Const ERROR_NO_KEY_FILE As Long = vbObjectError + 16

''
'The relative path from the application's location where the key config file is.
Private Const KEY_CONFIG_FILE As String = "\init\BindKeys.bin"

''
'Array of mapped keys
Private mappedKeys(0 To NUM_CUSTOM_CONFIGS, 1 To NUM_CUSTOM_KEYS) As Integer

''
'Active Key Configuration
Private ActiveConfig As Byte


''
'Index of keys on array
'
'@param mKeyUp Key for moving character up
'@param mKeyDown Key for moving character down
'@param mKeyLeft Key for moving character left
'@param mKeyRight Key for moving character right
'@param mKeyToggleMusic Key for toggling background music
'@param mKeyToggleSound Key for toggling sound
'@param mKeyToggleFxs Key for togglig the 3D Fxs
'@param mKeyRequestRefresh Key for sending a request refreshing packet
'@param mKeyToggleNames Key for toggling characters' names
'@param mKeyGetObject Key for retrieving an object from the ground
'@param mKeyEquipObject Key for equiping an inventory object
'@param mKeyTamAnimal Key for taming animals
'@param mKeySteal Key for Stealing other users' gold or objects
'@param mKeyToggleSafeMode Key for toggling Safe Mode
'@param mKeyToggleResuscitationSafe Key for toggling Resuscitation Safe
'@param mKeyHide Key for hiding the character
'@param mKeyDropObject Key for dropping inventory objects
'@param mKeyUseObject Key for using inventory objects
'@param mKeyAttack Key for attacking other Users/NPCs
'@param mKeyTalk Key for talking (to all)
'@param mKeyTalkWithGuild Key for talking (only to guild members)
'@param mKeyTakeScreenShot Key for taking a snapshot of the screen
'@param mKeyShowOptions Key for showing the "Options" form
'@param mKeyMeditate Key for start/stop meditation
'@param mKeyCastSpellMacro Key for start/stop the Cast Spell Macro
'@param mKeyWorkMacro Key for start/stop the Work Macro
'@param mKeyExitGame Key for close the game
Public Enum eKeyType
    mKeyUp = 1
    mKeyDown
    mKeyLeft
    mKeyRight
    mKeyToggleMusic
    mKeyToggleSound
    mKeyToggleFxs
    mKeyRequestRefresh
    mKeyToggleNames
    mKeyGetObject
    mKeyEquipObject
    mKeyTamAnimal
    mKeySteal
    mKeyToggleSafeMode
    mKeyToggleResuscitationSafe
    mKeyHide
    mKeyDropObject
    mKeyUseObject
    mKeyAttack
    mKeyTalk
    mKeyTalkWithGuild
    mKeyTakeScreenShot
    mKeyShowOptions
    mKeyMeditate
    mKeyCastSpellMacro
    mKeyWorkMacro
    mKeyExitGame
End Enum

''
'Shows a readable name for a specific Key code
'
'@param KeyCode Key code to process
'@return An User readable name for the key code or a null string if the key is not valid
'@remarks All keys not listed here are considered not valid

Public Function ReadableName(ByVal KeyCode As Integer) As String
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyA To KeyCodeConstants.vbKeyZ
            ReadableName = Chr$(KeyCode)
        
        Case KeyCodeConstants.vbKeyNumpad0 To KeyCodeConstants.vbKeyNumpad9
            ReadableName = KeyCode - KeyCodeConstants.vbKeyNumpad0 & " (teclado numérico)"
        
        Case KeyCodeConstants.vbKeyF1 To KeyCodeConstants.vbKeyF16
            ReadableName = "F" & (KeyCode - KeyCodeConstants.vbKeyF1 + 1)
        
        Case KeyCodeConstants.vbKeyMultiply
            ReadableName = "* (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyAdd
            ReadableName = "+ (teclado numerico)"
        
        Case KeyCodeConstants.vbKeySeparator
            'Note: Separator appears in MSDN as the Enter key for
            'the numpad, but MS recognizes it as the normal
            'Enter key
            ReadableName = "Enter (teclado numérico)"
        
        Case KeyCodeConstants.vbKeySubtract
            ReadableName = "- (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyDecimal
            ReadableName = ". (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyDivide
            ReadableName = "/ (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyShift
            ReadableName = "Shift"
        
        Case KeyCodeConstants.vbKeyControl
            ReadableName = "Control"
        
        Case KeyCodeConstants.vbKeyMenu 'Alt
            ReadableName = "Alt"
        
        Case KeyCodeConstants.vbKeyPause
            ReadableName = "Pausa"
        
        Case KeyCodeConstants.vbKeyCapital
            ReadableName = "Blq Mayus"
        
        Case KeyCodeConstants.vbKeyEscape
            ReadableName = "Esc"
        
        Case KeyCodeConstants.vbKeyPageUp
            ReadableName = "Ret Pag"
        
        Case KeyCodeConstants.vbKeyPageDown
            ReadableName = "Av Pag"
        
        Case KeyCodeConstants.vbKeyEnd
            ReadableName = "Fin"
        
        Case KeyCodeConstants.vbKeyHome
            ReadableName = "Inicio"
        
        Case KeyCodeConstants.vbKeyLeft
            ReadableName = "Izquierda"
        
        Case KeyCodeConstants.vbKeyUp
            ReadableName = "Arriba"
        
        Case KeyCodeConstants.vbKeyRight
            ReadableName = "Derecha"
        
        Case KeyCodeConstants.vbKeyDown
            ReadableName = "Abajo"
        
        Case KeyCodeConstants.vbKeyInsert
            ReadableName = "Insertar"
        
        Case KeyCodeConstants.vbKeyNumlock
            ReadableName = "Blq Num"
        
        Case KeyCodeConstants.vbKeyScrollLock
            ReadableName = "Blq Despl"
        
        Case KeyCodeConstants.vbKeySpace
            ReadableName = "Barra Espaciadora"
        
        Case KeyCodeConstants.vbKeyBack
            ReadableName = "Borrar"
        
        Case KeyCodeConstants.vbKeyReturn
            ReadableName = "Enter"
        
        Case KeyCodeConstants.vbKeyDelete
            ReadableName = "Suprimir"
        
        Case Else
            'In all other cases, key is considered not valid.
            ReadableName = Constants.vbNullString
    End Select
End Function

''
'Reads custom keys from a file.
'
'@remarks   If the custom keys file is not found ERROR_NO_KEY_FILE is raised.

Public Sub LoadCustomKeys()
'***************************************************
'Author: Rapsodius
'Last Modification: 18/11/2010
'16/11/2010: Amraphen - Adapté el procedimiento para cargar todas las configuraciones de teclas.
'18/11/2010: Amraphen - Modifiqué la forma en que se cargan los bindings y ahora carga los del formato viejo.
'***************************************************
Dim i As Long
Dim j As Long
Dim FileNum As Integer

     'Si no existe el archivo lo creamos con los bindings default:
    If Not FileExist(App.path & KEY_CONFIG_FILE, vbArchive) Then
        For i = 0 To NUM_CUSTOM_CONFIGS
            Call LoadDefaults(i)
        Next i
        
        ActiveConfig = 1
        
        Call SaveCustomKeys
    Else 'Existe el archivo, verificamos qué formato tiene:
        
        FileNum = FreeFile
    
        If FileLen(App.path & KEY_CONFIG_FILE) < (NUM_CUSTOM_CONFIGS * NUM_CUSTOM_KEYS * 2) + 1 Then
            'Formato viejo
            Open App.path & KEY_CONFIG_FILE For Binary Access Read Lock Read Write As FileNum
            
            ActiveConfig = 1
            
            For i = 0 To NUM_CUSTOM_CONFIGS
                Call LoadDefaults(i)
            Next i
            
            For j = 1 To NUM_CUSTOM_KEYS
                Get FileNum, , mappedKeys(ActiveConfig, j)
            Next j
        Else
            'Formato nuevo
            Open App.path & KEY_CONFIG_FILE For Binary Access Read Lock Read Write As FileNum
            
            Get FileNum, , ActiveConfig
            
            For i = 0 To NUM_CUSTOM_CONFIGS
                For j = 1 To NUM_CUSTOM_KEYS
                    Get FileNum, , mappedKeys(i, j)
                Next j
            Next i
        End If
        
        Close FileNum
        
        Call SaveCustomKeys
    End If
End Sub

''
' Saves custom keys to a file

Public Sub SaveCustomKeys()
'***************************************************
'Author: Rapsodius
'Last Modification: 18/11/10
'16/11/2010: Amraphen - Adapté el procedimiento para guardar todas las configuraciones de teclas.
'18/11/2010: Amraphen - Ahora se guarda la configuración activa.
'***************************************************
Dim i As Long
Dim j As Long
Dim FileNum As Integer

    FileNum = FreeFile()
    
    Open App.path & KEY_CONFIG_FILE For Binary Access Write As FileNum
    
    'Guardo la configuración activa:
    Put FileNum, , ActiveConfig
    
    'Guardo las distintas configuraciones:
    For i = 0 To NUM_CUSTOM_CONFIGS
        For j = 1 To NUM_CUSTOM_KEYS
            Put FileNum, , mappedKeys(i, j)
        Next j
    Next i
    
    Close FileNum
End Sub

''
'Gets a key code
'
'@param index Index of the key code to retrieve
'@return The Key code that belongs to index

Public Property Get BindedKey(ByVal Index As eKeyType) As Integer
'***************************************************
'Author: Rapsodius
'Last Modification: 16/11/2010
'16/11/2010: Amraphen - Adapté la propiedad para que cargue el bind de la tecla en la configuración activa.
'***************************************************
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
    
    BindedKey = mappedKeys(ActiveConfig, Index)
End Property

''
'Sets a key code
'
'@param index Index of the key to change
'@param NewVal New key code

Public Property Let BindedKey(ByVal Index As eKeyType, ByVal NewVal As Integer)
'***************************************************
'Author: Rapsodius
'Last Modification: 16/11/2010
'16/11/2010: Amraphen - Adapté la propiedad para que guarde el bind de la tecla en la configuración activa.
'***************************************************
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
    
    If LenB(ReadableName(NewVal)) = 0 Then 'If key is not valid...
        Exit Property                      'Rejects the new assignment
    End If
    
    If KeyAssigned(NewVal) Then 'Also reject it in case key is already assigned
        Exit Property
    End If
    
    mappedKeys(ActiveConfig, Index) = NewVal
End Property

''
'Loads default keys
'
'@param KeyConfig Index of the key configuration to change

Public Sub LoadDefaults(ByVal KeyConfig As Byte)
'***************************************************
'Author: Rapsodius
'Last Modification: 16/11/2010
'16/11/2010: Amraphen - Adapté el procedimiento para que cargue los bindings por default en una sola configuración.
'***************************************************
    mappedKeys(KeyConfig, eKeyType.mKeyUp) = KeyCodeConstants.vbKeyUp
    mappedKeys(KeyConfig, eKeyType.mKeyDown) = KeyCodeConstants.vbKeyDown
    mappedKeys(KeyConfig, eKeyType.mKeyLeft) = KeyCodeConstants.vbKeyLeft
    mappedKeys(KeyConfig, eKeyType.mKeyRight) = KeyCodeConstants.vbKeyRight
    mappedKeys(KeyConfig, eKeyType.mKeyToggleMusic) = KeyCodeConstants.vbKeyM
    mappedKeys(KeyConfig, eKeyType.mKeyToggleSound) = KeyCodeConstants.vbKeyS
    mappedKeys(KeyConfig, eKeyType.mKeyToggleFxs) = KeyCodeConstants.vbKeyF
    mappedKeys(KeyConfig, eKeyType.mKeyRequestRefresh) = KeyCodeConstants.vbKeyL
    mappedKeys(KeyConfig, eKeyType.mKeyToggleNames) = KeyCodeConstants.vbKeyN
    mappedKeys(KeyConfig, eKeyType.mKeyGetObject) = KeyCodeConstants.vbKeyA
    mappedKeys(KeyConfig, eKeyType.mKeyEquipObject) = KeyCodeConstants.vbKeyE
    mappedKeys(KeyConfig, eKeyType.mKeyTamAnimal) = KeyCodeConstants.vbKeyD
    mappedKeys(KeyConfig, eKeyType.mKeySteal) = KeyCodeConstants.vbKeyR
    mappedKeys(KeyConfig, eKeyType.mKeyToggleSafeMode) = KeyCodeConstants.vbKeyMultiply
    mappedKeys(KeyConfig, eKeyType.mKeyToggleResuscitationSafe) = KeyCodeConstants.vbKeyEnd
    mappedKeys(KeyConfig, eKeyType.mKeyHide) = KeyCodeConstants.vbKeyO
    mappedKeys(KeyConfig, eKeyType.mKeyDropObject) = KeyCodeConstants.vbKeyT
    mappedKeys(KeyConfig, eKeyType.mKeyUseObject) = KeyCodeConstants.vbKeyU
    mappedKeys(KeyConfig, eKeyType.mKeyAttack) = KeyCodeConstants.vbKeyControl
    mappedKeys(KeyConfig, eKeyType.mKeyTalk) = KeyCodeConstants.vbKeyReturn
    mappedKeys(KeyConfig, eKeyType.mKeyTalkWithGuild) = KeyCodeConstants.vbKeyDelete
    mappedKeys(KeyConfig, eKeyType.mKeyTakeScreenShot) = KeyCodeConstants.vbKeyF2
    mappedKeys(KeyConfig, eKeyType.mKeyShowOptions) = KeyCodeConstants.vbKeyF5
    mappedKeys(KeyConfig, eKeyType.mKeyMeditate) = KeyCodeConstants.vbKeyF6
    mappedKeys(KeyConfig, eKeyType.mKeyCastSpellMacro) = KeyCodeConstants.vbKeyF7
    mappedKeys(KeyConfig, eKeyType.mKeyWorkMacro) = KeyCodeConstants.vbKeyF8
    mappedKeys(KeyConfig, eKeyType.mKeyExitGame) = KeyCodeConstants.vbKeyF12
End Sub

''
'Returns whether a key is binded to an action
'
'@param KeyCode Key code to check
'@return True if key is mapped, false otherwise

Public Function KeyAssigned(ByVal KeyCode As Integer) As Boolean
'***************************************************
'Author: Rapsodius
'Last Modification: 16/11/2010
'16/11/2010: Amraphen - Adapté la función para que trabaje sobre la configuración activa.
'***************************************************
Dim Counter As Long
    
    KeyAssigned = False
    
    For Counter = 1 To NUM_CUSTOM_KEYS
        If mappedKeys(ActiveConfig, Counter) = KeyCode Then
            KeyAssigned = True
            Exit For
        End If
    Next Counter
End Function

''
'Retrieves the number of customizable keys
'
'@return The number of customizable keys

Public Property Get KeyCount() As Byte
'***************************************************
'Author: Rapsodius
'Last Modification: 16/11/2010
'16/11/2010: Amraphen - Cambié el nombre para evitar confusiones.
'***************************************************
    KeyCount = NUM_CUSTOM_KEYS
End Property

'
'@return The number of customizable key configurations

Public Property Get ConfigCount() As Byte
'***************************************************
'Author: Amraphen
'Last Modification: 16/11/2010
'
'***************************************************
    ConfigCount = NUM_CUSTOM_CONFIGS
End Property

''
'Sets the currently active key configuration index
'
'@param NewVal New currently active key configuration index

Public Property Let CurrentConfig(ByVal NewVal As Byte)
'***************************************************
'Author: Amraphen
'Last Modification: 16/11/2010
'
'***************************************************
    If NewVal < 0 Or NewVal > NUM_CUSTOM_CONFIGS Then Exit Property
    
    ActiveConfig = NewVal
End Property

''
'Retrieves the currently active key configuration index
'
'@return The currently active key configuration index

Public Property Get CurrentConfig() As Byte
'***************************************************
'Author: Amraphen
'Last Modification: 16/11/2010
'
'***************************************************
    CurrentConfig = ActiveConfig
End Property

''
'Constructor. Loads Keys from file

Private Sub Class_Initialize()
'***************************************************
'Author: Rapsodius
'Last Modification: 18/11/2010
'16/11/2010: Amraphen - Adapté el procedimiento para que trabaje sobre las distintas configuraciones de teclas.
'18/11/2010: Amraphen - Ahora el manejo de errores se hace desde el sub LoadCustomKeys
'***************************************************

    Call LoadCustomKeys
End Sub

''
'Destructor. Saves Keys to file

Private Sub Class_Terminate()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Call SaveCustomKeys
End Sub

