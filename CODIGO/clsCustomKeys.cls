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
' Creation Date: 08-03-07
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
'Customizes User Keys. Complete revision by Maraxus to fit coding standards and fix bugs.
'
'@author Rapsodius
'@version 1.1.0
'@remarks None
'@date 20070803

Option Explicit

''
'Number of Customizable Keys
Private Const NUM_CUSTOM_KEYS As Byte = 25

''
'Error Raised in case the key file doesn't exist
Private Const ERROR_NO_KEY_FILE As Long = vbObjectError + 16

''
'The relative path from the application's location where the key config file is.
Private Const KEY_CONFIG_FILE As String = "\init\BindKeys.bin"


''
'Array of mapped keys
Private mappedKeys(1 To NUM_CUSTOM_KEYS) As Integer


''
'Index of keys on array
'
'@param mKeyUp Key for moving character up
'@param mKeyDown Key for moving character down
'@param mKeyLeft Key for moving character left
'@param mKeyRight Key for moving character right
'@param mKeyToggleMusic Key for toggling background music
'@param mKeyRequestRefresh Key for sending a request refreshing packet
'@param mKeyToggleNames Key for toggling characters' names
'@param mKeyGetObject Key for retrieving an object from the ground
'@param mKeyToggleCombatMode Key for toggling combat mode
'@param mKeyEquipObject Key for equiping an inventory object
'@param mKeyTamAnimal Key for taming animals
'@param mKeySteal Key for Stealing other users' gold or objects
'@param mKeyToggleSafeMode Key for toggling Safe Mode
'@param mKeyHide Key for hiding the character
'@param mKeyDropObject Key for dropping inventory objects
'@param mKeyUseObject Key for using inventory objects
'@param mKeyAttack Key for attacking other Users/NPCs
'@param mKeyTalk Key for talking (to all)
'@param mKeyTalkWithGuild Key for talking (only to guild members)
'@param mKeyToggleFPS Key for toggling the FPS counter
'@param mKeyTakeScreenShot Key for taking a snapshot of the screen
'@param mKeyShowOptions Key for showing the "Options" form
'@param mKeyMeditate Key for start/stop meditation
'@param mKeyCastSpellMacro Key for start/stop the Cast Spell Macro
'@param mKeyWorkMacro Key for start/stop the Work Macro
Public Enum eKeyType
    '%%%%Movement Keys%%%%
    mKeyUp = 1
    mKeyDown = 2
    mKeyLeft = 3
    mKeyRight = 4
    '%%%%Movement Keys%%%%
    
    '%%%%Option Keys%%%%
    mKeyToggleMusic = 5
    mKeyRequestRefresh = 6
    mKeyToggleNames = 7
    '%%%%Option Keys%%%%
    
    '%%%%Action Keys%%%%
    mKeyGetObject = 8
    mKeyToggleCombatMode = 9
    mKeyEquipObject = 10
    mKeyTamAnimal = 11
    mKeySteal = 12
    mKeyToggleSafeMode = 13
    mKeyHide = 14
    mKeyDropObject = 15
    mKeyUseObject = 16
    mKeyAttack = 17
    '%%%%Action Keys%%%%
    
    '%%%%Talk Keys%%%%
    mKeyTalk = 18
    mKeyTalkWithGuild = 19
    '%%%%Talk Keys%%%%
    
    '%%%%Misc. Keys%%%%
    mKeyToggleFPS = 20
    mKeyTakeScreenShot = 21
    mKeyShowOptions = 22
    mKeyMeditate = 23
    mKeyCastSpellMacro = 24
    mKeyWorkMacro = 25
    '%%%%Misc. Keys%%%%
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
        '%%%%%%%%%%ALPHANUMERIC CHARACTERS%%%%%%%%%%
        Case KeyCodeConstants.vbKeyA To KeyCodeConstants.vbKeyZ 'A-Z
            ReadableName = Chr$(KeyCode)
        
        Case KeyCodeConstants.vbKeyNumpad0 To KeyCodeConstants.vbKeyNumpad9 '0-9
            ReadableName = KeyCode - KeyCodeConstants.vbKeyNumpad0 & " (teclado numérico)"
        
        '%%%%%%%%%%ALPHANUMERIC CHARACTERS%%%%%%%%%%
        
        '%%%%%%%%%%FUNCTION KEYS%%%%%%%%%%
        Case KeyCodeConstants.vbKeyF1 To KeyCodeConstants.vbKeyF16
            'I wonder how many people has keyboards with F16...
            ReadableName = "F" & KeyCode - KeyCodeConstants.vbKeyF1 + 1
        
        '%%%%%%%%%%FUNCTION KEYS%%%%%%%%%%
        
        '%%%%%%%%%%NUMPAD KEYS%%%%%%%%%%
        Case KeyCodeConstants.vbKeyMultiply
            ReadableName = "* (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyAdd
            ReadableName = "+ (teclado numerico)"
        
        Case KeyCodeConstants.vbKeySeparator
            'Note: Separator appears in MSDN as the Enter key for
            'the numpad, but windows recognizes it as the normal
            'Enter key
            ReadableName = "Enter (teclado numérico)"
        
        Case KeyCodeConstants.vbKeySubtract
            ReadableName = "- (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyDecimal
            ReadableName = ". (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyDivide
            ReadableName = "/ (teclado numerico)"
        
        '%%%%%%%%%%NUMPAD KEYS%%%%%%%%%%
        
        '%%%%%%%%%%MODIFIER KEYS%%%%%%%%%%
        Case KeyCodeConstants.vbKeyShift
            ReadableName = "Shift"
        
        Case KeyCodeConstants.vbKeyControl
            ReadableName = "Control"
        
        Case KeyCodeConstants.vbKeyMenu 'Alt
            ReadableName = "Alt"
        
        '%%%%%%%%%%MODIFIER KEYS%%%%%%%%%%
        
        '%%%%%%%%%%MISC. KEYS%%%%%%%%%%
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
        '%%%%%%%%%%MISC. KEYS%%%%%%%%%%
        
        Case Else
            'In all other cases, key is considered not valid.
            ReadableName = Constants.vbNullString
    End Select
End Function

''
'Reads custom keys from a file

Public Sub LoadCustomKeys()
'**************************************************************
' LoadCustomKeys - Reads binded keys from file, raises an
' error if file doesn't exist or if there is an invalid key
'
' Creation Date: 08-21-07
' Created and Implemented by Rapsodius
'**************************************************************
    Dim i As Long
    Dim FileNum As Integer
    FileNum = FreeFile()
        
    Open App.path & KEY_CONFIG_FILE For Binary Access Read Lock Read Write As FileNum
    
    For i = 1 To NUM_CUSTOM_KEYS
        Get FileNum, , mappedKeys(i)
        If LenB(ReadableName(mappedKeys(i))) = 0 Then
            Close FileNum
            Call Err.Raise(ERROR_NO_KEY_FILE)
            Exit Sub
        End If
    Next i
    
    Close FileNum
End Sub

''
' Saves custom keys to a file

Public Sub SaveCustomKeys()
'**************************************************************
' SaveCustomKeys - Saves binded keys to the file
'
' Creation Date: 08-04-07
' Created and Implemented by Rapsodius
'**************************************************************
    Dim FileNum As Integer
    FileNum = FreeFile()
    
    Open App.path & KEY_CONFIG_FILE For Binary Access Write As FileNum
    
    Put FileNum, , mappedKeys()
    
    Close #FileNum
End Sub

''
'Gets a key code
'
'@param index Index of the key code to retrieve
'@return The Key code that belongs to index

Public Property Get BindedKey(ByVal index As eKeyType) As Integer
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
' Creation Date: 08-04-07
' Created and Implemented by Rapsodius
'**************************************************************
    If index < 1 Or index > NUM_CUSTOM_KEYS Then Exit Property
    BindedKey = mappedKeys(index)
End Property

''
'Sets a key code
'
'@param index Index of the key to change
'@param NewVal New key code

Public Property Let BindedKey(ByVal index As eKeyType, ByVal NewVal As Integer)
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
' Creation Date: 08-05-07
' Last Modification: 08-14-07 by Rapsodius
' Created and Implemented by Rapsodius
'**************************************************************
    If index < 1 Or index > NUM_CUSTOM_KEYS Then Exit Property
    
    If LenB(ReadableName(NewVal)) = 0 Then 'If key is not valid...
        Exit Property                      'Rejects the new assignment
    End If
    
    If KeyAssigned(NewVal) Then 'Also reject it in case key is already assigned
        Exit Property
    End If
    
    mappedKeys(index) = NewVal
End Property

''
'Loads default keys

Public Sub LoadDefaults()
'**************************************************************
' LoadDefaults - Loads default key codes in case of error
' or if the user requests it.
'
' Creation Date: 08-05-07
' Created and Implemented by Rapsodius
'**************************************************************
    mappedKeys(eKeyType.mKeyUp) = KeyCodeConstants.vbKeyUp
    mappedKeys(eKeyType.mKeyDown) = KeyCodeConstants.vbKeyDown
    mappedKeys(eKeyType.mKeyLeft) = KeyCodeConstants.vbKeyLeft
    mappedKeys(eKeyType.mKeyRight) = KeyCodeConstants.vbKeyRight
    mappedKeys(eKeyType.mKeyToggleMusic) = KeyCodeConstants.vbKeyM
    mappedKeys(eKeyType.mKeyRequestRefresh) = KeyCodeConstants.vbKeyL
    mappedKeys(eKeyType.mKeyToggleNames) = KeyCodeConstants.vbKeyN
    mappedKeys(eKeyType.mKeyGetObject) = KeyCodeConstants.vbKeyA
    mappedKeys(eKeyType.mKeyToggleCombatMode) = KeyCodeConstants.vbKeyC
    mappedKeys(eKeyType.mKeyEquipObject) = KeyCodeConstants.vbKeyE
    mappedKeys(eKeyType.mKeyTamAnimal) = KeyCodeConstants.vbKeyD
    mappedKeys(eKeyType.mKeySteal) = KeyCodeConstants.vbKeyR
    mappedKeys(eKeyType.mKeyToggleSafeMode) = KeyCodeConstants.vbKeyMultiply
    mappedKeys(eKeyType.mKeyHide) = KeyCodeConstants.vbKeyO
    mappedKeys(eKeyType.mKeyDropObject) = KeyCodeConstants.vbKeyT
    mappedKeys(eKeyType.mKeyUseObject) = KeyCodeConstants.vbKeyU
    mappedKeys(eKeyType.mKeyAttack) = KeyCodeConstants.vbKeyControl
    mappedKeys(eKeyType.mKeyTalk) = KeyCodeConstants.vbKeyReturn
    mappedKeys(eKeyType.mKeyTalkWithGuild) = KeyCodeConstants.vbKeyDelete
    mappedKeys(eKeyType.mKeyToggleFPS) = KeyCodeConstants.vbKeyF4
    mappedKeys(eKeyType.mKeyTakeScreenShot) = KeyCodeConstants.vbKeyF2
    mappedKeys(eKeyType.mKeyShowOptions) = KeyCodeConstants.vbKeyF5
    mappedKeys(eKeyType.mKeyMeditate) = KeyCodeConstants.vbKeyF6
    mappedKeys(eKeyType.mKeyCastSpellMacro) = KeyCodeConstants.vbKeyF7
    mappedKeys(eKeyType.mKeyWorkMacro) = KeyCodeConstants.vbKeyF8
End Sub

''
'Returns whether a key is binded to an action
'
'@param KeyCode Key code to check
'@return True if key is mapped, false otherwise

Public Function KeyAssigned(ByVal KeyCode As Integer) As Boolean
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
' Creation Date: 08-07-07
' Created and Implemented by Rapsodius
'**************************************************************
    Dim Counter As Long
    
    KeyAssigned = False
    
    For Counter = 1 To NUM_CUSTOM_KEYS
        If mappedKeys(Counter) = KeyCode Then
            KeyAssigned = True
            Exit For
        End If
    Next Counter
End Function

''
'Retrieves the number of customizable keys
'
'@return The number of customizable keys

Public Property Get Count() As Byte
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Count = NUM_CUSTOM_KEYS
End Property

''
'Constructor. Loads Keys from file

Private Sub Class_Initialize()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
On Error GoTo ErrHand:
    Call LoadCustomKeys
    Exit Sub
    
ErrHand:
    If Err.Number = ERROR_NO_KEY_FILE Then
        Call LoadDefaults
        Call SaveCustomKeys
        Resume Next
    End If
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
