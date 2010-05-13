VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuildDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const p_SEGUNDOS = 7

Private p_CANTIDADDIALOGOS As Byte

Private Type t_GuildDlg
    Texto As String
    Segundos As Byte
End Type

Private p_Dialogos() As t_GuildDlg
Private p_Activo As Boolean

Private usedFont As StdFont

Public Property Let CantidadDialogos(ByVal v As Byte)
    If v > 0 Then
        ReDim Preserve p_Dialogos(1 To v) As t_GuildDlg
        p_CANTIDADDIALOGOS = v
    End If
End Property

''
' Removes all dialogs from screen.

Public Sub RemoveDialogs()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/02/07
'Removes all dialogs
'***************************************************
    Dim i As Long
    
    For i = 1 To p_CANTIDADDIALOGOS
        p_Dialogos(i).Texto = vbNullString
    Next i
End Sub

''
' Retrieves the number of dialogs to be displayed on screen.
'
' @return   The number of dialogs to be displayed on screen.

Public Property Get CantidadDialogos() As Byte
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 04/01/07
'Retrieves the number of dialogs to be displayed on screen
'***************************************************
    CantidadDialogos = p_CANTIDADDIALOGOS
End Property

Public Property Let Activo(ByVal v As Boolean)
    p_Activo = v
End Property

Public Property Get Activo() As Boolean
    Activo = p_Activo
End Property

Private Sub Class_Initialize()
    p_CANTIDADDIALOGOS = 5
    ReDim p_Dialogos(1 To p_CANTIDADDIALOGOS) As t_GuildDlg
    p_Activo = True
End Sub

Public Sub Draw()
Dim i As Byte
Dim Y As Long

If Not p_Activo Then Exit Sub

Y = 250
'500 600 700 800 900
    For i = 1 To p_CANTIDADDIALOGOS
        If Trim$(p_Dialogos(i).Texto) <> vbNullString Then
            If Left$(p_Dialogos(i).Texto, 1) = Chr(3) Then
                Call Mod_TileEngine.RenderText(260, Y + (i * 10), mid(p_Dialogos(i).Texto, 2), vbGreen, usedFont)
            Else
                Call Mod_TileEngine.RenderText(260, Y + (i * 10), p_Dialogos(i).Texto, vbYellow, usedFont)
            End If
        End If
    Next i
End Sub

Public Sub PassTimer()
Dim i As Byte

For i = 1 To p_CANTIDADDIALOGOS
    If p_Dialogos(i).Segundos > 0 Then
        p_Dialogos(i).Segundos = p_Dialogos(i).Segundos - 1
    Else
        p_Dialogos(i).Texto = vbNullString
    End If
Next i
    
End Sub

''
' Splits dialogs into lines fitting properly the render area and inserts them

Public Sub PushBackText(ByVal s As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/04/2009
'Splits dialogs into lines fitting properly the render area and inserts them
'07/04/2009: Now cuts the string properly in spite of not fitting in the screen. This avoids an infite loop.
'***************************************************
    Dim str As String
    Dim tmp As Integer
    
    str = s
    
    ' If it's too long to fit, split it
    Do While frmMain.TextWidth(str) > 500
        tmp = InStrRev(str, " ")
        
        If tmp = 0 Then Exit Do
        
        str = Left$(str, tmp - 1)
    Loop
    
    'Show message and continue with the rest
    Call RealPushBackText(str)
    
    If LenB(str) <> LenB(s) Then
        Call PushBackText(Right$(s, Len(s) - Len(str) - 1))
    End If
End Sub

Private Sub RealPushBackText(ByVal s As String)
Dim i As Byte
Dim Vacio As Boolean

If p_CANTIDADDIALOGOS > 0 Then
    i = p_CANTIDADDIALOGOS
    Vacio = True
    While i > 0 And Vacio
        Vacio = p_Dialogos(i).Texto = vbNullString
        If Vacio Then i = i - 1
    Wend
    If i = p_CANTIDADDIALOGOS Then
        'hay q scrollear, estamos llenos
        i = 1
        While i < p_CANTIDADDIALOGOS
            p_Dialogos(i) = p_Dialogos(i + 1)
            i = i + 1
        Wend
        p_Dialogos(i).Texto = s
        p_Dialogos(i).Segundos = p_SEGUNDOS
    Else
        p_Dialogos(i + 1).Texto = s
        p_Dialogos(i + 1).Segundos = p_SEGUNDOS
    End If
End If

End Sub

''
' Retrieves the font being used to render dialogs.
'
' @return   The font being used to render dialogs.

Public Property Get font() As StdFont
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 04/27/06
'Retrieves the font being used to render dialogs.
'**************************************************************
    font = usedFont
End Property

''
' Sets the font being used to render dialogs.
'
' @param    nValue  The new font to be used to render dialogs.

Public Property Let font(ByRef nValue As StdFont)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 04/27/06
'Sets the font being used to render dialogs.
'**************************************************************
    Set usedFont = nValue
End Property
