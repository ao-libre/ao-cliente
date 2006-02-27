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

Private p_CANTIDADDIALOGOS As Byte
Private Const p_SEGUNDOS = 7
Private Type t_GuildDlg
    Texto As String
    Segundos As Byte
End Type
Private p_Dialogos() As t_GuildDlg
Private p_Activo As Boolean

Public Property Let CantidadDialogos(ByVal v As Byte)
    If v > 0 Then
    ReDim Preserve p_Dialogos(1 To v) As t_GuildDlg
    End If
    p_CANTIDADDIALOGOS = v
End Property

Public Property Let Activo(ByVal v As Boolean)
    p_Activo = v
End Property

Public Property Get Activo() As Boolean
    Activo = p_Activo
End Property

Public Sub Class_Initialize()
    p_CANTIDADDIALOGOS = 5
    ReDim p_Dialogos(1 To p_CANTIDADDIALOGOS) As t_GuildDlg
    p_Activo = True
End Sub

Public Sub Draw(ByRef C As cDialogos)
Dim I As Byte
Dim Y As Long

If Not p_Activo Then Exit Sub

Y = 250
'500 600 700 800 900
    For I = 1 To p_CANTIDADDIALOGOS
        If Trim$(p_Dialogos(I).Texto) <> vbNullString Then
            If Left$(p_Dialogos(I).Texto, 1) = Chr(3) Then
                Call C.DrawText(260, Y + (I * 10), Mid(p_Dialogos(I).Texto, 2), vbGreen)
            Else
                Call C.DrawText(260, Y + (I * 10), p_Dialogos(I).Texto, vbYellow)
            End If
        End If
    Next I
End Sub

Public Sub PassTimer()
Dim I As Byte

For I = 1 To p_CANTIDADDIALOGOS
    If p_Dialogos(I).Segundos > 0 Then
        p_Dialogos(I).Segundos = p_Dialogos(I).Segundos - 1
    Else
        p_Dialogos(I).Texto = vbNullString
    End If
Next I
    
End Sub

Public Sub PushBackText(ByVal s As String)
Dim I As Byte
Dim Vacio As Boolean
Dim Multiplicador As Integer
If Left$(s, 1) = Chr(3) Then
    Multiplicador = 7
Else
    Multiplicador = 1
End If

If p_CANTIDADDIALOGOS > 0 Then
    I = p_CANTIDADDIALOGOS
    Vacio = True
    While I > 0 And Vacio
        Vacio = p_Dialogos(I).Texto = vbNullString
        If Vacio Then I = I - 1
    Wend
    If I = p_CANTIDADDIALOGOS Then
        'hay q scrollear, estamos llenos
        I = 1
        While I < p_CANTIDADDIALOGOS
            p_Dialogos(I) = p_Dialogos(I + 1)
            I = I + 1
        Wend
        p_Dialogos(I).Texto = s
        p_Dialogos(I).Segundos = p_SEGUNDOS * Multiplicador
    Else
        p_Dialogos(I + 1).Texto = s
        p_Dialogos(I + 1).Segundos = p_SEGUNDOS * Multiplicador
    End If
End If

End Sub

