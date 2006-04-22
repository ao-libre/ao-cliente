VERSION 5.00
Begin VB.Form frmPanelGm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel GM"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4635
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Ver oro en banco"
      Height          =   315
      Index           =   19
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Show SOS"
      Height          =   315
      Index           =   18
      Left            =   3420
      TabIndex        =   21
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Boveda"
      Height          =   315
      Index           =   17
      Left            =   2340
      TabIndex        =   20
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Ban X ip"
      Height          =   315
      Index           =   16
      Left            =   1260
      TabIndex        =   19
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Penas"
      Height          =   315
      Index           =   15
      Left            =   180
      TabIndex        =   18
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "IP 2 NICK"
      Height          =   315
      Index           =   14
      Left            =   1260
      TabIndex        =   17
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "NICK 2 IP"
      Height          =   315
      Index           =   13
      Left            =   180
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "UNBAN"
      Height          =   315
      Index           =   12
      Left            =   3420
      TabIndex        =   15
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "CARCEL"
      Height          =   315
      Index           =   11
      Left            =   3420
      TabIndex        =   14
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "SKILLS"
      Height          =   315
      Index           =   10
      Left            =   1260
      TabIndex        =   13
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "INV"
      Height          =   315
      Index           =   9
      Left            =   180
      TabIndex        =   12
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "INFO"
      Height          =   315
      Index           =   8
      Left            =   3420
      TabIndex        =   11
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "N.ENE."
      Height          =   315
      Index           =   7
      Left            =   180
      TabIndex        =   10
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "DONDE"
      Height          =   315
      Index           =   6
      Left            =   3420
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "HORA"
      Height          =   315
      Index           =   5
      Left            =   2340
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Guardar comentario"
      Height          =   315
      Index           =   4
      Left            =   180
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "IRA"
      Height          =   315
      Index           =   3
      Left            =   1260
      TabIndex        =   6
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "SUM"
      Height          =   315
      Index           =   2
      Left            =   2340
      TabIndex        =   5
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "BAN"
      Height          =   315
      Index           =   1
      Left            =   2340
      TabIndex        =   4
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "ECHAR"
      Height          =   315
      Index           =   0
      Left            =   2340
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "Actualiza"
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   4035
   End
   Begin VB.ComboBox cboListaUsus 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3435
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   120
      X2              =   120
      Y1              =   540
      Y2              =   1380
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   4440
      X2              =   4440
      Y1              =   540
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2280
      X2              =   2280
      Y1              =   960
      Y2              =   1380
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2280
      X2              =   4440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4440
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   2280
      Y1              =   1380
      Y2              =   1380
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccion_Click(index As Integer)
Dim Ok As Boolean, Tmp As String, Tmp2 As String
Dim Nick As String

Nick = cboListaUsus.Text

Select Case index
Case 0 '/ECHAR nick
    Call SendData("/ECHAR " & Nick)
Case 1 '/ban motivo@nick
    Tmp = InputBox("Motivo ?", "")
    If MsgBox("Esta seguro que desea banear al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/BAN " & Tmp & "@" & Nick)
    End If
Case 2 '/sum nick
    Call SendData("/SUM " & Nick)
Case 3 '/ira nick
    Call SendData("/IRA " & Nick)
Case 4 '/rem
    Tmp = InputBox("Comentario ?", "")
    Call SendData("/REM " & Tmp)
Case 5 '/hora
    Call SendData("/HORA")
Case 6 '/donde nick
    Call SendData("/DONDE " & Nick)
Case 7 '/nene
    Tmp = InputBox("Mapa ?", "")
    Call SendData("/NENE " & Trim(Tmp))
Case 8 '/info nick
    Call SendData("/INFO " & Nick)
Case 9 '/inv nick
    Call SendData("/INV " & cboListaUsus.Text)
Case 10 '/skills nick
    Call SendData("/SKILLS " & Nick)
Case 11 '/carcel minutos nick
    Tmp = InputBox("Minutos ? (hasta 30)", "")
    Tmp2 = InputBox("Razon ?", "")
    If MsgBox("Esta seguro que desea encarcelar al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/CARCEL " & Nick & "@" & Tmp2 & "@" & Tmp)
    End If
Case 12 '/unban nick
    If MsgBox("Esta seguro que desea removerle el ban al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/UNBAN " & Nick)
    End If
Case 13 '/nick2ip nick
    Call SendData("/NICK2IP " & Nick)
Case 14 '/ip2nick ip
    Call SendData("/IP2NICK " & Nick)
Case 15 '/penas
    Call SendData("/PENAS " & cboListaUsus.Text)
Case 16 'Ban X ip
    Tmp = InputBox("Ingrese el motivo del ban", "Ban X IP")
    If MsgBox("Esta seguro que desea banear el (ip o personaje) " & Nick & "Por IP?", vbYesNo) = vbYes Then
        Nick = Replace(Nick, " ", "+")
        Call SendData("/BANIP " & Nick & Tmp)
    End If
Case 17 ' MUESTA BOBEDA
    Call SendData("/BOV " & Nick)
Case 18 ' Sos
    Call SendData("/SHOW SOS")
Case 19 ' Balance
    Call SendData("/BAL " & cboListaUsus.Text)
End Select


End Sub

Private Sub cmdActualiza_Click()
Call SendData("LISTUSU")

End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Call cmdActualiza_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub
