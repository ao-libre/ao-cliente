VERSION 5.00
Begin VB.Form frmCambiaMotd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar MOTD"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2580
      TabIndex        =   3
      Top             =   3300
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   660
      TabIndex        =   2
      Top             =   3300
      Width           =   1455
   End
   Begin VB.TextBox txtMotd 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   660
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No te olvides de poner los codigos de colores al final de cada linea!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "frmCambiaMotd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Dim T() As String
Dim I As Long, N As Long, Pos As Long

If Len(txtMotd.Text) >= 2 Then
    If Right(txtMotd.Text, 2) = vbCrLf Then txtMotd.Text = Left(txtMotd.Text, Len(txtMotd.Text) - 2)
End If

T = Split(txtMotd.Text, vbCrLf)

'hola~1~1~1~1~1

For I = LBound(T) To UBound(T)
    N = 0
    Pos = InStr(1, T(I), "~")
    Do While Pos > 0 And Pos < Len(T(I))
        N = N + 1
        Pos = InStr(Pos + 1, T(I), "~")
    Loop
    If N <> 5 Then
        MsgBox "Error en el formato de la linea " & I + 1 & "."
        Exit Sub
    End If
Next I

Call SendData("ZMOTD" & txtMotd.Text)
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me

End Sub
