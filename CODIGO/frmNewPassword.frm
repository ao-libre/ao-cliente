VERSION 5.00
Begin VB.Form frmNewPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   495
      Left            =   293
      TabIndex        =   3
      Top             =   2400
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   293
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   293
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   293
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Confirmar contraseña nueva:"
      Height          =   255
      Left            =   293
      TabIndex        =   6
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña nueva:"
      Height          =   255
      Left            =   293
      TabIndex        =   5
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Contraseña anterior:"
      Height          =   255
      Left            =   293
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmNewPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Text2.Text <> Text3.Text Then
        Call MsgBox("Las contraseñas no coinciden", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Cambiar Contraseña")
        Exit Sub
    End If
    
    Call WriteChangePassword(Text1.Text, Text2.Text)
    Unload Me
End Sub
