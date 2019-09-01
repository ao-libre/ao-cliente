VERSION 5.00
Begin VB.Form frmColor 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResultado 
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtG 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtR 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtA 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdConvertir 
      Caption         =   "Convertir"
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label lblB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   480
      Width           =   195
   End
   Begin VB.Label lblG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   210
   End
   Begin VB.Label lblR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1730
      TabIndex        =   3
      Top             =   480
      Width           =   195
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   195
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConvertir_Click()

    If Len(txtA.Text) <> 0 Then
        txtResultado.Text = D3DColorARGB(txtA.Text, txtR.Text, txtG.Text, txtB.Text)
    Else
        txtResultado.Text = D3DColorXRGB(txtR.Text, txtG.Text, txtB.Text)
    End If
    
End Sub
