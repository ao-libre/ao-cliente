VERSION 5.00
Begin VB.Form frmMessageTxt 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensajes Predefinidos"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cancelCmd 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2400
      TabIndex        =   21
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton okCmd 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   9
      Left            =   1080
      TabIndex        =   10
      Top             =   3440
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   9
      Top             =   3080
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   7
      Left            =   1080
      TabIndex        =   8
      Top             =   2720
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   7
      Top             =   2360
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   6
      Top             =   2000
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   5
      Top             =   1640
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   4
      Top             =   1280
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   920
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   560
      Width           =   3400
   End
   Begin VB.TextBox messageTxt 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   200
      Width           =   3400
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mensaje 10:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   870
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mensaje 9:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mensaje 8:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   780
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mensaje 7:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mensaje 6:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mensaje 5:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mensaje 4:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mensaje 3:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mensaje 2:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Mensaje 1:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "frmMessageTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cancelCmd_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    For i = 0 To 9
        messageTxt(i) = CustomMessages.Message(i)
    Next i
End Sub

Private Sub okCmd_Click()
On Error GoTo ErrHandler
    Dim i As Long
    
    For i = 0 To 9
        CustomMessages.Message(i) = messageTxt(i)
    Next i
    
    Unload Me
Exit Sub

ErrHandler:
    'Did detected an invalid message??
    If Err.number = CustomMessages.InvalidMessageErrCode Then
        Call MsgBox("El Mensaje " & CStr(i + 1) & " es inválido. Modifiquelo por favor.")
    End If
End Sub
