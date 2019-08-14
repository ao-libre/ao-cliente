VERSION 5.00
Begin VB.Form FrmRetos 
   BackColor       =   &H80000008&
   Caption         =   "Panel Retos"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCompa 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtCompa 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtOponente 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtOponente 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtGld 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "0"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtOponente 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   9
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblSelected 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENVIAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Index           =   3
      Left            =   1320
      TabIndex        =   15
      Top             =   3960
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compañero 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oponente 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   1125
   End
   Begin VB.Label lblCompa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compañero 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1275
   End
   Begin VB.Label lblOponente2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oponente 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1125
   End
   Begin VB.Label lblOro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monedas de Oro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblOponente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oponente 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label lblSelected 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3vs3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   555
   End
   Begin VB.Label lblSelected 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2vs2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   555
   End
   Begin VB.Label lblSelected 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1vs1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "FrmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SelectedIndex As Byte

Private Sub Form_Load()
    SelectedIndex = 0
    
End Sub

Private Sub lblSelected_Click(Index As Integer)

    Dim ErrorMsg As String
    Dim ListUser As String
    
    Select Case Index
        Case 0 ' 1vs1
            txtOponente(1).Visible = False
            txtOponente(2).Visible = False
            txtCompa(0).Visible = False
            txtCompa(1).Visible = False
            SelectedIndex = 1
        Case 1 ' 2vs2
            txtOponente(1).Visible = True
            txtOponente(2).Visible = False
            txtCompa(0).Visible = True
            txtCompa(1).Visible = False
            SelectedIndex = 2
        Case 2 ' 3vs3
            txtOponente(1).Visible = True
            txtOponente(2).Visible = True
            txtCompa(0).Visible = True
            txtCompa(1).Visible = True
            SelectedIndex = 3
        Case 3 ' Enviar Reto
            If Not CheckDataReto(SelectedIndex, ListUser, ErrorMsg) Then
                MsgBox ErrorMsg
                Exit Sub
            End If
            
            Call Protocol.WriteFightSend(ListUser, Val(txtGld.Text))
    End Select
End Sub

Private Function CheckDataReto(ByVal Selected As Byte, _
                                ByRef ListUser As String, _
                                ByRef ErrorMsg As String) As Boolean
    CheckDataReto = False
    
    Dim A As Long
    
    If Val(txtGld.Text) < 0 Then
        ErrorMsg = "La apuesta mínima es por 0 monedas de oro"
        Exit Function
    End If
    
    If Len(txtOponente(0).Text) <= 0 Then
        ErrorMsg = "Debes seleccionar al oponente n°1"
        Exit Function
    End If
    
    ListUser = txtOponente(0).Text
    
    Select Case Selected
        Case 2
            If Len(txtOponente(1).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar al oponente n°2"
                Exit Function
            End If
            
            If Len(txtCompa(0).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar a tu compañero"
                Exit Function
            End If
            
            ListUser = txtOponente(0).Text & "-" & txtOponente(1).Text & "-" & txtCompa(0).Text
        Case 3
            If Len(txtOponente(1).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar al oponente n°2"
                Exit Function
            End If
            
            If Len(txtOponente(2).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar al oponente n°3"
                Exit Function
            End If
            
            If Len(txtCompa(0).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar a tu compañero n°2"
                Exit Function
            End If
            
            If Len(txtCompa(1).Text) <= 0 Then
                ErrorMsg = "Debes seleccionar a tu compañero n°3"
                Exit Function
            End If
            
            ListUser = txtOponente(0).Text & "-" & txtOponente(1) & "-" & txtOponente(2) & "-" & txtCompa(0).Text & "-" & txtCompa(1).Text
    End Select
    
    
    CheckDataReto = True
End Function
