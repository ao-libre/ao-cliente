VERSION 5.00
Begin VB.Form frmAmbientEditor 
   Caption         =   "Editor de Ambiente"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Luces"
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   3135
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   840
         Max             =   10
         Min             =   1
         TabIndex        =   23
         Top             =   720
         Value           =   1
         Width           =   2055
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Borrar Luz Actual"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2340
         TabIndex        =   19
         Text            =   "255"
         Top             =   375
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1455
         TabIndex        =   18
         Text            =   "255"
         Top             =   375
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   585
         TabIndex        =   17
         Text            =   "255"
         Top             =   375
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Crear Luz en Posicion Actual"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Rango:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "R:           G:           B:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Meteo"
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   3135
      Begin VB.CheckBox Check3 
         Caption         =   "Llueve"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Nieve"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Aplicar"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   2895
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         Max             =   150
         Min             =   -30
         TabIndex        =   12
         Top             =   960
         Value           =   30
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Usar Niebla en el Mapa"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Grado de Niebla"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Luz Ambiental"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton Command7 
         Caption         =   "Aplicar"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   8
         Text            =   "255"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Text            =   "255"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Text            =   "255"
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Usar Luz propia"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Usar Luz del Dia"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "R:           G:           B:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar Ambiente"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6720
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recargar Ambiente"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Width           =   3135
   End
End
Attribute VB_Name = "frmAmbientEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    
    If Check1.Value = Checked Then
        HScroll1.Enabled = True
    Else
        HScroll1.Enabled = False
    End If
    
End Sub

Private Sub Check2_Click()
    
    If Check2.Value = Checked Then
        CurMapAmbient.Snow = True
    Else
        CurMapAmbient.Snow = False
    End If

End Sub

Private Sub Check3_Click()
    
    If Check3.Value = Checked Then
        CurMapAmbient.Rain = True
    Else
        CurMapAmbient.Rain = False
    End If

End Sub

Private Sub Command1_Click()
    Call Init_Ambient(UserMap)
End Sub

Private Sub Command10_Click()
    
    With CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Light
        .b = 0
        .g = 0
        .r = 0
        .range = 0
    End With
    
    Call Delete_Light_To_Map(UserPos.X, UserPos.Y)
    Call LightRenderAll
    
End Sub

Private Sub Command2_Click()
    Save_Ambient UserMap
    DoEvents
    
    Init_Ambient UserMap
End Sub

Private Sub Command7_Click()
    
    With CurMapAmbient
    
        If Option1(0).Value = True Then
            
            .UseDayAmbient = True
        
            With .OwnAmbientLight
        
                .a = 255
                .r = 0
                .g = 0
                .b = 0
        
            End With
        
        Else
            
            .UseDayAmbient = False
            
            With .OwnAmbientLight
            
                .a = 255
                .r = Val(Text1(0).Text)
                .g = Val(Text1(1).Text)
                .b = Val(Text1(2).Text)
            
            End With
            
        End If
    
    End With
    
    DoEvents
    
    Call Apply_OwnAmbient
End Sub

Private Sub Command8_Click()
    
    With CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Light
        .b = Val(Text4.Text)
        .g = Val(Text3.Text)
        .r = Val(Text2.Text)
        .range = Val(HScroll2.Value)
    End With
    
    Call Create_Light_To_Map(UserPos.X, UserPos.Y, Val(HScroll2.Value), Val(Text2.Text), Val(Text3.Text), Val(Text4.Text))
    
End Sub

Private Sub Command9_Click()
    
    If Check1.Value = Unchecked Then
        CurMapAmbient.Fog = -1
    Else
        CurMapAmbient.Fog = Val(HScroll1.Value)
    End If
    
End Sub
