VERSION 5.00
Begin VB.Form frmBuscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscador de Objetos y NPC's"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRespawn 
      Caption         =   "Respawn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6840
      TabIndex        =   12
      Text            =   "1"
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Limpiarlistas 
      Caption         =   "Limpiar Listas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   7335
   End
   Begin VB.ListBox ListCrearNpcs 
      Height          =   3375
      ItemData        =   "frmBuscar.frx":0000
      Left            =   960
      List            =   "frmBuscar.frx":0002
      TabIndex        =   8
      Top             =   2640
      Width           =   735
   End
   Begin VB.ListBox ListCrearObj 
      Height          =   3375
      ItemData        =   "frmBuscar.frx":0004
      Left            =   120
      List            =   "frmBuscar.frx":0006
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   3765
      ItemData        =   "frmBuscar.frx":0008
      Left            =   1800
      List            =   "frmBuscar.frx":000A
      TabIndex        =   5
      Top             =   2400
      Width           =   5655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Buscar NPCs"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar Objetos."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      MaskColor       =   &H8000000F&
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox NPCs 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   7335
   End
   Begin VB.TextBox Objetos 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6000
      TabIndex        =   13
      Top             =   1155
      Width           =   795
   End
   Begin VB.Label CrearObjetos 
      Alignment       =   2  'Center
      Caption         =   "Objetos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label CrearNPCs 
      Alignment       =   2  'Center
      Caption         =   "NPC's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Crear 
      Alignment       =   2  'Center
      Caption         =   "Crear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Buscador"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Menu mnuCrearO 
      Caption         =   "Crear Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuCrearObj 
         Caption         =   "Crear Objeto?"
      End
   End
   Begin VB.Menu mnuCrearN 
      Caption         =   "Crear NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuCrearNPC 
         Caption         =   "Crear NPC?"
      End
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblTitle.Caption = JsonLanguage.Item("FRMBUSCAR_TITLE").Item("TEXTO")
    Crear.Caption = JsonLanguage.Item("FRMBUSCAR_CREAR").Item("TEXTO")
    CrearNPCs.Caption = JsonLanguage.Item("FRMBUSCAR_CREARNPCS").Item("TEXTO")
    CrearObjetos.Caption = JsonLanguage.Item("FRMBUSCAR_CREAROBJETOS").Item("TEXTO")
    Label2.Caption = JsonLanguage.Item("FRMBUSCAR_CANTIDAD").Item("TEXTO")
    chkRespawn.Caption = JsonLanguage.Item("FRMBUSCAR_RESPAWN").Item("TEXTO")
    Command1.Caption = JsonLanguage.Item("FRMBUSCAR_BUSCAROBJETO").Item("TEXTO")
    Command2.Caption = JsonLanguage.Item("FRMBUSCAR_BUSCARNPCS").Item("TEXTO")
    Limpiarlistas.Caption = JsonLanguage.Item("FRMBUSCAR_LIMPIARLISTAS").Item("TEXTO")
    
    'Fix: Missing args. Select an option by default.
    ListCrearObj.SelCount = 0
    ListCrearNpcs.SelCount = 0
End Sub

Private Sub Command1_Click()
       
        ' traduccion de 'objeto'
        Dim tObjeto As String
            tObjeto = JsonLanguage.Item("OBJETO").Item("TEXTO")
        
        ' Seamos un poco mas especificos y evitemos un overflow ^-^
        If Len(Objetos.Text) < 4 Then
            MsgBox Replace$(JsonLanguage.Item("ERROR_BUSCAR_MUY_CORTO").Item("TEXTO"), "VAR_TARGET", tObjeto), vbApplicationModal
            Exit Sub
        End If
        
        'Limpiamos las listas antes.
        Call Limpiarlistas_Click
        
        If Len(Objetos.Text) <> 0 Then
            Call WriteSearchObj(Objetos.Text)
        End If

End Sub

Private Sub Command2_Click()
        
        ' Seamos un poco mas especificos y evitemos un overflow ^-^
        If Len(NPCs.Text) < 4 Then
            MsgBox Replace$(JsonLanguage.Item("ERROR_BUSCAR_MUY_CORTO").Item("TEXTO"), "VAR_TARGET", "NPC"), vbApplicationModal
            Exit Sub
        End If
        
        'Limpiamos las listas antes.
        Call Limpiarlistas_Click
        
        If Len(NPCs.Text) > 0 Then
            Call WriteSearchNpc(NPCs.Text)
        End If

End Sub

Private Sub Limpiarlistas_Click()

        ListCrearNpcs.Clear
        
        List1.Clear
        
        ListCrearObj.Clear
End Sub

Private Sub ListCrearNpcs_MouseDown(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

        If Button = vbLeftButton Then
            Call PopupMenu(mnuCrearN)
        End If

End Sub

Private Sub ListCrearObj_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)

        If Button = vbLeftButton Then
            Call PopupMenu(mnuCrearO)
        End If

End Sub

Private Sub mnuCrearObj_Click()

        If ListCrearObj.Visible And LenB(ListCrearObj.Text) > 0 And LenB(txtCantidad.Text) > 0 Then
            Call WriteCreateItem(ListCrearObj.Text, txtCantidad.Text)
        End If

End Sub

Private Sub mnuCrearNPC_Click()

        If ListCrearNpcs.Visible And LenB(ListCrearNpcs.Text) > 0 Then
            Call WriteCreateNPC(ListCrearNpcs.Text, CBool(chkRespawn.Value))
        End If

End Sub

'Parche: Al cerrar el formulario tambien te desconecta hahahaha ^_^'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub
