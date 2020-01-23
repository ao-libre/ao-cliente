VERSION 5.00
Begin VB.Form frmBuscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscador de Objetos y NPC's"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Limpiarlista 
      Caption         =   "Limpiar lista"
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
      TabIndex        =   4
      Top             =   5640
      Width           =   4935
   End
   Begin VB.ListBox Resultados 
      Height          =   3765
      ItemData        =   "frmBuscar.frx":0000
      Left            =   120
      List            =   "frmBuscar.frx":0002
      TabIndex        =   3
      Top             =   1800
      Width           =   4935
   End
   Begin VB.CommandButton BuscarNPCs 
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton BuscarObjetos 
      Caption         =   "Buscar Objetos"
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
      Left            =   120
      MaskColor       =   &H8000000F&
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Busqueda 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Ingresa parte del nombre del objeto o NPC a buscar"
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Info 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<Clic derecho sobre el item para crear>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   4935
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
      Left            =   1320
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Menu mnuCrearO 
      Caption         =   "Crear Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuCrearObj 
         Caption         =   "Crear 1"
         Index           =   0
      End
      Begin VB.Menu mnuCrearObj 
         Caption         =   "Crear 10"
         Index           =   1
      End
      Begin VB.Menu mnuCrearObj 
         Caption         =   "Crear 100"
         Index           =   2
      End
      Begin VB.Menu mnuCrearObj 
         Caption         =   "Crear N"
         Index           =   3
      End
   End
   Begin VB.Menu mnuCrearN 
      Caption         =   "Crear NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuCrearNPC 
         Caption         =   "Crear NPC"
         Index           =   0
      End
      Begin VB.Menu mnuCrearNPC 
         Caption         =   "Crear con Respawn"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const LB_ITEMFROMPOINT = &H1A9

Private MensajeBusqueda As Boolean
Private BusquedaObjetos As Boolean

Public Sub AddItem(ByVal num As Integer, ByVal obj As Boolean, data As String)
    Resultados.AddItem data
    Resultados.ItemData(Resultados.ListCount - 1) = num
    
    If Not Info.Visible Or obj <> BusquedaObjetos Then
        If obj Then
            Info.Caption = JsonLanguage.item("FRMBUSCAR_INFO_OBJ").item("TEXTO")
        Else
            Info.Caption = JsonLanguage.item("FRMBUSCAR_INFO_NPC").item("TEXTO")
        End If
        Info.Visible = True
    End If
    
    BusquedaObjetos = obj
End Sub

Private Sub Form_Load()
   Call LoadTextsForm
   
   Info.Visible = False
   MensajeBusqueda = True
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub LoadTextsForm()
    lblTitle.Caption = JsonLanguage.item("FRMBUSCAR_TITLE").item("TEXTO")
    BuscarObjetos.Caption = JsonLanguage.item("FRMBUSCAR_BUSCAROBJETO").item("TEXTO")
    BuscarNPCs.Caption = JsonLanguage.item("FRMBUSCAR_BUSCARNPCS").item("TEXTO")
    Limpiarlista.Caption = JsonLanguage.item("FRMBUSCAR_LIMPIARLISTAS").item("TEXTO")
    Busqueda.Text = JsonLanguage.item("FRMBUSCAR_BUSQUEDA_TOOLTIP").item("TEXTO")
End Sub

Private Sub BuscarObjetos_Click()
    ' traduccion de 'objeto'
    Dim tObjeto As String
    tObjeto = JsonLanguage.item("OBJETO").item("TEXTO")
    
    ' Seamos un poco mas especificos y evitemos un overflow ^-^
    If Len(Busqueda.Text) < 2 Then
        MsgBox Replace$(JsonLanguage.item("ERROR_BUSCAR_MUY_CORTO").item("TEXTO"), "VAR_TARGET", tObjeto), vbApplicationModal
        Exit Sub
    End If
    
    'Limpiamos la lista antes.
    Resultados.Clear
    
    If Not MensajeBusqueda And Len(Busqueda.Text) <> 0 Then
        Call WriteSearchObj(Busqueda.Text)
    End If
End Sub

Private Sub BuscarNPCs_Click()
    ' Seamos un poco mas especificos y evitemos un overflow ^-^
    If Len(Busqueda.Text) < 2 Then
        MsgBox Replace$(JsonLanguage.item("ERROR_BUSCAR_MUY_CORTO").item("TEXTO"), "VAR_TARGET", "NPC"), vbApplicationModal
        Exit Sub
    End If
    
    'Limpiamos la lista antes.
    Resultados.Clear
    
    If Not MensajeBusqueda And Len(Busqueda.Text) <> 0 Then
        Call WriteSearchNpc(Busqueda.Text)
    End If

End Sub

Private Sub Busqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If MensajeBusqueda Then
        Busqueda = vbNullString
        Busqueda.ForeColor = vbBlack
        MensajeBusqueda = False
    End If
End Sub

Private Sub Limpiarlista_Click()
    Resultados.Clear
End Sub

Private Sub Resultados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Index As Long
    Dim PosX As Long, PosY As Long

    ' Detectamos el clic derecho para simular la seleccion
    If Button = vbRightButton Then
        ' Convertir a pixeles
        PosX = CLng(X / Screen.TwipsPerPixelX)
        PosY = CLng(Y / Screen.TwipsPerPixelY)

        ' Mensaje directo al hWnd usando WinAPI
        Index = SendMessage(Resultados.hWnd, LB_ITEMFROMPOINT, 0, ByVal ((PosY * 65536) + PosX))

        ' Si seleccionamos un item valido
        If Index < Resultados.ListCount Then
            Resultados.ListIndex = Index
        End If
    End If
End Sub

Private Sub Resultados_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Resultados.ListIndex >= 0 Then
        If BusquedaObjetos Then
            PopupMenu mnuCrearO
        Else
            PopupMenu mnuCrearN
        End If
    End If
End Sub

Private Sub mnuCrearObj_Click(Index As Integer)
On Error GoTo errhandler

    Dim Numero As Integer
    Dim cantidad As Integer
    
    Select Case Index
        Case 0
            cantidad = 1
        Case 1
            cantidad = 10
        Case 2
            cantidad = 100
        Case 3
            cantidad = Val(InputBox(JsonLanguage.item("FRMBUSCAR_INGRESE_CANTIDAD").item("TEXTO"), JsonLanguage.item("FRMBUSCAR_CANTIDAD").item("TEXTO"), vbApplicationModal))
            
            If cantidad <= 0 Then
                Exit Sub
            ElseIf cantidad > MAX_INVENTORY_OBJS Then
                cantidad = MAX_INVENTORY_OBJS
            End If
    End Select

    'Parche para evitar que no se seleccione un item y al querer crearlo explote el juego (Recox)
    If Resultados.ListIndex < 0 Then
        MsgBox (JsonLanguage.item("FRMBUSCAR_SELECCIONE_ITEM").item("TEXTO"))
        Exit Sub
    End If
    
    Numero = Resultados.ItemData(Resultados.ListIndex)
    
    If Numero > 0 Then
        Call WriteCreateItem(Resultados.ItemData(Resultados.ListIndex), cantidad)
    End If

    Exit Sub

errhandler:
    cantidad = MAX_INVENTORY_OBJS
    Resume Next
End Sub

Private Sub mnuCrearNPC_Click(Index As Integer)
    Dim Numero As Integer
    Dim Respawn As Boolean
    
    If Index = 1 Then Respawn = True
    
    'Parche para evitar que no se seleccione un item y al querer crearlo explote el juego (Recox)
    If Resultados.ListIndex < 0 Then
        MsgBox (JsonLanguage.item("FRMBUSCAR_SELECCIONE_ITEM").item("TEXTO"))
        Exit Sub
    End If
    
    Numero = Resultados.ItemData(Resultados.ListIndex)
    
    If Numero > 0 Then
        Call WriteCreateNPC(Numero, Respawn)
    End If
End Sub

'Parche: Al cerrar el formulario tambien te desconecta hahahaha ^_^'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub
