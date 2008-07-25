VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgMapDungeon 
      Height          =   4935
      Left            =   0
      Top             =   0
      Width           =   8775
   End
   Begin VB.Image imgMap 
      Height          =   4935
      Left            =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp 'Cambiamos el "nivel" del mapa, al estilo Zelda ;D
            ToggleImgMaps
        Case Else
            Unload Me
    End Select
End Sub
Private Sub ToggleImgMaps()
    imgMap.Visible = Not imgMap.Visible
    imgMapDungeon.Visible = Not imgMapDungeon.Visible
End Sub

Private Sub Form_Load()
    On Error GoTo Error
    
    'Cargamos las imagenes de los mapas
    imgMap.Picture = LoadPicture(App.path & "\Graficos\mapa1.jpg")
    imgMapDungeon.Picture = LoadPicture(App.path & "\Graficos\mapa2.jpg")
    
    'Ajustamos el tamaño del formulario a la imagen más grande
    If imgMap.Width > imgMapDungeon.Width Then
        Me.Width = imgMap.Width
    Else
        Me.Width = imgMapDungeon.Width
    End If
    
    If imgMap.Height > imgMapDungeon.Height Then
        Me.Height = imgMap.Height
    Else
        Me.Height = imgMapDungeon.Height
    End If
    
    'Movemos ambas imágenes al centro del formulario
    imgMap.Left = Me.Width / 2 - imgMap.Width / 2
    imgMap.Top = Me.Height / 2 - imgMap.Height / 2
    
    imgMapDungeon.Left = Me.Width / 2 - imgMapDungeon.Width / 2
    imgMapDungeon.Top = Me.Height / 2 - imgMapDungeon.Height / 2
    
    imgMapDungeon.Visible = False
    Exit Sub
Error:
    MsgBox Err.Description, vbInformation, "Error: " & Err.number
    Unload Me
End Sub
