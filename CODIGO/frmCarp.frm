VERSION 5.00
Begin VB.Form frmCarp 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   4
      Left            =   5430
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   14
      Top             =   3930
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   3
      Left            =   5430
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   3135
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   2
      Left            =   5430
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   12
      Top             =   2340
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   1
      Left            =   5430
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   11
      Top             =   1545
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtCantItems 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5175
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "Ingrese la cantidad total de items a construir."
      Top             =   2925
      Width           =   1050
   End
   Begin VB.ComboBox cboItemsCiclo 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   1
      Left            =   870
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   10
      Top             =   1545
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMaderas0 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1710
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   9
      Top             =   1545
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.VScrollBar Scroll 
      Height          =   3135
      Left            =   450
      TabIndex        =   0
      Top             =   1410
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picMaderas1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1710
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   8
      Top             =   2340
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   2
      Left            =   870
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   2340
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMaderas2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1710
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   6
      Top             =   3135
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   3
      Left            =   870
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   3135
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMaderas3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1710
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      Top             =   3930
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   4
      Left            =   870
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   3930
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCantidadCiclo 
      Height          =   645
      Left            =   5160
      Top             =   3435
      Width           =   1110
   End
   Begin VB.Image imgPestania 
      Height          =   255
      Index           =   1
      Left            =   1680
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image imgPestania 
      Height          =   255
      Index           =   0
      Left            =   720
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   975
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   4
      Left            =   5280
      Top             =   3780
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   3
      Left            =   5280
      Top             =   2985
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   2
      Left            =   5280
      Top             =   2190
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   1
      Left            =   5280
      Top             =   1395
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoMaderas 
      Height          =   780
      Index           =   4
      Left            =   1560
      Top             =   3780
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgMarcoMaderas 
      Height          =   780
      Index           =   3
      Left            =   1560
      Top             =   2985
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgMarcoMaderas 
      Height          =   780
      Index           =   2
      Left            =   1560
      Top             =   2190
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgMarcoMaderas 
      Height          =   780
      Index           =   1
      Left            =   1560
      Top             =   1395
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   4
      Left            =   720
      Top             =   3780
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   3
      Left            =   720
      Top             =   2985
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   2
      Left            =   720
      Top             =   2190
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   1
      Left            =   720
      Top             =   1395
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgCerrar 
      Height          =   360
      Left            =   2760
      Top             =   4650
      Width           =   1455
   End
   Begin VB.Image imgConstruir3 
      Height          =   420
      Left            =   3150
      Top             =   3960
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgConstruir2 
      Height          =   420
      Left            =   3150
      Top             =   3180
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgConstruir1 
      Height          =   420
      Left            =   3150
      Top             =   2370
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgConstruir0 
      Height          =   420
      Left            =   3150
      Top             =   1560
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgChkMacro 
      Height          =   420
      Left            =   5415
      MousePointer    =   99  'Custom
      Top             =   1860
      Width           =   435
   End
   Begin VB.Image imgMejorar0 
      Height          =   420
      Left            =   3150
      Top             =   1560
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgMejorar1 
      Height          =   420
      Left            =   3150
      Top             =   2370
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgMejorar2 
      Height          =   420
      Left            =   3150
      Top             =   3180
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgMejorar3 
      Height          =   420
      Left            =   3150
      Top             =   3960
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Dim Cargando As Boolean
Private clsFormulario As clsFormMovementManager

Private Enum ePestania
    ieItems
    ieMejorar
End Enum

Private picCheck As Picture
Private picRecuadroItem As Picture
Private picRecuadroMaderas As Picture

Private Pestanias(1) As Picture
Private UltimaPestania As Byte

Private cBotonCerrar As clsGraphicalButton
Private cBotonConstruir(0 To 4) As clsGraphicalButton
Private cBotonMejorar(0 To 4) As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private UsarMacro As Boolean

Private Sub Form_Load()
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Call LoadDefaultValues
    
    Me.Picture = LoadPicture(App.path & "\graficos\VentanaCarpinteriaItems.jpg")
    LoadButtons
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    Dim Index As Long
    
    GrhPath = DirGraficos

    Set Pestanias(ePestania.ieItems) = LoadPicture(GrhPath & "VentanaCarpinteriaItems.jpg")
    Set Pestanias(ePestania.ieMejorar) = LoadPicture(GrhPath & "VentanaCarpinteriaMejorar.jpg")
    
    Set picCheck = LoadPicture(GrhPath & "CheckBoxCarpinteria.jpg")
    
    Set picRecuadroItem = LoadPicture(GrhPath & "RecuadroItemsCarpinteria.jpg")
    Set picRecuadroMaderas = LoadPicture(GrhPath & "RecuadroMadera.jpg")
    
    For Index = 1 To MAX_LIST_ITEMS
        imgMarcoItem(Index).Picture = picRecuadroItem
        imgMarcoUpgrade(Index).Picture = picRecuadroItem
        imgMarcoMaderas(Index).Picture = picRecuadroMaderas
    Next Index
    
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonConstruir(0) = New clsGraphicalButton
    Set cBotonConstruir(1) = New clsGraphicalButton
    Set cBotonConstruir(2) = New clsGraphicalButton
    Set cBotonConstruir(3) = New clsGraphicalButton
    Set cBotonMejorar(0) = New clsGraphicalButton
    Set cBotonMejorar(1) = New clsGraphicalButton
    Set cBotonMejorar(2) = New clsGraphicalButton
    Set cBotonMejorar(3) = New clsGraphicalButton

    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarCarpinteria.jpg", _
                                    GrhPath & "BotonCerrarRolloverCarpinteria.jpg", _
                                    GrhPath & "BotonCerrarClickCarpinteria.jpg", Me)
                                    
    Call cBotonConstruir(0).Initialize(imgConstruir0, GrhPath & "BotonConstruirCarpinteria.jpg", _
                                    GrhPath & "BotonConstruirRolloverCarpinteria.jpg", _
                                    GrhPath & "BotonConstruirClickCarpinteria.jpg", Me)
                                    
    Call cBotonConstruir(1).Initialize(imgConstruir1, GrhPath & "BotonConstruirCarpinteria.jpg", _
                                    GrhPath & "BotonConstruirRolloverCarpinteria.jpg", _
                                    GrhPath & "BotonConstruirClickCarpinteria.jpg", Me)
                                    
    Call cBotonConstruir(2).Initialize(imgConstruir2, GrhPath & "BotonConstruirCarpinteria.jpg", _
                                    GrhPath & "BotonConstruirRolloverCarpinteria.jpg", _
                                    GrhPath & "BotonConstruirClickCarpinteria.jpg", Me)
                                    
    Call cBotonConstruir(3).Initialize(imgConstruir3, GrhPath & "BotonConstruirCarpinteria.jpg", _
                                    GrhPath & "BotonConstruirRolloverCarpinteria.jpg", _
                                    GrhPath & "BotonConstruirClickCarpinteria.jpg", Me)
    
    Call cBotonMejorar(0).Initialize(imgMejorar0, GrhPath & "BotonMejorarCarpinteria.jpg", _
                                    GrhPath & "BotonMejorarRolloverCarpinteria.jpg", _
                                    GrhPath & "BotonMejorarClickCarpinteria.jpg", Me)
    
    Call cBotonMejorar(1).Initialize(imgMejorar1, GrhPath & "BotonMejorarCarpinteria.jpg", _
                                    GrhPath & "BotonMejorarRolloverCarpinteria.jpg", _
                                    GrhPath & "BotonMejorarClickCarpinteria.jpg", Me)
    
    Call cBotonMejorar(2).Initialize(imgMejorar2, GrhPath & "BotonMejorarCarpinteria.jpg", _
                                    GrhPath & "BotonMejorarRolloverCarpinteria.jpg", _
                                    GrhPath & "BotonMejorarClickCarpinteria.jpg", Me)
    
    Call cBotonMejorar(3).Initialize(imgMejorar3, GrhPath & "BotonMejorarCarpinteria.jpg", _
                                    GrhPath & "BotonMejorarRolloverCarpinteria.jpg", _
                                    GrhPath & "BotonMejorarClickCarpinteria.jpg", Me)
                                    
    imgCantidadCiclo.Picture = LoadPicture(GrhPath & "ConstruirPorCiclo.jpg")
    
    imgChkMacro.Picture = picCheck
    
    imgPestania(ePestania.ieItems).MouseIcon = picMouseIcon
    imgPestania(ePestania.ieMejorar).MouseIcon = picMouseIcon
    
    imgChkMacro.MouseIcon = picMouseIcon
End Sub

Private Sub LoadDefaultValues()
    
    Dim MaxConstItem As Integer
    Dim i As Integer

    Cargando = True
    
    MaxConstItem = CInt((UserLvl - 2) * 0.2)
    MaxConstItem = IIf(MaxConstItem < 1, 1, MaxConstItem)
    MaxConstItem = IIf(UserClase = eClass.Worker, MaxConstItem, 1)
    
    For i = 1 To MaxConstItem
        cboItemsCiclo.AddItem i
    Next i
    
    cboItemsCiclo.ListIndex = 0
    
    Scroll.value = 0
    
    UsarMacro = True
    
    UltimaPestania = ePestania.ieItems
    
    Cargando = False
End Sub


Private Sub Construir(ByVal Index As Integer)

    Dim ItemIndex As Integer
    Dim CantItemsCiclo As Integer
    
    If Scroll.Visible = True Then ItemIndex = Scroll.value
    ItemIndex = ItemIndex + Index
    
    Select Case UltimaPestania
        Case ePestania.ieItems
        
            If UsarMacro Then
                CantItemsCiclo = Val(cboItemsCiclo.Text)
                MacroBltIndex = ObjCarpintero(ItemIndex).OBJIndex
                frmMain.ActivarMacroTrabajo
            Else
                ' Que cosntruya el maximo, total si sobra no importa, valida el server
                CantItemsCiclo = Val(cboItemsCiclo.List(cboItemsCiclo.ListCount - 1))
            End If
            
            Call WriteInitCrafting(Val(txtCantItems.Text), CantItemsCiclo)
            Call WriteCraftCarpenter(ObjCarpintero(ItemIndex).OBJIndex)
            
        Case ePestania.ieMejorar
            Call WriteItemUpgrade(CarpinteroMejorar(ItemIndex).OBJIndex)
    End Select
        
    Unload Me

End Sub

Public Sub HideExtraControls(ByVal NumItems As Integer, Optional ByVal Upgrading As Boolean = False)
    Dim i As Integer
    
    picMaderas0.Visible = (NumItems >= 1)
    picMaderas1.Visible = (NumItems >= 2)
    picMaderas2.Visible = (NumItems >= 3)
    picMaderas3.Visible = (NumItems >= 4)
    
    imgConstruir0.Visible = (NumItems >= 1 And Not Upgrading)
    imgConstruir1.Visible = (NumItems >= 2 And Not Upgrading)
    imgConstruir2.Visible = (NumItems >= 3 And Not Upgrading)
    imgConstruir3.Visible = (NumItems >= 4 And Not Upgrading)
    
    imgMejorar0.Visible = (NumItems >= 1 And Upgrading)
    imgMejorar1.Visible = (NumItems >= 2 And Upgrading)
    imgMejorar2.Visible = (NumItems >= 3 And Upgrading)
    imgMejorar3.Visible = (NumItems >= 4 And Upgrading)

    
    For i = 1 To MAX_LIST_ITEMS
        picItem(i).Visible = (NumItems >= i)
        imgMarcoItem(i).Visible = (NumItems >= i)
        imgMarcoMaderas(i).Visible = (NumItems >= i)

        ' Upgrade
        imgMarcoUpgrade(i).Visible = (NumItems >= i And Upgrading)
        picUpgrade(i).Visible = (NumItems >= i And Upgrading)
    Next i
    
    If NumItems > MAX_LIST_ITEMS Then
        Scroll.Visible = True
        Cargando = True
        Scroll.max = NumItems - MAX_LIST_ITEMS
        Cargando = False
    Else
        Scroll.Visible = False
    End If
    
    txtCantItems.Visible = Not Upgrading
    cboItemsCiclo.Visible = Not Upgrading And UsarMacro
    imgChkMacro.Visible = Not Upgrading
    imgCantidadCiclo.Visible = Not Upgrading And UsarMacro
End Sub

Private Sub RenderItem(ByRef Pic As PictureBox, ByVal GrhIndex As Long)
    Dim SR As RECT
    Dim DR As RECT
    
    With GrhData(GrhIndex)
        SR.Left = .sX
        SR.Top = .sY
        SR.Right = SR.Left + .pixelWidth
        SR.Bottom = SR.Top + .pixelHeight
    End With
    
    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.Bottom = 32
    
    Call DrawGrhtoHdc(Pic.hdc, GrhIndex, SR, DR)
    Pic.Refresh
End Sub

Public Sub RenderList(ByVal Inicio As Integer)
Dim i As Long
Dim NumItems As Integer

NumItems = UBound(ObjCarpintero)
Inicio = Inicio - 1

For i = 1 To MAX_LIST_ITEMS
    If i + Inicio <= NumItems Then
        With ObjCarpintero(i + Inicio)
            ' Agrego el item
            Call RenderItem(picItem(i), .GrhIndex)
            picItem(i).ToolTipText = .Name
        
            ' Inventario de leños
            Call InvMaderasCarpinteria(i).SetItem(1, 0, .Madera, 0, MADERA_GRH, 0, 0, 0, 0, 0, 0, "Leña")
            Call InvMaderasCarpinteria(i).SetItem(2, 0, .MaderaElfica, 0, MADERA_ELFICA_GRH, 0, 0, 0, 0, 0, 0, "Leña élfica")
        End With
    End If
Next i
End Sub

Public Sub RenderUpgradeList(ByVal Inicio As Integer)
Dim i As Long
Dim NumItems As Integer

NumItems = UBound(CarpinteroMejorar)
Inicio = Inicio - 1

For i = 1 To MAX_LIST_ITEMS
    If i + Inicio <= NumItems Then
        With CarpinteroMejorar(i + Inicio)
            ' Agrego el item
            Call RenderItem(picItem(i), .GrhIndex)
            picItem(i).ToolTipText = .Name
            
            Call RenderItem(picUpgrade(i), .UpgradeGrhIndex)
            picUpgrade(i).ToolTipText = .UpgradeName
        
            ' Inventario de leños
            Call InvMaderasCarpinteria(i).SetItem(1, 0, .Madera, 0, MADERA_GRH, 0, 0, 0, 0, 0, 0, "Leña")
            Call InvMaderasCarpinteria(i).SetItem(2, 0, .MaderaElfica, 0, MADERA_ELFICA_GRH, 0, 0, 0, 0, 0, 0, "Leña élfica")
        End With
    End If
Next i
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    For i = 1 To MAX_LIST_ITEMS
        Set InvMaderasCarpinteria(i) = Nothing
    Next i

    MirandoCarpinteria = False
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgChkMacro_Click()
    UsarMacro = Not UsarMacro
    
    If UsarMacro Then
        imgChkMacro.Picture = picCheck
    Else
        Set imgChkMacro.Picture = Nothing
    End If
    
    cboItemsCiclo.Visible = UsarMacro
    imgCantidadCiclo.Visible = UsarMacro
End Sub

Private Sub imgConstruir0_Click()
    Call Construir(1)
End Sub

Private Sub imgConstruir1_Click()
    Call Construir(2)
End Sub

Private Sub imgConstruir2_Click()
    Call Construir(3)
End Sub

Private Sub imgConstruir3_Click()
    Call Construir(4)
End Sub

Private Sub imgMejorar0_Click()
    Call Construir(1)
End Sub

Private Sub imgMejorar1_Click()
    Call Construir(2)
End Sub

Private Sub imgMejorar2_Click()
    Call Construir(3)
End Sub

Private Sub imgMejorar3_Click()
    Call Construir(4)
End Sub

Private Sub imgPestania_Click(Index As Integer)
    Dim i As Integer
    Dim NumItems As Integer
    
    If Cargando Then Exit Sub
    If UltimaPestania = Index Then Exit Sub
    
    Scroll.value = 0
    
    Select Case Index
        Case ePestania.ieItems
            ' Background
            Me.Picture = Pestanias(ePestania.ieItems)
            
            NumItems = UBound(ObjCarpintero)
        
            Call HideExtraControls(NumItems)
            
            ' Cargo inventarios e imagenes
            Call RenderList(1)
            

        Case ePestania.ieMejorar
            ' Background
            Me.Picture = Pestanias(ePestania.ieMejorar)
            
            NumItems = UBound(CarpinteroMejorar)
            
            Call HideExtraControls(NumItems, True)
            
            Call RenderUpgradeList(1)
    End Select

    UltimaPestania = Index

End Sub

Private Sub Scroll_Change()
    Dim i As Long
    
    If Cargando Then Exit Sub
    
    i = Scroll.value
    ' Cargo inventarios e imagenes
    
    Select Case UltimaPestania
        Case ePestania.ieItems
            Call RenderList(i + 1)
        Case ePestania.ieMejorar
            Call RenderUpgradeList(i + 1)
    End Select
End Sub
