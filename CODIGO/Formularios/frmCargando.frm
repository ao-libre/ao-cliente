VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   668
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Status 
      Height          =   2385
      Left            =   2610
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   3210
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   4207
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmCargando.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox LOGO 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   240
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   240
      Width           =   9600
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matias Fernando Pequeno
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
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

Public NoInternetConnection As Boolean
Private VersionNumberMaster As String
Private VersionNumberLocal As String

Private Sub Form_Load()
    Me.Analizar
    Me.Picture = LoadPicture(Game.path(Interfaces) & "VentanaCargando.jpg")

    'Solo hay 9 imagenes de cargando, cambiar 14 por el numero maximo si se quiere cambiar
    LOGO.Picture = LoadPicture(Game.path(Interfaces) & "ImagenCargando" & RandomNumber(1, 14) & ".jpg")
End Sub

Function Analizar()
On Error Resume Next
    Dim binaryFileToOpen As String
    Dim isLastVersion As Boolean

    isLastVersion = CheckIfRunningLastVersion
    
    If NoInternetConnection Then
        MsgBox "No hay conexion a internet, verificar que tengas internet/No Internet connection, please verify"
        Exit Function
    End If
           
    If Not isLastVersion = True Then
        If MsgBox("Tu version no es la actual, Deseas ejecutar el actualizador?. - Tu version: " & VersionNumberLocal & " Ultima version: " & VersionNumberMaster & " -- Your version is not up to date, open the launcher to update? ", vbYesNo) = vbYes Then
            binaryFileToOpen = GetVar(Game.path(INIT) & "Config.ini", "Launcher", "fileToOpen")
            Call ShellExecute(Me.hWnd, "open", App.path & binaryFileToOpen, "", "", 1)
            End
        End If
    End If
End Function

Private Function CheckIfRunningLastVersion() As Boolean
On Error Resume Next

    Dim responseGithub As String
    Dim JsonObject     As Object
    
    Set Inet = New clsInet
    
    responseGithub = Inet.OpenRequest("https://api.github.com/repos/ao-libre/ao-cliente/releases/latest", "GET")
    responseGithub = Inet.Execute
    responseGithub = Inet.GetResponseAsString
    
    Set JsonObject = JSON.parse(responseGithub)
    
    VersionNumberMaster = JsonObject.Item("tag_name")
    VersionNumberLocal = GetVar(Game.path(INIT) & "Config.ini", "Cliente", "VersionTagRelease")

    If VersionNumberMaster = VersionNumberLocal Then
        CheckIfRunningLastVersion = True
    Else
        CheckIfRunningLastVersion = False
    End If
End Function
