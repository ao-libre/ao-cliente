VERSION 5.00
Begin VB.Form frmShowProcess 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmShowProcess"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCaptions 
      Height          =   5325
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   6255
   End
   Begin VB.ListBox lstProcess 
      Height          =   5325
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmShowProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
