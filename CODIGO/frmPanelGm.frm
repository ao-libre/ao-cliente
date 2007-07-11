VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPanelGm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel GM"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   4200
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   7695
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdNOREAL 
         Caption         =   "/NOREAL"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   86
         Top             =   6480
         Width           =   1815
      End
      Begin VB.CommandButton cmdNOCAOS 
         Caption         =   "/NOCAOS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   85
         Top             =   6480
         Width           =   1815
      End
      Begin VB.CommandButton cmdKICKCONSE 
         Caption         =   "/KICKCONSE"
         CausesValidation=   0   'False
         Height          =   675
         Left            =   2520
         TabIndex        =   84
         Top             =   6960
         Width           =   1335
      End
      Begin VB.CommandButton cmdACEPTCONSECAOS 
         Caption         =   "/ACEPTCONSECAOS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   83
         Top             =   7320
         Width           =   2295
      End
      Begin VB.CommandButton cmdACEPTCONSE 
         Caption         =   "/ACEPTCONSE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   82
         Top             =   6960
         Width           =   2295
      End
      Begin VB.ComboBox cboListaUsus 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   54
         Top             =   480
         Width           =   3675
      End
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "Actualizar"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   120
         Width           =   3675
      End
      Begin VB.CommandButton cmdIRCERCA 
         Caption         =   "/IRCERCA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   52
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdDONDE 
         Caption         =   "/DONDE"
         CausesValidation=   0   'False
         Height          =   675
         Left            =   120
         TabIndex        =   51
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdPENAS 
         Caption         =   "/PENAS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdTELEP 
         Caption         =   "/TELEP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   49
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdSILENCIAR 
         Caption         =   "/SILENCIAR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   48
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdIRA 
         Caption         =   "/IRA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   47
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdCARCEL 
         Caption         =   "/CARCEL"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   46
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdADVERTENCIA 
         Caption         =   "/ADVERTENCIA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   45
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdINFO 
         Caption         =   "/INFO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   44
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdSTAT 
         Caption         =   "/STAT"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   43
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdBAL 
         Caption         =   "/BAL"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   42
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdINV 
         Caption         =   "/INV"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdBOV 
         Caption         =   "/BOV"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   40
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdSKILLS 
         Caption         =   "/SKILLS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   39
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdREVIVIR 
         Caption         =   "/REVIVIR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdPERDON 
         Caption         =   "/PERDON"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   37
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdECHAR 
         Caption         =   "/ECHAR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdEJECUTAR 
         Caption         =   "/EJECUTAR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   35
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdBAN 
         Caption         =   "/BAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdUNBAN 
         Caption         =   "/UNBAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   33
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdSUM 
         Caption         =   "/SUM"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   32
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdNICK2IP 
         Caption         =   "/NICK2IP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdESTUPIDO 
         Caption         =   "/ESTUPIDO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton cmdNOESTUPIDO 
         Caption         =   "/NOESTUPIDO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   29
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton cmdBORRARPENA 
         Caption         =   "/BORRARPENA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   28
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdLASTIP 
         Caption         =   "/LASTIP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   27
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdCONDEN 
         Caption         =   "/CONDEN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdRAJAR 
         Caption         =   "/RAJAR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   25
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdRAJARCLAN 
         Caption         =   "/RAJARCLAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   24
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton cmdLASTEMAIL 
         Caption         =   "/LASTEMAIL"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   23
         Top             =   2880
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8160
      Width           =   4215
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdGMSG 
         Caption         =   "/GMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdHORA 
         Caption         =   "/HORA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdRMSG 
         Caption         =   "/RMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdREALMSG 
         Caption         =   "/REALMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdCAOSMSG 
         Caption         =   "/CAOSMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdCIUMSG 
         Caption         =   "/CIUMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdTALKAS 
         Caption         =   "/TALKAS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdMOTDCAMBIA 
         Caption         =   "/MOTDCAMBIA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSMSG 
         Caption         =   "/SMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2415
      Index           =   5
      Left            =   120
      TabIndex        =   55
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdCC 
         Caption         =   "/CC"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   72
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton cmdLIMPIAR 
         Caption         =   "/LIMPIAR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   71
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCT 
         Caption         =   "/CT"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   70
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdDT 
         Caption         =   "/DT"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   69
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdLLUVIA 
         Caption         =   "/LLUVIA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   68
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdMASSDEST 
         Caption         =   "/MASSDEST"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   67
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdPISO 
         Caption         =   "/PISO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   66
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdCI 
         Caption         =   "/CI"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   65
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdDEST 
         Caption         =   "/DEST"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   64
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2415
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdONLINEREAL 
         Caption         =   "/ONLINEREAL"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   21
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdONLINECAOS 
         Caption         =   "/ONLINECAOS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   20
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdNENE 
         Caption         =   "/NENE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdSHOW_SOS 
         Caption         =   "/SHOW SOS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   18
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdTRABAJANDO 
         Caption         =   "/TRABAJANDO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdOCULTANDO 
         Caption         =   "/OCULTANDO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdONLINEGM 
         Caption         =   "/ONLINEGM"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   15
         Top             =   2040
         Width           =   3015
      End
      Begin VB.CommandButton cmdONLINEMAP 
         Caption         =   "/ONLINEMAP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdBORRAR_SOS 
         Caption         =   "/BORRAR SOS"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   6
      Left            =   120
      TabIndex        =   56
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdSHOWCMSG 
         Caption         =   "/SHOWCMSG"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   80
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdBANCLAN 
         Caption         =   "/BANCLAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   79
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton cmdMIEMBROSCLAN 
         Caption         =   "/MIEMBROSCLAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   78
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdBANIPRELOAD 
         Caption         =   "/BANIPRELOAD"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   77
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdBANIPLIST 
         Caption         =   "/BANIPLIST"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   76
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdIP2NICK 
         Caption         =   "/IP2NICK"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   75
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdBANIP 
         Caption         =   "/BANIP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   74
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdUNBANIP 
         Caption         =   "/UNBANIP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   73
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdSHOWNAME 
         Caption         =   "/SHOWNAME"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   63
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdREM 
         Caption         =   "/REM"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   600
         TabIndex        =   62
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CommandButton cmdINVISIBLE 
         Caption         =   "/INVISIBLE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   600
         TabIndex        =   61
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdSETDESC 
         Caption         =   "/SETDESC"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   60
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdNAVE 
         Caption         =   "/NAVE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   600
         TabIndex        =   59
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdCHATCOLOR 
         Caption         =   "/CHATCOLOR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   600
         TabIndex        =   58
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdIGNORADO 
         Caption         =   "/IGNORADO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   57
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      CausesValidation=   0   'False
      Height          =   1935
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3413
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Message"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Player"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Me"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "World"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Admin"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmPanelGm.frm
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

''
' IMPORTANT!!!
' To prevent the combo list of usernames from closing when a conole message arrives, the Validate event allways
' sets the Cancel arg to True. This, combined with setting the CausesValidation of the RichTextBox to True
' makes the trick. However, in order to be able to use other commands, ALL OTHER controls in this form must have the
' CuasesValidation parameter set to false (unless you want to code your custom flag system to know when to allow or not the loose of focus).

Private Sub cboListaUsus_Validate(Cancel As Boolean)
    Cancel = True
End Sub

Private Sub cmdACEPTCONSE_Click()
    '/ACEPTCONSE
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        If MsgBox("¿Seguro desea aceptar a " & nick & " como consejero real?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteAcceptRoyalCouncilMember(nick)
End Sub

Private Sub cmdACEPTCONSECAOS_Click()
    '/ACEPTCONSECAOS
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        If MsgBox("¿Seguro desea aceptar a " & nick & " como consejero del caos?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteAcceptChaosCouncilMember(nick)
End Sub

Private Sub cmdADVERTENCIA_Click()
    '/ADVERTENCIA
    Dim tStr As String
    Dim nick As String

    nick = cboListaUsus.Text
        
    If LenB(nick) <> 0 Then
        tStr = InputBox("Escriba el motivo de la advertencia.", "Advertir a " & nick)
                
        If LenB(tStr) <> 0 Then
            'We use the Parser to control the command format
            Call ParseUserCommand("/ADVERTENCIA " & nick & "@" & tStr)
        End If
    End If
End Sub

Private Sub cmdBAL_Click()
    '/BAL
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteRequestCharGold(nick)
End Sub

Private Sub cmdBAN_Click()
    '/BAN
    Dim tStr As String
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then
        tStr = InputBox("Escriba el motivo del ban.", "BAN a " & nick)
                
        If LenB(tStr) <> 0 Then _
            If MsgBox("¿Seguro desea banear a " & nick & "?", vbYesNo, "Atencion!") = vbYes Then _
                Call WriteBanChar(nick, tStr)
    End If
End Sub

Private Sub cmdBANCLAN_Click()
    '/BANCLAN
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre del clan.", "Banear clan")
    If LenB(tStr) <> 0 Then _
        If MsgBox("¿Seguro desea banear al clan " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteGuildBan(tStr)
End Sub

Private Sub cmdBANIP_Click()
    '/BANIP
    Dim tStr As String
    Dim reason As String
    
    tStr = InputBox("Escriba el ip o el nick del PJ.", "Banear IP")
    
    reason = InputBox("Escriba el motivo del ban.", "Banear IP")
    
    If LenB(tStr) <> 0 Then _
        If MsgBox("¿Seguro desea banear la ip " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call ParseUserCommand("/BANIP " & tStr & " " & reason) 'We use the Parser to control the command format
End Sub

Private Sub cmdBANIPLIST_Click()
    '/BANIPLIST
    Call WriteBannedIPList
End Sub

Private Sub cmdBANIPRELOAD_Click()
    '/BANIPRELOAD
    Call WriteBannedIPReload
End Sub

Private Sub cmdBORRAR_SOS_Click()
    '/BORRAR SOS
    If MsgBox("¿Seguro desea borrar el SOS?", vbYesNo, "Atencion!") = vbYes Then _
        Call WriteCleanSOS
End Sub

Private Sub cmdBORRARPENA_Click()
    '/BORRARPENA
    Dim tStr As String
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then
        tStr = InputBox("Indique el número de la pena a borrar.", "Borrar pena")
        If LenB(tStr) <> 0 Then _
            If MsgBox("¿Seguro desea borrar la pena " & tStr & " a " & nick & "?", vbYesNo, "Atencion!") = vbYes Then _
                Call ParseUserCommand("/BORRARPENA " & nick & "@" & tStr) 'We use the Parser to control the command format
    End If
End Sub

Private Sub cmdBOV_Click()
    '/BOV
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteRequestCharBank(nick)
End Sub

Private Sub cmdCAOSMSG_Click()
    '/CAOSMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba un Mensaje.", "Mensaje por consola LegionOscura")
    If LenB(tStr) <> 0 Then _
        Call WriteChaosLegionMessage(tStr)
End Sub

Private Sub cmdCARCEL_Click()
    '/CARCEL
    Dim tStr As String
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then
        tStr = InputBox("Escriba el motivo de la pena.", "Carcel a " & nick)
                
        If LenB(tStr) <> 0 Then
            tStr = tStr & "@" & InputBox("Indique el tiempo de condena (entre 0 y 60 minutos).", "Carcel a " & nick)
            'We use the Parser to control the command format
            Call ParseUserCommand("/CARCEL " & nick & "@" & tStr)
        End If
    End If
End Sub

Private Sub cmdCC_Click()
    '/CC
    Call WriteSpawnListRequest
End Sub

Private Sub cmdCHATCOLOR_Click()
    '/CHATCOLOR
    Dim tStr As String
    
    tStr = InputBox("Defina el color (R G B). Deje en blanco para usar el default.", "Cambiar color del chat")
    
    Call ParseUserCommand("/CHATCOLOR " & tStr) 'We use the Parser to control the command format
End Sub

Private Sub cmdCI_Click()
    '/CI
    Dim tStr As String
    
    tStr = InputBox("Indique el número del objeto a crear.", "Crear Objeto")
    If LenB(tStr) <> 0 Then _
        If MsgBox("¿Seguro desea crear el objeto " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call ParseUserCommand("/CI " & tStr) 'We use the Parser to control the command format
End Sub

Private Sub cmdCIUMSG_Click()
    '/CIUMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba un Mensaje.", "Mensaje por consola Ciudadanos")
    If LenB(tStr) <> 0 Then _
        Call WriteCitizenMessage(tStr)
End Sub

Private Sub cmdCONDEN_Click()
    '/CONDEN
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        If MsgBox("¿Seguro desea volver criminal a " & nick & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteTurnCriminal(nick)
End Sub

Private Sub cmdCT_Click()
    '/CT
    Dim tStr As String
    
    tStr = InputBox("Indique la posición donde lleva el portal (MAPA X Y).", "Crear Portal")
    If LenB(tStr) <> 0 Then _
        Call ParseUserCommand("/CT " & tStr) 'We use the Parser to control the command format
End Sub

Private Sub cmdDEST_Click()
    '/DEST
    If MsgBox("¿Seguro desea destruir el objeto sobre el que esta parado?", vbYesNo, "Atencion!") = vbYes Then _
        Call WriteDestroyItems
End Sub

Private Sub cmdDONDE_Click()
    '/DONDE
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteWhere(nick)
End Sub

Private Sub cmdDT_Click()
    'DT
    If MsgBox("¿Seguro desea destruir el portal?", vbYesNo, "Atencion!") = vbYes Then _
        Call WriteTeleportDestroy
End Sub

Private Sub cmdECHAR_Click()
    '/ECHAR
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteKick(nick)
End Sub

Private Sub cmdEJECUTAR_Click()
    '/EJECUTAR
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        If MsgBox("¿Seguro desea ejecutar a " & nick & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteExecute(nick)
End Sub

Private Sub cmdESTUPIDO_Click()
    '/ESTUPIDO
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteMakeDumb(nick)
End Sub

Private Sub cmdGMSG_Click()
    '/GMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de GM")
    If LenB(tStr) <> 0 Then _
        Call WriteGMMessage(tStr)
End Sub

Private Sub cmdHORA_Click()
    '/HORA
    Call Protocol.WriteServerTime
End Sub

Private Sub cmdIGNORADO_Click()
    '/IGNORADO
    Call WriteIgnored
End Sub

Private Sub cmdINFO_Click()
    '/INFO
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteRequestCharInfo(nick)
End Sub

Private Sub cmdINV_Click()
    '/INV
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteRequestCharInventory(nick)
End Sub

Private Sub cmdINVISIBLE_Click()
    '/INVISIBLE
    Call WriteInvisible
End Sub

Private Sub cmdIP2NICK_Click()
    '/IP2NICK
    Dim tStr As String
    
    tStr = InputBox("Escriba la ip.", "IP to Nick")
    If LenB(tStr) <> 0 Then _
        Call ParseUserCommand("/IP2NICK " & tStr) 'We use the Parser to control the command format
End Sub

Private Sub cmdIRA_Click()
    '/IRA
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteGoToChar(nick)
End Sub

Private Sub cmdIRCERCA_Click()
    '/IRCERCA
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteGoNearby(nick)
End Sub

Private Sub cmdKICKCONSE_Click()
    'KICKCONSE
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        If MsgBox("¿Seguro desea destituir a " & nick & " de su cargo de consejero?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteCouncilKick(nick)
End Sub

Private Sub cmdLASTEMAIL_Click()
    '/LASTEMAIL
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteRequestCharMail(nick)
End Sub

Private Sub cmdLASTIP_Click()
    '/LASTIP
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteLastIP(nick)
End Sub

Private Sub cmdLIMPIAR_Click()
    '/LIMPIAR
    Call WriteCleanWorld
End Sub

Private Sub cmdLLUVIA_Click()
    '/LLUVIA
    Call WriteRainToggle
End Sub

Private Sub cmdMASSDEST_Click()
    '/MASSDEST
    If MsgBox("¿Seguro desea destruir todos los items del mapa?", vbYesNo, "Atencion!") = vbYes Then _
        Call WriteDestroyAllItemsInArea
End Sub

Private Sub cmdMIEMBROSCLAN_Click()
    '/MIEMBROSCLAN
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre del clan.", "Lista de miembros del clan")
    If LenB(tStr) <> 0 Then _
        Call WriteGuildMemberList(tStr)
End Sub

Private Sub cmdMOTDCAMBIA_Click()
    '/MOTDCAMBIA
    Call WriteChangeMOTD
End Sub

Private Sub cmdNAVE_Click()
    '/NAVE
    Call WriteNavigateToggle
End Sub

Private Sub cmdNENE_Click()
    '/NENE
    Dim tStr As String
    
    tStr = InputBox("Indique el mapa.", "Número de NPCs enemigos.")
    If LenB(tStr) <> 0 Then _
        Call ParseUserCommand("/NENE " & tStr) 'We use the Parser to control the command format
End Sub

Private Sub cmdNICK2IP_Click()
    '/NICK2IP
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteNickToIP(nick)
End Sub

Private Sub cmdNOCAOS_Click()
    '/NOCAOS
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        If MsgBox("¿Seguro desea expulsar a " & nick & " de la legión oscura?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteChaosLegionKick(nick)
End Sub

Private Sub cmdNOESTUPIDO_Click()
    '/NOESTUPIDO
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteMakeDumbNoMore(nick)
End Sub

Private Sub cmdNOREAL_Click()
    '/NOREAL
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        If MsgBox("¿Seguro desea expulsar a " & nick & " de la armada real?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteRoyalArmyKick(nick)
End Sub

Private Sub cmdOCULTANDO_Click()
    '/OCULTANDO
    Call WriteHiding
End Sub

Private Sub cmdONLINECAOS_Click()
    '/ONLINECAOS
    Call WriteOnlineChaosLegion
End Sub

Private Sub cmdONLINEGM_Click()
    '/ONLINEGM
    Call WriteOnlineGM
End Sub

Private Sub cmdONLINEMAP_Click()
    '/ONLINEMAP
    Call WriteOnlineMap
End Sub

Private Sub cmdONLINEREAL_Click()
    '/ONLINEREAL
    Call WriteOnlineRoyalArmy
End Sub

Private Sub cmdPENAS_Click()
    '/PENAS
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WritePunishments(nick)
End Sub

Private Sub cmdPERDON_Click()
    '/PERDON
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteForgive(nick)
End Sub

Private Sub cmdPISO_Click()
    '/PISO
    Call WriteItemsInTheFloor
End Sub

Private Sub cmdRAJAR_Click()
    '/RAJAR
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        If MsgBox("¿Seguro desea resetear la facción de " & nick & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteResetFactions(nick)
End Sub

Private Sub cmdRAJARCLAN_Click()
    '/RAJARCLAN
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        If MsgBox("¿Seguro desea expulsar a " & nick & " de su clan?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteRemoveCharFromGuild(nick)
End Sub

Private Sub cmdREALMSG_Click()
    '/REALMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba un Mensaje.", "Mensaje por consola ArmadaReal")
    If LenB(tStr) <> 0 Then _
        Call WriteRoyalArmyMessage(tStr)
End Sub

Private Sub cmdREM_Click()
    '/REM
    Dim tStr As String
    
    tStr = InputBox("Escriba el comentario.", "Comentario en el logGM")
    If LenB(tStr) <> 0 Then _
        Call WriteComment(tStr)
End Sub

Private Sub cmdREVIVIR_Click()
    '/REVIVIR
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteReviveChar(nick)
End Sub

Private Sub cmdRMSG_Click()
    '/RMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de RoleMaster")
    If LenB(tStr) <> 0 Then _
        Call WriteServerMessage(tStr)
End Sub

Private Sub cmdSETDESC_Click()
    '/SETDESC
    Dim tStr As String
    
    tStr = InputBox("Escriba una DESC.", "Set Description")
    If LenB(tStr) <> 0 Then _
        Call WriteSetCharDescription(tStr)
End Sub

Private Sub cmdSHOW_SOS_Click()
    '/SHOW SOS
    Call WriteSOSShowList
End Sub

Private Sub cmdSHOWCMSG_Click()
    '/SHOWCMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre del clan que desea escuchar.", "Escuchar los mensajes del clan")
    If LenB(tStr) <> 0 Then _
        Call WriteShowGuildMessages(tStr)
End Sub

Private Sub cmdSHOWNAME_Click()
    '/SHOWNAME
    Call WriteShowName
End Sub

Private Sub cmdSILENCIAR_Click()
    '/SILENCIAR
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteSilence(nick)
End Sub

Private Sub cmdSKILLS_Click()
    '/SKILLS
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteRequestCharSkills(nick)
End Sub

Private Sub cmdSMSG_Click()
    '/SMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba el mensaje.", "Mensaje de sistema")
    If LenB(tStr) <> 0 Then _
        Call WriteSystemMessage(tStr)
End Sub

Private Sub cmdSTAT_Click()
    '/STAT
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteRequestCharStats(nick)
End Sub

Private Sub cmdSUM_Click()
    '/SUM
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        Call WriteSummonChar(nick)
End Sub

Private Sub cmdTALKAS_Click()
    '/TALKAS
    Dim tStr As String
    
    tStr = InputBox("Escriba un Mensaje.", "Hablar por NPC")
    If LenB(tStr) <> 0 Then _
        Call WriteTalkAsNPC(tStr)
End Sub

Private Sub cmdTELEP_Click()
    '/TELEP
    Dim tStr As String
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then
        tStr = InputBox("Indique la posición (MAPA X Y).", "Transportar a " & nick)
        If LenB(tStr) <> 0 Then _
            Call ParseUserCommand("/TELEP " & nick & " " & tStr) 'We use the Parser to control the command format
    End If
End Sub

Private Sub cmdTRABAJANDO_Click()
    '/TRABAJANDO
    Call WriteWorking
End Sub

Private Sub cmdUNBAN_Click()
    '/UNBAN
    Dim nick As String

    nick = cboListaUsus.Text
    
    If LenB(nick) <> 0 Then _
        If MsgBox("¿Seguro desea unbanear a " & nick & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call WriteUnbanChar(nick)
End Sub

Private Sub cmdUNBANIP_Click()
    '/UNBANIP
    Dim tStr As String
    
    tStr = InputBox("Escriba el ip.", "Unbanear IP")
    If LenB(tStr) <> 0 Then _
        If MsgBox("¿Seguro desea unbanear la ip " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
            Call ParseUserCommand("/UNBANIP " & tStr) 'We use the Parser to control the command format
End Sub

Private Sub Form_Load()
    Call showTab(1)
    Call cmdActualiza_Click
End Sub

Private Sub cmdActualiza_Click()
    Call WriteRequestUserList
    Call FlushBuffer
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub TabStrip_Click()
    Call showTab(TabStrip.SelectedItem.index)
End Sub

Private Sub showTab(TabId As Byte)
    Dim i As Byte
    
    For i = 1 To Frame.UBound
        Frame(i).Visible = (i = TabId)
    Next i
    
    With Frame(TabId)
        frmPanelGm.Height = .Height + 1280
        TabStrip.Height = .Height + 480
        cmdCerrar.Top = .Height + 465
    End With
End Sub
