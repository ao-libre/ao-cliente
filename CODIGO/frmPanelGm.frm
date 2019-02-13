VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPanelGm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel GM"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   4215
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2415
      Index           =   5
      Left            =   120
      TabIndex        =   55
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdDE 
         Caption         =   "/DE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2280
         TabIndex        =   110
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdCC 
         Caption         =   "/CC"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   720
         TabIndex        =   72
         Top             =   720
         Width           =   1095
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
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   7695
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdConsulta 
         Caption         =   "/CONSULTA"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   87
         Top             =   1320
         Width           =   1095
      End
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
         Left            =   120
         TabIndex        =   52
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdDONDE 
         Caption         =   "/DONDE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   51
         Top             =   1320
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
         Left            =   1320
         TabIndex        =   49
         Top             =   960
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
         Top             =   1320
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
         Left            =   2520
         TabIndex        =   32
         Top             =   960
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
         Caption         =   "/MODIFICARPENA"
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
      Height          =   7395
      Index           =   7
      Left            =   120
      TabIndex        =   88
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ACTUALIZAR"
         Height          =   495
         Left            =   2160
         TabIndex        =   109
         Top             =   2100
         Width           =   1695
      End
      Begin VB.TextBox txtNuevaDescrip 
         Height          =   765
         Left            =   120
         MaxLength       =   40
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   107
         Top             =   6120
         Width           =   3735
      End
      Begin VB.CommandButton cmdAddFollow 
         Caption         =   "Agregar Seguimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   105
         Top             =   6960
         Width           =   3735
      End
      Begin VB.TextBox txtNuevoUsuario 
         Height          =   285
         Left            =   120
         TabIndex        =   104
         Top             =   5580
         Width           =   3735
      End
      Begin VB.CommandButton cmdAddObs 
         Caption         =   "Agregar Observación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   102
         Top             =   4800
         Width           =   3735
      End
      Begin VB.TextBox txtObs 
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   101
         Top             =   3780
         Width           =   3735
      End
      Begin VB.TextBox txtDescrip 
         Height          =   675
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   40
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   99
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtCreador 
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   1620
         Width           =   1695
      End
      Begin VB.TextBox txtTimeOn 
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtIP 
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   540
         Width           =   1695
      End
      Begin VB.ListBox lstUsers 
         Height          =   2400
         Left            =   120
         TabIndex        =   89
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   108
         Top             =   60
         Width           =   660
      End
      Begin VB.Label Label9 
         Caption         =   "Descripción"
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
         TabIndex        =   106
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4200
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label Label8 
         Caption         =   "Usuario"
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
         TabIndex        =   103
         Top             =   5340
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   100
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   98
         Top             =   2700
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Creador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   96
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Logueado Hace:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   94
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "IP:"
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
         Left            =   2160
         TabIndex        =   92
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Online"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   2880
         TabIndex        =   91
         Top             =   60
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Usuarios Marcados"
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
         TabIndex        =   90
         Top             =   0
         Width           =   1695
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
         NumTabs         =   7
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
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seguimientos"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSeguimientos 
      Caption         =   "Seguimientos"
      Begin VB.Menu mnuIra 
         Caption         =   "Ir Cerca"
      End
      Begin VB.Menu mnuSum 
         Caption         =   "Sumonear"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Eliminar Seguimiento"
      End
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
    
    On Error GoTo cboListaUsus_Validate_Err
    
    Cancel = True

    
    Exit Sub

cboListaUsus_Validate_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cboListaUsus_Validate"
    End If
Resume Next
    
End Sub

Private Sub cmdACEPTCONSE_Click()
    '/ACEPTCONSE
    
    On Error GoTo cmdACEPTCONSE_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("¿Seguro desea aceptar a " & Nick & " como consejero real?", vbYesNo, "Atencion!") = vbYes Then Call WriteAcceptRoyalCouncilMember(Nick)

    
    Exit Sub

cmdACEPTCONSE_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdACEPTCONSE_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdACEPTCONSECAOS_Click()
    '/ACEPTCONSECAOS
    
    On Error GoTo cmdACEPTCONSECAOS_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("¿Seguro desea aceptar a " & Nick & " como consejero del caos?", vbYesNo, "Atencion!") = vbYes Then Call WriteAcceptChaosCouncilMember(Nick)

    
    Exit Sub

cmdACEPTCONSECAOS_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdACEPTCONSECAOS_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdAddFollow_Click()
    
    On Error GoTo cmdAddFollow_Click_Err
    
    Dim i As Long

    For i = 0 To lstUsers.ListCount

        If UCase$(lstUsers.List(i)) = UCase$(txtNuevoUsuario.Text) Then
            Call MsgBox("¡El usuario ya está en la lista!", vbOKOnly + vbExclamation)
            Exit Sub

        End If

    Next i
            
    If LenB(txtNuevoUsuario.Text) = 0 Then
        Call MsgBox("¡Escribe el nombre de un usuario!", vbOKOnly + vbExclamation)
        Exit Sub

    End If
    
    If LenB(txtNuevaDescrip.Text) = 0 Then
        Call MsgBox("¡Escribe el motivo del seguimiento!", vbOKOnly + vbExclamation)
        Exit Sub

    End If
    
    Call WriteRecordAdd(txtNuevoUsuario.Text, txtNuevaDescrip.Text)
    
    txtNuevoUsuario.Text = vbNullString
    txtNuevaDescrip.Text = vbNullString

    
    Exit Sub

cmdAddFollow_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdAddFollow_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdAddObs_Click()
    
    On Error GoTo cmdAddObs_Click_Err
    
    Dim Obs As String
    
    Obs = InputBox("Ingrese la observación", "Nueva Observación")
    
    If LenB(Obs) = 0 Then
        Call MsgBox("¡Escribe una observación!", vbOKOnly + vbExclamation)
        Exit Sub

    End If
    
    If lstUsers.ListIndex = -1 Then
        Call MsgBox("¡Seleccione un seguimiento!", vbOKOnly + vbExclamation)
        Exit Sub

    End If
    
    Call WriteRecordAddObs(lstUsers.ListIndex + 1, Obs)

    
    Exit Sub

cmdAddObs_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdAddObs_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdADVERTENCIA_Click()
    '/ADVERTENCIA
    
    On Error GoTo cmdADVERTENCIA_Click_Err
    
    Dim tStr As String
    Dim Nick As String

    Nick = cboListaUsus.Text
        
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo de la advertencia.", "Advertir a " & Nick)
                
        If LenB(tStr) <> 0 Then
            'We use the Parser to control the command format
            Call ParseUserCommand("/ADVERTENCIA " & Nick & "@" & tStr)

        End If

    End If

    
    Exit Sub

cmdADVERTENCIA_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdADVERTENCIA_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdBAL_Click()
    '/BAL
    
    On Error GoTo cmdBAL_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharGold(Nick)

    
    Exit Sub

cmdBAL_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdBAL_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdBAN_Click()
    '/BAN
    
    On Error GoTo cmdBAN_Click_Err
    
    Dim tStr As String
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo del ban.", "BAN a " & Nick)
                
        If LenB(tStr) <> 0 Then If MsgBox("¿Seguro desea banear a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteBanChar(Nick, tStr)

    End If

    
    Exit Sub

cmdBAN_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdBAN_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdBANCLAN_Click()
    '/BANCLAN
    
    On Error GoTo cmdBANCLAN_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre del clan.", "Banear clan")

    If LenB(tStr) <> 0 Then If MsgBox("¿Seguro desea banear al clan " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteGuildBan(tStr)

    
    Exit Sub

cmdBANCLAN_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdBANCLAN_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdBANIP_Click()
    '/BANIP
    
    On Error GoTo cmdBANIP_Click_Err
    
    Dim tStr   As String
    Dim Reason As String
    
    tStr = InputBox("Escriba el ip o el nick del PJ.", "Banear IP")
    
    Reason = InputBox("Escriba el motivo del ban.", "Banear IP")
    
    If LenB(tStr) <> 0 Then If MsgBox("¿Seguro desea banear la ip " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/BANIP " & tStr & " " & Reason) 'We use the Parser to control the command format

    
    Exit Sub

cmdBANIP_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdBANIP_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdBANIPLIST_Click()
    '/BANIPLIST
    
    On Error GoTo cmdBANIPLIST_Click_Err
    
    Call WriteBannedIPList

    
    Exit Sub

cmdBANIPLIST_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdBANIPLIST_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdBANIPRELOAD_Click()
    '/BANIPRELOAD
    
    On Error GoTo cmdBANIPRELOAD_Click_Err
    
    Call WriteBannedIPReload

    
    Exit Sub

cmdBANIPRELOAD_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdBANIPRELOAD_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdBORRAR_SOS_Click()
    
    On Error GoTo cmdBORRAR_SOS_Click_Err
    

    '/BORRAR SOS
    If MsgBox("¿Seguro desea borrar el SOS?", vbYesNo, "Atencion!") = vbYes Then Call WriteCleanSOS

    
    Exit Sub

cmdBORRAR_SOS_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdBORRAR_SOS_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdBORRARPENA_Click()
    '/BORRARPENA
    
    On Error GoTo cmdBORRARPENA_Click_Err
    
    Dim tStr As String
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Indique el número de la pena a borrar.", "Borrar pena")

        If LenB(tStr) <> 0 Then If MsgBox("¿Seguro desea borrar la pena " & tStr & " a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/BORRARPENA " & Nick & "@" & tStr) 'We use the Parser to control the command format

    End If

    
    Exit Sub

cmdBORRARPENA_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdBORRARPENA_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdBOV_Click()
    '/BOV
    
    On Error GoTo cmdBOV_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharBank(Nick)

    
    Exit Sub

cmdBOV_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdBOV_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdCAOSMSG_Click()
    '/CAOSMSG
    
    On Error GoTo cmdCAOSMSG_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba un Mensaje.", "Mensaje por consola LegionOscura")

    If LenB(tStr) <> 0 Then Call WriteChaosLegionMessage(tStr)

    
    Exit Sub

cmdCAOSMSG_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdCAOSMSG_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdCARCEL_Click()
    '/CARCEL
    
    On Error GoTo cmdCARCEL_Click_Err
    
    Dim tStr As String
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo de la pena.", "Carcel a " & Nick)
                
        If LenB(tStr) <> 0 Then
            tStr = tStr & "@" & InputBox("Indique el tiempo de condena (entre 0 y 60 minutos).", "Carcel a " & Nick)
            'We use the Parser to control the command format
            Call ParseUserCommand("/CARCEL " & Nick & "@" & tStr)

        End If

    End If

    
    Exit Sub

cmdCARCEL_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdCARCEL_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdCC_Click()
    '/CC
    
    On Error GoTo cmdCC_Click_Err
    
    Call WriteSpawnListRequest

    
    Exit Sub

cmdCC_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdCC_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdCHATCOLOR_Click()
    '/CHATCOLOR
    
    On Error GoTo cmdCHATCOLOR_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Defina el color (R G B). Deje en blanco para usar el default.", "Cambiar color del chat")
    
    Call ParseUserCommand("/CHATCOLOR " & tStr) 'We use the Parser to control the command format

    
    Exit Sub

cmdCHATCOLOR_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdCHATCOLOR_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdCI_Click()
    '/CI
    
    On Error GoTo cmdCI_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Indique el número del objeto a crear.", "Crear Objeto")

    If LenB(tStr) <> 0 Then If MsgBox("¿Seguro desea crear el objeto " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/CI " & tStr) 'We use the Parser to control the command format

    
    Exit Sub

cmdCI_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdCI_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdCIUMSG_Click()
    '/CIUMSG
    
    On Error GoTo cmdCIUMSG_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba un Mensaje.", "Mensaje por consola Ciudadanos")

    If LenB(tStr) <> 0 Then Call WriteCitizenMessage(tStr)

    
    Exit Sub

cmdCIUMSG_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdCIUMSG_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdCONDEN_Click()
    '/CONDEN
    
    On Error GoTo cmdCONDEN_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("¿Seguro desea volver criminal a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteTurnCriminal(Nick)

    
    Exit Sub

cmdCONDEN_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdCONDEN_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdConsulta_Click()
    
    On Error GoTo cmdConsulta_Click_Err
    
    WriteConsultation

    
    Exit Sub

cmdConsulta_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdConsulta_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdCT_Click()
    '/CT
    
    On Error GoTo cmdCT_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Indique la posición donde lleva el portal (MAPA X Y).", "Crear Portal")

    If LenB(tStr) <> 0 Then Call ParseUserCommand("/CT " & tStr) 'We use the Parser to control the command format

    
    Exit Sub

cmdCT_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdCT_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdDE_Click()
    
    On Error GoTo cmdDE_Click_Err
    

    '/DE
    If MsgBox("¿Seguro desea destruir el Tile Exit?", vbYesNo, "Atencion!") = vbYes Then Call WriteExitDestroy

    
    Exit Sub

cmdDE_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdDE_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdDEST_Click()
    
    On Error GoTo cmdDEST_Click_Err
    

    '/DEST
    If MsgBox("¿Seguro desea destruir el objeto sobre el que esta parado?", vbYesNo, "Atencion!") = vbYes Then Call WriteDestroyItems

    
    Exit Sub

cmdDEST_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdDEST_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdDONDE_Click()
    '/DONDE
    
    On Error GoTo cmdDONDE_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteWhere(Nick)

    
    Exit Sub

cmdDONDE_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdDONDE_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdDT_Click()
    
    On Error GoTo cmdDT_Click_Err
    

    'DT
    If MsgBox("¿Seguro desea destruir el portal?", vbYesNo, "Atencion!") = vbYes Then Call WriteTeleportDestroy

    
    Exit Sub

cmdDT_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdDT_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdECHAR_Click()
    '/ECHAR
    
    On Error GoTo cmdECHAR_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteKick(Nick)

    
    Exit Sub

cmdECHAR_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdECHAR_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdEJECUTAR_Click()
    '/EJECUTAR
    
    On Error GoTo cmdEJECUTAR_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("¿Seguro desea ejecutar a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteExecute(Nick)

    
    Exit Sub

cmdEJECUTAR_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdEJECUTAR_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdESTUPIDO_Click()
    '/ESTUPIDO
    
    On Error GoTo cmdESTUPIDO_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteMakeDumb(Nick)

    
    Exit Sub

cmdESTUPIDO_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdESTUPIDO_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdGMSG_Click()
    '/GMSG
    
    On Error GoTo cmdGMSG_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de GM")

    If LenB(tStr) <> 0 Then Call WriteGMMessage(tStr)

    
    Exit Sub

cmdGMSG_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdGMSG_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdHORA_Click()
    '/HORA
    
    On Error GoTo cmdHORA_Click_Err
    
    Call Protocol.WriteServerTime

    
    Exit Sub

cmdHORA_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdHORA_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdIGNORADO_Click()
    '/IGNORADO
    
    On Error GoTo cmdIGNORADO_Click_Err
    
    Call WriteIgnored

    
    Exit Sub

cmdIGNORADO_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdIGNORADO_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdINFO_Click()
    '/INFO
    
    On Error GoTo cmdINFO_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharInfo(Nick)

    
    Exit Sub

cmdINFO_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdINFO_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdINV_Click()
    '/INV
    
    On Error GoTo cmdINV_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharInventory(Nick)

    
    Exit Sub

cmdINV_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdINV_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdINVISIBLE_Click()
    '/INVISIBLE
    
    On Error GoTo cmdINVISIBLE_Click_Err
    
    Call WriteInvisible

    
    Exit Sub

cmdINVISIBLE_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdINVISIBLE_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdIP2NICK_Click()
    '/IP2NICK
    
    On Error GoTo cmdIP2NICK_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba la ip.", "IP to Nick")

    If LenB(tStr) <> 0 Then Call ParseUserCommand("/IP2NICK " & tStr) 'We use the Parser to control the command format

    
    Exit Sub

cmdIP2NICK_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdIP2NICK_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdIRA_Click()
    '/IRA
    
    On Error GoTo cmdIRA_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteGoToChar(Nick)

    
    Exit Sub

cmdIRA_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdIRA_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdIRCERCA_Click()
    '/IRCERCA
    
    On Error GoTo cmdIRCERCA_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteGoNearby(Nick)

    
    Exit Sub

cmdIRCERCA_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdIRCERCA_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdKICKCONSE_Click()
    'KICKCONSE
    
    On Error GoTo cmdKICKCONSE_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("¿Seguro desea destituir a " & Nick & " de su cargo de consejero?", vbYesNo, "Atencion!") = vbYes Then Call WriteCouncilKick(Nick)

    
    Exit Sub

cmdKICKCONSE_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdKICKCONSE_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdLASTEMAIL_Click()
    '/LASTEMAIL
    
    On Error GoTo cmdLASTEMAIL_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharMail(Nick)

    
    Exit Sub

cmdLASTEMAIL_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdLASTEMAIL_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdLASTIP_Click()
    '/LASTIP
    
    On Error GoTo cmdLASTIP_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteLastIP(Nick)

    
    Exit Sub

cmdLASTIP_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdLASTIP_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdLLUVIA_Click()
    '/LLUVIA
    
    On Error GoTo cmdLLUVIA_Click_Err
    
    Call WriteRainToggle

    
    Exit Sub

cmdLLUVIA_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdLLUVIA_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdMASSDEST_Click()
    
    On Error GoTo cmdMASSDEST_Click_Err
    

    '/MASSDEST
    If MsgBox("¿Seguro desea destruir todos los items del mapa?", vbYesNo, "Atencion!") = vbYes Then Call WriteDestroyAllItemsInArea

    
    Exit Sub

cmdMASSDEST_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdMASSDEST_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdMIEMBROSCLAN_Click()
    '/MIEMBROSCLAN
    
    On Error GoTo cmdMIEMBROSCLAN_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre del clan.", "Lista de miembros del clan")

    If LenB(tStr) <> 0 Then Call WriteGuildMemberList(tStr)

    
    Exit Sub

cmdMIEMBROSCLAN_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdMIEMBROSCLAN_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdMOTDCAMBIA_Click()
    '/MOTDCAMBIA
    
    On Error GoTo cmdMOTDCAMBIA_Click_Err
    
    Call WriteChangeMOTD

    
    Exit Sub

cmdMOTDCAMBIA_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdMOTDCAMBIA_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdNAVE_Click()
    '/NAVE
    
    On Error GoTo cmdNAVE_Click_Err
    
    Call WriteNavigateToggle

    
    Exit Sub

cmdNAVE_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdNAVE_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdNENE_Click()
    '/NENE
    
    On Error GoTo cmdNENE_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Indique el mapa.", "Número de NPCs enemigos.")

    If LenB(tStr) <> 0 Then Call ParseUserCommand("/NENE " & tStr) 'We use the Parser to control the command format

    
    Exit Sub

cmdNENE_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdNENE_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdNICK2IP_Click()
    '/NICK2IP
    
    On Error GoTo cmdNICK2IP_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteNickToIP(Nick)

    
    Exit Sub

cmdNICK2IP_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdNICK2IP_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdNOCAOS_Click()
    '/NOCAOS
    
    On Error GoTo cmdNOCAOS_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("¿Seguro desea expulsar a " & Nick & " de la legión oscura?", vbYesNo, "Atencion!") = vbYes Then Call WriteChaosLegionKick(Nick)

    
    Exit Sub

cmdNOCAOS_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdNOCAOS_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdNOESTUPIDO_Click()
    '/NOESTUPIDO
    
    On Error GoTo cmdNOESTUPIDO_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteMakeDumbNoMore(Nick)

    
    Exit Sub

cmdNOESTUPIDO_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdNOESTUPIDO_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdNOREAL_Click()
    '/NOREAL
    
    On Error GoTo cmdNOREAL_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("¿Seguro desea expulsar a " & Nick & " de la armada real?", vbYesNo, "Atencion!") = vbYes Then Call WriteRoyalArmyKick(Nick)

    
    Exit Sub

cmdNOREAL_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdNOREAL_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdOCULTANDO_Click()
    '/OCULTANDO
    
    On Error GoTo cmdOCULTANDO_Click_Err
    
    Call WriteHiding

    
    Exit Sub

cmdOCULTANDO_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdOCULTANDO_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdONLINECAOS_Click()
    '/ONLINECAOS
    
    On Error GoTo cmdONLINECAOS_Click_Err
    
    Call WriteOnlineChaosLegion

    
    Exit Sub

cmdONLINECAOS_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdONLINECAOS_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdONLINEGM_Click()
    '/ONLINEGM
    
    On Error GoTo cmdONLINEGM_Click_Err
    
    Call WriteOnlineGM

    
    Exit Sub

cmdONLINEGM_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdONLINEGM_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdONLINEMAP_Click()
    '/ONLINEMAP
    
    On Error GoTo cmdONLINEMAP_Click_Err
    
    Call WriteOnlineMap(UserMap)

    
    Exit Sub

cmdONLINEMAP_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdONLINEMAP_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdONLINEREAL_Click()
    '/ONLINEREAL
    
    On Error GoTo cmdONLINEREAL_Click_Err
    
    Call WriteOnlineRoyalArmy

    
    Exit Sub

cmdONLINEREAL_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdONLINEREAL_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdPENAS_Click()
    '/PENAS
    
    On Error GoTo cmdPENAS_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WritePunishments(Nick)

    
    Exit Sub

cmdPENAS_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdPENAS_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdPERDON_Click()
    '/PERDON
    
    On Error GoTo cmdPERDON_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteForgive(Nick)

    
    Exit Sub

cmdPERDON_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdPERDON_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdPISO_Click()
    '/PISO
    
    On Error GoTo cmdPISO_Click_Err
    
    Call WriteItemsInTheFloor

    
    Exit Sub

cmdPISO_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdPISO_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdRAJAR_Click()
    '/RAJAR
    
    On Error GoTo cmdRAJAR_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("¿Seguro desea resetear la facción de " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteResetFactions(Nick)

    
    Exit Sub

cmdRAJAR_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdRAJAR_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdRAJARCLAN_Click()
    '/RAJARCLAN
    
    On Error GoTo cmdRAJARCLAN_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("¿Seguro desea expulsar a " & Nick & " de su clan?", vbYesNo, "Atencion!") = vbYes Then Call WriteRemoveCharFromGuild(Nick)

    
    Exit Sub

cmdRAJARCLAN_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdRAJARCLAN_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdREALMSG_Click()
    '/REALMSG
    
    On Error GoTo cmdREALMSG_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba un Mensaje.", "Mensaje por consola ArmadaReal")

    If LenB(tStr) <> 0 Then Call WriteRoyalArmyMessage(tStr)

    
    Exit Sub

cmdREALMSG_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdREALMSG_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdRefresh_Click()
    
    On Error GoTo cmdRefresh_Click_Err
    
    Call ClearRecordDetails
    Call WriteRecordListRequest

    
    Exit Sub

cmdRefresh_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdRefresh_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdREM_Click()
    '/REM
    
    On Error GoTo cmdREM_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba el comentario.", "Comentario en el logGM")

    If LenB(tStr) <> 0 Then Call WriteComment(tStr)

    
    Exit Sub

cmdREM_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdREM_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdREVIVIR_Click()
    '/REVIVIR
    
    On Error GoTo cmdREVIVIR_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteReviveChar(Nick)

    
    Exit Sub

cmdREVIVIR_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdREVIVIR_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdRMSG_Click()
    '/RMSG
    
    On Error GoTo cmdRMSG_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de RoleMaster")

    If LenB(tStr) <> 0 Then Call WriteServerMessage(tStr)

    
    Exit Sub

cmdRMSG_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdRMSG_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdSETDESC_Click()
    '/SETDESC
    
    On Error GoTo cmdSETDESC_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba una DESC.", "Set Description")

    If LenB(tStr) <> 0 Then Call WriteSetCharDescription(tStr)

    
    Exit Sub

cmdSETDESC_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdSETDESC_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdSHOW_SOS_Click()
    '/SHOW SOS
    
    On Error GoTo cmdSHOW_SOS_Click_Err
    
    Call WriteSOSShowList

    
    Exit Sub

cmdSHOW_SOS_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdSHOW_SOS_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdSHOWCMSG_Click()
    '/SHOWCMSG
    
    On Error GoTo cmdSHOWCMSG_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre del clan que desea escuchar.", "Escuchar los mensajes del clan")

    If LenB(tStr) <> 0 Then Call WriteShowGuildMessages(tStr)

    
    Exit Sub

cmdSHOWCMSG_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdSHOWCMSG_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdSHOWNAME_Click()
    '/SHOWNAME
    
    On Error GoTo cmdSHOWNAME_Click_Err
    
    Call WriteShowName

    
    Exit Sub

cmdSHOWNAME_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdSHOWNAME_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdSILENCIAR_Click()
    '/SILENCIAR
    
    On Error GoTo cmdSILENCIAR_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteSilence(Nick)

    
    Exit Sub

cmdSILENCIAR_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdSILENCIAR_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdSKILLS_Click()
    '/SKILLS
    
    On Error GoTo cmdSKILLS_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharSkills(Nick)

    
    Exit Sub

cmdSKILLS_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdSKILLS_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdSMSG_Click()
    '/SMSG
    
    On Error GoTo cmdSMSG_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba el mensaje.", "Mensaje de sistema")

    If LenB(tStr) <> 0 Then Call WriteSystemMessage(tStr)

    
    Exit Sub

cmdSMSG_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdSMSG_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdSTAT_Click()
    '/STAT
    
    On Error GoTo cmdSTAT_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharStats(Nick)

    
    Exit Sub

cmdSTAT_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdSTAT_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdSUM_Click()
    '/SUM
    
    On Error GoTo cmdSUM_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteSummonChar(Nick)

    
    Exit Sub

cmdSUM_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdSUM_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdTALKAS_Click()
    '/TALKAS
    
    On Error GoTo cmdTALKAS_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba un Mensaje.", "Hablar por NPC")

    If LenB(tStr) <> 0 Then Call WriteTalkAsNPC(tStr)

    
    Exit Sub

cmdTALKAS_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdTALKAS_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdTELEP_Click()
    '/TELEP
    
    On Error GoTo cmdTELEP_Click_Err
    
    Dim tStr As String
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Indique la posición (MAPA X Y).", "Transportar a " & Nick)

        If LenB(tStr) <> 0 Then Call ParseUserCommand("/TELEP " & Nick & " " & tStr) 'We use the Parser to control the command format

    End If

    
    Exit Sub

cmdTELEP_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdTELEP_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdTRABAJANDO_Click()
    '/TRABAJANDO
    
    On Error GoTo cmdTRABAJANDO_Click_Err
    
    Call WriteWorking

    
    Exit Sub

cmdTRABAJANDO_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdTRABAJANDO_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdUNBAN_Click()
    '/UNBAN
    
    On Error GoTo cmdUNBAN_Click_Err
    
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("¿Seguro desea unbanear a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteUnbanChar(Nick)

    
    Exit Sub

cmdUNBAN_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdUNBAN_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdUNBANIP_Click()
    '/UNBANIP
    
    On Error GoTo cmdUNBANIP_Click_Err
    
    Dim tStr As String
    
    tStr = InputBox("Escriba el ip.", "Unbanear IP")

    If LenB(tStr) <> 0 Then If MsgBox("¿Seguro desea unbanear la ip " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/UNBANIP " & tStr) 'We use the Parser to control the command format

    
    Exit Sub

cmdUNBANIP_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdUNBANIP_Click"
    End If
Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call showTab(1)
    
    'Actualiza los usuarios online
    Call cmdActualiza_Click
    
    'Actualiza los seguimientos
    Call cmdRefresh_Click
    
    'Oculta el menú usado para el PopUp
    mnuSeguimientos.Visible = False

    
    Exit Sub

Form_Load_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "Form_Load"
    End If
Resume Next
    
End Sub

Private Sub cmdActualiza_Click()
    
    On Error GoTo cmdActualiza_Click_Err
    
    Call WriteRequestUserList
    Call FlushBuffer

    
    Exit Sub

cmdActualiza_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdActualiza_Click"
    End If
Resume Next
    
End Sub

Private Sub cmdCerrar_Click()
    
    On Error GoTo cmdCerrar_Click_Err
    
    Unload Me

    
    Exit Sub

cmdCerrar_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "cmdCerrar_Click"
    End If
Resume Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error GoTo Form_QueryUnload_Err
    
    Unload Me

    
    Exit Sub

Form_QueryUnload_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "Form_QueryUnload"
    End If
Resume Next
    
End Sub

Private Sub lstUsers_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    
    On Error GoTo lstUsers_MouseUp_Err
    

    If Button = vbRightButton Then
        PopupMenu mnuSeguimientos
    Else

        If lstUsers.ListIndex <> -1 Then
            Call ClearRecordDetails
            Call WriteRecordDetailsRequest(lstUsers.ListIndex + 1)

        End If

    End If

    
    Exit Sub

lstUsers_MouseUp_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "lstUsers_MouseUp"
    End If
Resume Next
    
End Sub

Private Sub ClearRecordDetails()
    
    On Error GoTo ClearRecordDetails_Err
    
    txtIP.Text = vbNullString
    txtCreador.Text = vbNullString
    txtDescrip.Text = vbNullString
    txtObs.Text = vbNullString
    txtTimeOn.Text = vbNullString
    lblEstado.Caption = vbNullString

    
    Exit Sub

ClearRecordDetails_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "ClearRecordDetails"
    End If
Resume Next
    
End Sub

Private Sub mnuDelete_Click()
    
    On Error GoTo mnuDelete_Click_Err
    

    With lstUsers

        If .ListIndex = -1 Then
            Call MsgBox("¡Seleccione un usuario para remover el seguimiento!", vbOKOnly + vbExclamation)
            Exit Sub

        End If
        
        If MsgBox("¿Desea eliminar el seguimiento al personaje " & .List(.ListIndex) & "?", vbYesNo) = vbYes Then
            Call WriteRecordRemove(.ListIndex + 1)
            Call ClearRecordDetails

        End If

    End With

    
    Exit Sub

mnuDelete_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "mnuDelete_Click"
    End If
Resume Next
    
End Sub

Private Sub mnuIra_Click()
    
    On Error GoTo mnuIra_Click_Err
    

    With lstUsers

        If .ListIndex <> -1 Then
            Call WriteGoToChar(.List(.ListIndex))

        End If

    End With

    
    Exit Sub

mnuIra_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "mnuIra_Click"
    End If
Resume Next
    
End Sub

Private Sub mnuSum_Click()
    
    On Error GoTo mnuSum_Click_Err
    

    With lstUsers

        If .ListIndex <> -1 Then
            Call WriteSummonChar(.List(.ListIndex))

        End If

    End With

    
    Exit Sub

mnuSum_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "mnuSum_Click"
    End If
Resume Next
    
End Sub

Private Sub TabStrip_Click()
    
    On Error GoTo TabStrip_Click_Err
    
    Call showTab(TabStrip.SelectedItem.Index)

    
    Exit Sub

TabStrip_Click_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "TabStrip_Click"
    End If
Resume Next
    
End Sub

Private Sub showTab(TabId As Byte)
    
    On Error GoTo showTab_Err
    
    Dim i As Byte
    
    For i = 1 To Frame.UBound
        Frame(i).Visible = (i = TabId)
    Next i
    
    With Frame(TabId)
        frmPanelGm.Height = .Height + 1280
        TabStrip.Height = .Height + 480
        cmdCerrar.Top = .Height + 465

    End With

    
    Exit Sub

showTab_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "showTab"
    End If
Resume Next
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyUp_Err
    
    If KeyCode = vbKeyEscape Then
        Set clsFormulario = Nothing
        Unload Me
    End If

    Exit Sub

Form_KeyUp_Err:

    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "frmPanelGm" & "->" & "Form_KeyUp"

    End If

    Resume Next
    
End Sub

