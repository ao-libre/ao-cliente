VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPanelGm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel GM"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   4200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6600
      Width           =   4215
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   1575
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/GMSG"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/HORA"
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/RMSG"
         Height          =   315
         Index           =   36
         Left            =   1440
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/REALMSG"
         Height          =   315
         Index           =   43
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/CAOSMSG"
         Height          =   315
         Index           =   44
         Left            =   1440
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/CIUMSG"
         Height          =   315
         Index           =   45
         Left            =   2640
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/TALKAS"
         Height          =   315
         Index           =   46
         Left            =   2880
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/MOTDCAMBIA"
         Height          =   315
         Index           =   66
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/SMSG"
         Height          =   315
         Index           =   67
         Left            =   2640
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/SHOWNAME"
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   63
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/REM"
         Height          =   315
         Index           =   5
         Left            =   600
         TabIndex        =   62
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/INVISIBLE"
         Height          =   315
         Index           =   14
         Left            =   600
         TabIndex        =   61
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/SETDESC"
         Height          =   315
         Index           =   42
         Left            =   2280
         TabIndex        =   60
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/NAVE"
         Height          =   315
         Index           =   68
         Left            =   600
         TabIndex        =   59
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/CHATCOLOR"
         Height          =   315
         Index           =   75
         Left            =   600
         TabIndex        =   58
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/IGNORADO"
         Height          =   315
         Index           =   76
         Left            =   2280
         TabIndex        =   57
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   5
      Left            =   120
      TabIndex        =   55
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/CC"
         Height          =   315
         Index           =   34
         Left            =   720
         TabIndex        =   72
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/LIMPIAR"
         Height          =   315
         Index           =   35
         Left            =   2280
         TabIndex        =   71
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/CT"
         Height          =   315
         Index           =   39
         Left            =   720
         TabIndex        =   70
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/DT"
         Height          =   315
         Index           =   40
         Left            =   2280
         TabIndex        =   69
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/LLUVIA"
         Height          =   315
         Index           =   41
         Left            =   720
         TabIndex        =   68
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/MASSDEST"
         Height          =   315
         Index           =   47
         Left            =   2280
         TabIndex        =   67
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/PISO"
         Height          =   315
         Index           =   50
         Left            =   720
         TabIndex        =   66
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/CI"
         Height          =   315
         Index           =   60
         Left            =   720
         TabIndex        =   65
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/DEST"
         Height          =   315
         Index           =   61
         Left            =   2280
         TabIndex        =   64
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/ONLINEREAL"
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   21
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/ONLINECAOS"
         Height          =   315
         Index           =   3
         Left            =   2160
         TabIndex        =   20
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/NENE"
         Height          =   315
         Index           =   8
         Left            =   480
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/SHOW SOS"
         Height          =   315
         Index           =   11
         Left            =   480
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/TRABAJANDO"
         Height          =   315
         Index           =   15
         Left            =   480
         TabIndex        =   17
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/OCULTANDO"
         Height          =   315
         Index           =   16
         Left            =   2160
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/ONLINEGM"
         Height          =   315
         Index           =   26
         Left            =   480
         TabIndex        =   15
         Top             =   2160
         Width           =   3015
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/ONLINEMAP"
         Height          =   315
         Index           =   27
         Left            =   2160
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/BORRAR SOS"
         Height          =   315
         Index           =   74
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   3375
      Index           =   6
      Left            =   120
      TabIndex        =   56
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/SHOWCMSG"
         Height          =   315
         Index           =   73
         Left            =   480
         TabIndex        =   85
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/BANCLAN"
         Height          =   315
         Index           =   57
         Left            =   480
         TabIndex        =   84
         Top             =   3000
         Width           =   3015
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/MIEMBROSCLAN"
         Height          =   315
         Index           =   56
         Left            =   1920
         TabIndex        =   83
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/BANIPRELOAD"
         Height          =   315
         Index           =   55
         Left            =   1440
         TabIndex        =   82
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/BANIPLIST"
         Height          =   315
         Index           =   54
         Left            =   240
         TabIndex        =   81
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/IP2NICK"
         Height          =   315
         Index           =   38
         Left            =   240
         TabIndex        =   80
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/ACEPTCONSE"
         Height          =   315
         Index           =   48
         Left            =   240
         TabIndex        =   79
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/ACEPTCONSECAOS"
         Height          =   315
         Index           =   49
         Left            =   240
         TabIndex        =   78
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/KICKCONSE"
         Height          =   675
         Index           =   53
         Left            =   2520
         TabIndex        =   77
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/BANIP"
         Height          =   315
         Index           =   58
         Left            =   1440
         TabIndex        =   76
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/UNBANIP"
         Height          =   315
         Index           =   59
         Left            =   2520
         TabIndex        =   75
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/NOCAOS"
         Height          =   315
         Index           =   62
         Left            =   240
         TabIndex        =   74
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/NOREAL"
         Height          =   315
         Index           =   63
         Left            =   1440
         TabIndex        =   73
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   6135
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   3975
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
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   3675
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/IRCERCA"
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   52
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/DONDE"
         Height          =   675
         Index           =   7
         Left            =   120
         TabIndex        =   51
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/PENAS"
         Height          =   315
         Index           =   12
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/TELEP"
         Height          =   315
         Index           =   9
         Left            =   2520
         TabIndex        =   49
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/SILENCIAR"
         Height          =   315
         Index           =   10
         Left            =   1320
         TabIndex        =   48
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/IRA"
         Height          =   315
         Index           =   13
         Left            =   2520
         TabIndex        =   47
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/CARCEL"
         Height          =   315
         Index           =   17
         Left            =   1320
         TabIndex        =   46
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/ADVERTENCIA"
         Height          =   315
         Index           =   18
         Left            =   2520
         TabIndex        =   45
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/INFO"
         Height          =   315
         Index           =   19
         Left            =   120
         TabIndex        =   44
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/STAT"
         Height          =   315
         Index           =   20
         Left            =   1320
         TabIndex        =   43
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/BAL"
         Height          =   315
         Index           =   21
         Left            =   2520
         TabIndex        =   42
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/INV"
         Height          =   315
         Index           =   22
         Left            =   120
         TabIndex        =   41
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/BOV"
         Height          =   315
         Index           =   23
         Left            =   1320
         TabIndex        =   40
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/SKILLS"
         Height          =   315
         Index           =   24
         Left            =   2520
         TabIndex        =   39
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/REVIVIR"
         Height          =   315
         Index           =   25
         Left            =   120
         TabIndex        =   38
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/PERDON"
         Height          =   315
         Index           =   28
         Left            =   1320
         TabIndex        =   37
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/ECHAR"
         Height          =   315
         Index           =   29
         Left            =   120
         TabIndex        =   36
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/EJECUTAR"
         Height          =   315
         Index           =   30
         Left            =   1320
         TabIndex        =   35
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/BAN"
         Height          =   315
         Index           =   31
         Left            =   120
         TabIndex        =   34
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/UNBAN"
         Height          =   315
         Index           =   32
         Left            =   1320
         TabIndex        =   33
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/SUM"
         Height          =   315
         Index           =   33
         Left            =   1320
         TabIndex        =   32
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/NICK2IP"
         Height          =   315
         Index           =   37
         Left            =   120
         TabIndex        =   31
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/ESTUPIDO"
         Height          =   315
         Index           =   51
         Left            =   120
         TabIndex        =   30
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/NOESTUPIDO"
         Height          =   315
         Index           =   52
         Left            =   1320
         TabIndex        =   29
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/BORRARPENA"
         Height          =   315
         Index           =   64
         Left            =   2520
         TabIndex        =   28
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/LASTIP"
         Height          =   315
         Index           =   65
         Left            =   1320
         TabIndex        =   27
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/CONDEN"
         Height          =   315
         Index           =   69
         Left            =   120
         TabIndex        =   26
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/RAJAR"
         Height          =   315
         Index           =   70
         Left            =   2760
         TabIndex        =   25
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/RAJARCLAN"
         Height          =   315
         Index           =   71
         Left            =   2520
         TabIndex        =   24
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton cmdAccion 
         Caption         =   "/LASTEMAIL"
         Height          =   315
         Index           =   72
         Left            =   2520
         TabIndex        =   23
         Top             =   2880
         Width           =   1215
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   2055
      Left            =   0
      TabIndex        =   86
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3625
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
Option Explicit

Private Sub cmdAccion_Click(index As Integer)
    Dim Nick As String
    Dim tStr As String
    
    Nick = cboListaUsus.Text

    Select Case index
        Case 0 '/GMSG
            tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de GM")
            If LenB(tStr) <> 0 Then _
                Call WriteGMMessage(tStr)

        Case 1 '/SHOWNAME
            Call WriteShowName
        
        Case 2 '/ONLINEREAL
            Call WriteOnlineRoyalArmy
        
        Case 3 '/ONLINECAOS
            Call WriteOnlineChaosLegion
        
        Case 4 '/IRCERCA
            If LenB(Nick) <> 0 Then _
                Call WriteGoNearby(Nick)
                
        Case 5 '/REM
            tStr = InputBox("Escriba el comentario.", "Comentario en el logGM")
            If LenB(tStr) <> 0 Then _
                Call WriteComment(tStr)
        
        Case 6 '/HORA
            Call Protocol.WriteServerTime
        
        Case 7 '/DONDE
            If LenB(Nick) <> 0 Then _
                Call WriteWhere(Nick)
                
        Case 8 '/NENE
            tStr = InputBox("Indique el mapa.", "Número de NPCs enemigos.")
            If LenB(tStr) <> 0 Then _
                Call ParseUserCommand("/NENE " & tStr) 'We use the Parser to control the command format
            
        Case 9 '/TELEP
            If LenB(Nick) <> 0 Then
                tStr = InputBox("Indique la posición (MAPA X Y).", "Transportar a " & Nick)
                If LenB(tStr) <> 0 Then _
                    Call ParseUserCommand("/TELEP " & Nick & " " & tStr) 'We use the Parser to control the command format
            End If
        
        Case 10 '/SILENCIAR
            If LenB(Nick) <> 0 Then _
                Call WriteSilence(Nick)
        
        Case 11 '/SHOW SOS
            Call WriteSOSShowList
        
        Case 12 '/PENAS
            If LenB(Nick) <> 0 Then _
                Call WritePunishments(Nick)

        Case 13 '/IRA
            If LenB(Nick) <> 0 Then _
                Call WriteGoToChar(Nick)
        
        Case 14 '/INVISIBLE
            Call WriteInvisible

        Case 15 '/TRABAJANDO
            Call WriteWorking
        
        Case 16 '/OCULTANDO
            Call WriteHiding
        
        Case 17 '/CARCEL
            If LenB(Nick) <> 0 Then
                tStr = InputBox("Escriba el motivo de la pena.", "Carcel a " & Nick)
                
                If LenB(tStr) <> 0 Then
                    tStr = tStr & "@" & InputBox("Indique el tiempo de condena (entre 0 y 60 minutos).", "Carcel a " & Nick)
                    'We use the Parser to control the command format
                    Call ParseUserCommand("/CARCEL " & Nick & "@" & tStr)
                End If
            End If
        
        Case 18 '/ADVERTENCIA
            If LenB(Nick) <> 0 Then
                tStr = InputBox("Escriba el motivo de la advertencia.", "Advertir a " & Nick)
                
                If LenB(tStr) <> 0 Then
                    'We use the Parser to control the command format
                    Call ParseUserCommand("/ADVERTENCIA " & Nick & "@" & tStr)
                End If
            End If
                        
        Case 19 '/INFO
            If LenB(Nick) <> 0 Then _
                Call WriteRequestCharInfo(Nick)
        
        Case 20 '/STAT
            If LenB(Nick) <> 0 Then _
                Call WriteRequestCharStats(Nick)
        
        Case 21 '/BAL
            If LenB(Nick) <> 0 Then _
                Call WriteRequestCharGold(Nick)

        Case 22 '/INV"
            If LenB(Nick) <> 0 Then _
                Call WriteRequestCharInventory(Nick)

        Case 23 '/BOV
            If LenB(Nick) <> 0 Then _
                Call WriteRequestCharBank(Nick)

        Case 24 '/SKILLS
            If LenB(Nick) <> 0 Then _
                Call WriteRequestCharSkills(Nick)
                
        Case 25 '/REVIVIR
            If LenB(Nick) <> 0 Then _
                Call WriteReviveChar(Nick)

        Case 26 '/ONLINEGM
            Call WriteOnlineGM
        
        Case 27 '/ONLINEMAP
            Call WriteOnlineMap
        
        Case 28 '/PERDON
            If LenB(Nick) <> 0 Then _
                Call WriteForgive(Nick)

        Case 29 '/ECHAR
            If LenB(Nick) <> 0 Then _
                Call WriteKick(Nick)
                        
        Case 30 '/EJECUTAR
            If LenB(Nick) <> 0 Then _
                If MsgBox("¿Seguro desea ejecutar a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
                    Call WriteExecute(Nick)
                
        Case 31 '/BAN
            If LenB(Nick) <> 0 Then
                tStr = InputBox("Escriba el motivo del ban.", "BAN a " & Nick)
                
                If LenB(tStr) <> 0 Then _
                    If MsgBox("¿Seguro desea banear a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
                        Call WriteBanChar(Nick, tStr)
            End If
            
        Case 32 '/UNBAN
            If LenB(Nick) <> 0 Then _
                If MsgBox("¿Seguro desea unbanear a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
                    Call WriteUnbanChar(Nick)
                
        Case 33 '/SUM
            If LenB(Nick) <> 0 Then _
                Call WriteSummonChar(Nick)
                
        Case 34 '/CC
            Call WriteSpawnListRequest
                
        Case 35 '/LIMPIAR
            Call WriteCleanWorld
                
        Case 36 '/RMSG
            tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de RoleMaster")
            If LenB(tStr) <> 0 Then _
                Call WriteServerMessage(tStr)
                
        Case 37 '/NICK2IP
            If LenB(Nick) <> 0 Then _
                Call WriteNickToIP(Nick)
                
        Case 38 '/IP2NICK
            tStr = InputBox("Escriba la ip.", "IP to Nick")
            If LenB(tStr) <> 0 Then _
                Call ParseUserCommand("/IP2NICK " & tStr) 'We use the Parser to control the command format
                
        Case 39 '/CT
            tStr = InputBox("Indique la posición donde lleva el portal (MAPA X Y).", "Crear Portal")
            If LenB(tStr) <> 0 Then _
                Call ParseUserCommand("/CT " & tStr) 'We use the Parser to control the command format
                
        Case 40 'DT
            If MsgBox("¿Seguro desea destruir el portal?", vbYesNo, "Atencion!") = vbYes Then _
                Call WriteTeleportDestroy
        
        Case 41 '/LLUVIA
            Call WriteRainToggle
                
        Case 42 '/SETDESC
            tStr = InputBox("Escriba una DESC.", "Set Description")
            If LenB(tStr) <> 0 Then _
                Call WriteSetCharDescription(tStr)
                
        Case 43 '/REALMSG
            tStr = InputBox("Escriba un Mensaje.", "Mensaje por consola ArmadaReal")
            If LenB(tStr) <> 0 Then _
                Call WriteRoyalArmyMessage(tStr)
                 
        Case 44 '/CAOSMSG
            tStr = InputBox("Escriba un Mensaje.", "Mensaje por consola LegionOscura")
            If LenB(tStr) <> 0 Then _
                Call WriteChaosLegionMessage(tStr)
                
        Case 45 '/CIUMSG
            tStr = InputBox("Escriba un Mensaje.", "Mensaje por consola Ciudadanos")
            If LenB(tStr) <> 0 Then _
                Call WriteCitizenMessage(tStr)
                
        Case 46 '/TALKAS
            tStr = InputBox("Escriba un Mensaje.", "Hablar por NPC")
            If LenB(tStr) <> 0 Then _
                Call WriteTalkAsNPC(tStr)
        
        Case 47 '/MASSDEST
            If MsgBox("¿Seguro desea destruir todos los items del mapa?", vbYesNo, "Atencion!") = vbYes Then _
                Call WriteDestroyAllItemsInArea
    
        Case 48 '/ACEPTCONSE
            If LenB(Nick) <> 0 Then _
                If MsgBox("¿Seguro desea aceptar a " & Nick & " como consejero real?", vbYesNo, "Atencion!") = vbYes Then _
                    Call WriteAcceptRoyalCouncilMember(Nick)
                
        Case 49 '/ACEPTCONSECAOS
            If LenB(Nick) <> 0 Then _
                If MsgBox("¿Seguro desea aceptar a " & Nick & " como consejero del caos?", vbYesNo, "Atencion!") = vbYes Then _
                    Call WriteAcceptChaosCouncilMember(Nick)
                
        Case 50 '/PISO
            Call WriteItemsInTheFloor
                
        Case 51 '/ESTUPIDO
            If LenB(Nick) <> 0 Then _
                Call WriteMakeDumb(Nick)
                
        Case 52 '/NOESTUPIDO
            If LenB(Nick) <> 0 Then _
                Call WriteMakeDumbNoMore(Nick)
                
        Case 53 'KICKCONSE
            If LenB(Nick) <> 0 Then _
                If MsgBox("¿Seguro desea destituir a " & Nick & " de su cargo de consejero?", vbYesNo, "Atencion!") = vbYes Then _
                    Call WriteCouncilKick(Nick)
                
        Case 54 '/BANIPLIST
            Call WriteBannedIPList
                
        Case 55 '/BANIPRELOAD
            Call WriteBannedIPReload
                
        Case 56 '/MIEMBROSCLAN
            tStr = InputBox("Escriba el nombre del clan.", "Lista de miembros del clan")
            If LenB(tStr) <> 0 Then _
                Call WriteGuildMemberList(tStr)
                
        Case 57 '/BANCLAN
            tStr = InputBox("Escriba el nombre del clan.", "Banear clan")
            If LenB(tStr) <> 0 Then _
                If MsgBox("¿Seguro desea banear al clan " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
                    Call WriteGuildBan(tStr)
                
        Case 58 '/BANIP
            tStr = InputBox("Escriba el ip.", "Banear IP")
            If LenB(tStr) <> 0 Then _
                If MsgBox("¿Seguro desea banear la ip " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
                    Call ParseUserCommand("/BANIP " & tStr) 'We use the Parser to control the command format
                
        Case 59 '/UNBANIP
            tStr = InputBox("Escriba el ip.", "Unbanear IP")
            If LenB(tStr) <> 0 Then _
                If MsgBox("¿Seguro desea unbanear la ip " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
                    Call ParseUserCommand("/UNBANIP " & tStr) 'We use the Parser to control the command format
                
        Case 60 '/CI
            tStr = InputBox("Indique el número del objeto a crear.", "Crear Objeto")
            If LenB(tStr) <> 0 Then _
                If MsgBox("¿Seguro desea crear el objeto " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then _
                    Call ParseUserCommand("/CI " & tStr) 'We use the Parser to control the command format
                
        Case 61 '/DEST
            If MsgBox("¿Seguro desea destruir el objeto sobre el que esta parado?", vbYesNo, "Atencion!") = vbYes Then _
                Call WriteDestroyItems
                
        Case 62 '/NOCAOS
            If LenB(Nick) <> 0 Then _
                If MsgBox("¿Seguro desea expulsar a " & Nick & " de la legión oscura?", vbYesNo, "Atencion!") = vbYes Then _
                    Call WriteChaosLegionKick(Nick)
    
        Case 63 '/NOREAL
            If LenB(Nick) <> 0 Then _
                If MsgBox("¿Seguro desea expulsar a " & Nick & " de la armada real?", vbYesNo, "Atencion!") = vbYes Then _
                    Call WriteRoyalArmyKick(Nick)
                
        Case 64 '/BORRARPENA
            If LenB(Nick) <> 0 Then
                tStr = InputBox("Indique el número de la pena a borrar.", "Borrar pena")
                If LenB(tStr) <> 0 Then _
                    If MsgBox("¿Seguro desea borrar la pena " & tStr & " a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
                        Call ParseUserCommand("/BORRARPENA " & Nick & "@" & tStr) 'We use the Parser to control the command format
            End If

        Case 65 '/LASTIP
            If LenB(Nick) <> 0 Then _
                Call WriteLastIP(Nick)
    
        Case 66 '/MOTDCAMBIA
            Call WriteChangeMOTD
                
        Case 67 '/SMSG
            tStr = InputBox("Escriba el mensaje.", "Mensaje de sistema")
            If LenB(tStr) <> 0 Then _
                Call WriteSystemMessage(tStr)
            
        Case 68 '/NAVE
            Call WriteNavigateToggle
                
        Case 69 '/CONDEN
            If LenB(Nick) <> 0 Then _
                If MsgBox("¿Seguro desea volver criminal a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
                    Call WriteTurnCriminal(Nick)
                
        Case 70 '/RAJAR
            If LenB(Nick) <> 0 Then _
                If MsgBox("¿Seguro desea resetear la faccion de " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
                    Call WriteResetFactions(Nick)
                
        Case 71 '/RAJARCLAN
            If LenB(Nick) <> 0 Then _
                If MsgBox("¿Seguro desea expulsar a " & Nick & " de su clan?", vbYesNo, "Atencion!") = vbYes Then _
                    Call WriteRemoveCharFromGuild(Nick)
                
        Case 72 '/LASTEMAIL
            If LenB(Nick) <> 0 Then _
                Call WriteRequestCharMail(Nick)
                
        Case 73 '/SHOWCMSG
            tStr = InputBox("Escriba el nombre del clan que desea escuchar.", "Escuchar los mensajes del clan")
            If LenB(tStr) <> 0 Then _
                Call WriteShowGuildMessages(tStr)
                
        Case 74 '/BORRAR SOS
            If MsgBox("¿Seguro desea borrar el SOS?", vbYesNo, "Atencion!") = vbYes Then _
                Call WriteCleanSOS

        Case 75 '/CHATCOLOR
            tStr = InputBox("Defina el color (R G B).", "Cambiar color del chat")
            If LenB(tStr) <> 0 Then _
                Call ParseUserCommand("/CHATCOLOR " & tStr) 'We use the Parser to control the command format
                
        Case 76 '/IGNORADO
            Call WriteIgnored

     End Select
End Sub

Private Sub cmdActualiza_Click()
    Call WriteRequestUserList
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call showTab(1)
    Call cmdActualiza_Click
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
        If i = TabId Then
            Frame(i).Visible = True
        Else
            Frame(i).Visible = False
        End If
    Next i
    
    TabStrip.Height = Frame(TabId).Height + 480
    cmdCerrar.Top = Frame(TabId).Height + 465
    frmPanelGm.Height = Frame(TabId).Height + 1215
End Sub
