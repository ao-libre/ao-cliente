VERSION 5.00
Begin VB.Form frmCreditos 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Creditos Argentum Online Libre"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10020
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCredits 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   6975
      Left            =   480
      ScaleHeight     =   6975
      ScaleWidth      =   7455
      TabIndex        =   0
      Top             =   360
      Width           =   7455
   End
   Begin VB.Timer Timer1 
      Left            =   9360
      Top             =   6480
   End
   Begin AOLibre.uAOButton cCodigo 
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   6240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      TX              =   "Codigo Fuente"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCreditos.frx":0000
      PICF            =   "frmCreditos.frx":001C
      PICH            =   "frmCreditos.frx":0038
      PICV            =   "frmCreditos.frx":0054
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton cCerrar 
      Height          =   495
      Left            =   8040
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCreditos.frx":0070
      PICF            =   "frmCreditos.frx":008C
      PICH            =   "frmCreditos.frx":00A8
      PICV            =   "frmCreditos.frx":00C4
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CreditLine()    As String
Private CreditLeft()    As Long
Private ScrollSpeed     As Integer
Private LinesVisible    As Integer
Private CharHeight      As Integer
Private TotalLines      As Integer
Private FadeIn          As Long
Private FadeOut         As Long
Private ColText         As Long
Private cDiff1          As Long
Private cDiff2          As Double
Private cDiff3          As Double
Private ColorFades(100) As Long
Private Yscroll         As Long
Private LinesOffset     As Integer
Private StopScroll      As Boolean

Private Sub Form_Load()
On Error Resume Next
    Call LoadAOCustomControlsPictures(Me)
    Call LoadTextsForm
    
    Me.Picture = LoadPicture(Game.path(Interfaces) & "frmCargando.jpg")

    Call StartCredits

    Call Audio.PlayBackgroundMusic("Creditos", MusicTypes.Mp3)
End Sub

Private Sub LoadTextsForm()
    Me.cCodigo.Caption = JsonLanguage.item("BTN_CODIGO_FUENTE").item("TEXTO")
    Me.cCerrar.Caption = JsonLanguage.item("BTN_SALIR").item("TEXTO")
End Sub

Private Sub cCerrar_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call Audio.PlayBackgroundMusic("2", MusicTypes.Mp3)

    Unload Me
End Sub

Private Sub cCodigo_Click()
'***********************************
'IMPORTANTE!
'
'No debe eliminarse la posibilidad de bajar el codigo de su servidor de esta forma.
'Caso contrario estaran violando la licencia Affero GPL y con ella derechos de autor,
'incurriendo de esta forma en un delito punible por ley.
'
'Argentum Online es Libre, es de todos. Mantengamoslo asi. Si tanto te gusta el juego y queres los
'cambios que hacemos nosotros, comparti los tuyos. Es un cambio justo. Si no estas de acuerdo,
'no uses nuestro codigo, pues nadie te obliga o bien utiliza una version anterior a la 0.12.0.
'***********************************
   Call Audio.PlayWave(SND_CLICK)

   Call MsgBox("No debe eliminarse la posibilidad de bajar el codigo de su servidor. Caso contrario estaran violando la licencia Affero GPL y con ella derechos de autor, incurriendo de esta forma en un delito punible por ley." & vbCrLf & vbCrLf & vbCrLf & _
                "Argentum Online es Libre, es de todos. Mantengamoslo asi. Si tanto te gusta el juego y queres los cambios que hacemos nosotros, comparti los tuyos. Es un cambio justo.", vbCritical Or vbApplicationModal)
    

   Call ShellExecute(0, "Open", "https://github.com/ao-libre", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StopScroll = False
End Sub

Private Sub StartCredits()
On Error Resume Next
   Dim FileO       As Integer
   Dim FileName    As String
   Dim tmp         As String
   Dim i           As Integer

   Dim Rcol1       As Long
   Dim Gcol1       As Long
   Dim Bcol1       As Long

   Dim Rcol2       As Long
   Dim Gcol2       As Long
   Dim Bcol2       As Long

   Dim Rfade       As Long
   Dim Gfade       As Long
   Dim Bfade       As Long

   Dim PercentFade As Integer
   Dim TimeInterval As Integer
   Dim AlignText  As Integer

   PercentFade = 20
   TimeInterval = 10
   ScrollSpeed = 10
   AlignText = 2 '( 1=left 2=center 3=right )
   LinesVisible = (picCredits.Height / picCredits.TextHeight("A")) + 1

   For i = 1 To LinesVisible
      ReDim Preserve CreditLine(TotalLines) As String
      CreditLine(TotalLines) = tmp
      TotalLines = TotalLines + 1
   Next

   FileO = FreeFile
   FileName = Game.path(INIT) & "\Creditos.txt"
   
   If Dir(FileName) = vbNullString Then
      TotalLines = 5
      ReDim Preserve CreditLine(TotalLines) As String
      CreditLine(0) = ""
      CreditLine(1) = "Para mas informacion visita:"
      CreditLine(2) = "http://www.ArgentumOnline.org"
      CreditLine(3) = ""
      CreditLine(4) = "~"
   Else
      On Error GoTo errhandler
      Open FileName For Input As FileO
      While Not EOF(FileO)
         Line Input #FileO, tmp
         ReDim Preserve CreditLine(TotalLines) As String
         CreditLine(TotalLines) = tmp
         TotalLines = TotalLines + 1
         Wend
      Close #FileO
   End If

   Me.Timer1.Interval = TimeInterval
   LinesVisible = (picCredits.Height / picCredits.TextHeight("A")) + 1
   CharHeight = picCredits.TextHeight("A")

   If PercentFade <> 0 Then
      FadeOut = ((picCredits.Height / 100) * PercentFade) - CharHeight
      FadeIn = (picCredits.Height - FadeOut) - CharHeight - CharHeight
   Else
      FadeIn = picCredits.Height
      FadeOut = 0 - CharHeight
   End If

   ColText = picCredits.ForeColor
   cDiff1 = (picCredits.Height - (CharHeight - 10)) - FadeIn
   cDiff2 = 100 / cDiff1
   cDiff3 = 100 / FadeOut
   ReDim CreditLeft(TotalLines - 1)

   For i = 0 To TotalLines - 1
      Select Case AlignText
      Case 1
         CreditLeft(i) = 100
      Case 2
         CreditLeft(i) = (picCredits.Width - picCredits.TextWidth(CreditLine(i))) / 2
      Case 3
         CreditLeft(i) = picCredits.Width - picCredits.TextWidth(CreditLine(i)) - 100
      End Select
   Next i

   Rcol1 = picCredits.ForeColor Mod 256
   Gcol1 = (picCredits.ForeColor And vbGreen) / 256
   Bcol1 = (picCredits.ForeColor And vbBlue) / 65536
   Rcol2 = picCredits.BackColor Mod 256
   Gcol2 = (picCredits.BackColor And vbGreen) / 256
   Bcol2 = (picCredits.BackColor And vbBlue) / 65536
   
   For i = 0 To 100
      Rfade = Rcol2 + ((Rcol1 - Rcol2) / 100) * i: If Rfade < 0 Then Rfade = 0
      Gfade = Gcol2 + ((Gcol1 - Gcol2) / 100) * i: If Gfade < 0 Then Gfade = 0
      Bfade = Bcol2 + ((Bcol1 - Bcol2) / 100) * i: If Bfade < 0 Then Bfade = 0
      ColorFades(i) = RGB(Rfade, Gfade, Bfade)
   Next

   StopScroll = False
   Me.Timer1.Enabled = True

Exit Sub
errhandler:
   Close FileO

End Sub

Private Sub picCredits_Click()
    StopScroll = Not StopScroll
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
   Dim Ycurr       As Long
   Dim textLine    As Integer
   Dim ColPrct     As Long
   Dim i           As Integer

   If StopScroll = True Then Exit Sub

   picCredits.Cls
   Yscroll = Yscroll - ScrollSpeed
   If Yscroll < (0 - CharHeight) Then
      Yscroll = 0
      LinesOffset = LinesOffset + 1
      If LinesOffset > TotalLines - 1 Then LinesOffset = 0
   End If
   picCredits.CurrentY = Yscroll
   Ycurr = Yscroll
   For i = 1 To LinesVisible
      If Ycurr > FadeIn And Ycurr < picCredits.Height Then
         ColPrct = cDiff2 * (cDiff1 - (Ycurr - FadeIn))
         If ColPrct < 0 Then ColPrct = 0
         If ColPrct > 100 Then ColPrct = 100
         picCredits.ForeColor = ColorFades(ColPrct)
      ElseIf Ycurr < FadeOut Then
         ColPrct = cDiff3 * Ycurr
         If ColPrct < 0 Then ColPrct = 0
         If ColPrct > 100 Then ColPrct = 100
         picCredits.ForeColor = ColorFades(ColPrct)
      Else
         picCredits.ForeColor = ColText
      End If
      textLine = (i + LinesOffset) Mod TotalLines
      picCredits.CurrentX = CreditLeft(textLine)
      picCredits.Print CreditLine(textLine)
      Ycurr = Ycurr + CharHeight
   Next i
End Sub
