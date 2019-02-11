VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoundMapas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private Type p_tSoundMapInfo

    cantidad As Integer
    SoundIndex() As Integer
    Flags() As Long
    Probabilidad() As Single

End Type
 
Private Enum p_eSoundFlags

    ninguna = 0
    Lluvia = 1

End Enum
 
Private p_Mapas()     As p_tSoundMapInfo
Private Const NumMaps As Integer = 291
 
Private Sub Class_Initialize()
    'armar el array
    
    On Error GoTo Class_Initialize_Err
    
    ReDim p_Mapas(1 To NumMaps) As p_tSoundMapInfo
    Call LoadSoundMapInfo

    
    Exit Sub

Class_Initialize_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsSoundMapas" & "->" & "Class_Initialize"
    End If
Resume Next
    
End Sub
 
Public Sub LoadSoundMapInfo()
    
    On Error GoTo LoadSoundMapInfo_Err
    
    Dim i       As Integer
    Dim j       As Integer
    Dim Temps   As String
    Dim MAPFILE As String
   
    MAPFILE = App.path & "\Mapas\" & "MAPA"
   
    For i = 1 To NumMaps
        Temps = GetVar(MAPFILE & i & ".dat", "SONIDOS", "Cantidad")
       
        If IsNumeric(Temps) Then
            p_Mapas(i).cantidad = Val(Temps)
           
            ReDim p_Mapas(i).Flags(1 To p_Mapas(i).cantidad) As Long
            ReDim p_Mapas(i).Probabilidad(1 To p_Mapas(i).cantidad) As Single
            ReDim p_Mapas(i).SoundIndex(1 To p_Mapas(i).cantidad) As Integer
           
            For j = 1 To p_Mapas(i).cantidad
                p_Mapas(i).Flags(j) = Val(GetVar(MAPFILE & i & ".dat", "SONIDO" & j, "Flags"))
                p_Mapas(i).Probabilidad(j) = Val(GetVar(MAPFILE & i & ".dat", "SONIDO" & j, "Probabilidad"))
                p_Mapas(i).SoundIndex(j) = Val(GetVar(MAPFILE & i & ".dat", "SONIDO" & j, "Sonido"))
            Next j

        Else
            p_Mapas(i).cantidad = 0

        End If

    Next i

    
    Exit Sub

LoadSoundMapInfo_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsSoundMapas" & "->" & "LoadSoundMapInfo"
    End If
Resume Next
    
End Sub
 
Public Sub ReproducirSonidosDeMapas()
    'Nos fijamos el mapa del personaje, sus sonidos y lo aplicamos al mapa. Simple y óptimo
    
    On Error GoTo ReproducirSonidosDeMapas_Err
    
    Dim i          As Long
    Dim SonidoMapa As Byte
    Dim posX       As Byte
    Dim posY       As Byte
   
    'Sounds are played at a random position
    posX = RandomNumber(UserPos.X, UserPos.X + 20)
    posY = RandomNumber(UserPos.Y, UserPos.Y + 20)
   
    i = UserMap
   
    If p_Mapas(i).cantidad > 0 Then
        SonidoMapa = RandomNumber(1, p_Mapas(i).cantidad)

        If RandomNumber(1, 100) <= p_Mapas(i).Probabilidad(SonidoMapa) Then

            'tocarlo
            If bRain Then
                If p_Mapas(i).Flags(SonidoMapa) Xor p_eSoundFlags.Lluvia Then
                       
                    Call Audio.PlayWave(p_Mapas(i).SoundIndex(SonidoMapa) & ".wav", posX, posY)

                End If

            Else

                If p_Mapas(i).Flags(SonidoMapa) Xor p_eSoundFlags.ninguna Then
                       
                    Call Audio.PlayWave(p_Mapas(i).SoundIndex(SonidoMapa) & ".wav", posX, posY)

                End If

            End If

        End If

    End If

    
    Exit Sub

ReproducirSonidosDeMapas_Err:
    If Err.number <> 0 Then
        LogError Err.number, Err.Description, "clsSoundMapas" & "->" & "ReproducirSonidosDeMapas"
    End If
Resume Next
    
End Sub
