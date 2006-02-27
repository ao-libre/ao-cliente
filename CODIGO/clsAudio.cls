VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsAudio.cls - Controls all audio using DirectSound and DirectMusic
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@hotmail.com)
'**************************************************************

Option Explicit

Private Const NumSoundBuffers As Byte = 7

'WAVE Sounds
Dim DirectSound As DirectSound
Dim DSBuffers(1 To NumSoundBuffers) As DirectSoundBuffer
Dim LastSoundBufferUsed As Integer

Dim SoundPlaying As Boolean

Dim SoundPath As String

'Music
Dim Performance As DirectMusicPerformance
Dim Segment As DirectMusicSegment
Dim Loader As DirectMusicLoader
Dim SegState As DirectMusicSegmentState

Dim Music As Boolean        'True if music is supported by HW

Dim MusicPath As String

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'
'***************************************************


End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Destroy's all DirectMusic and DirectSound objects
'***************************************************
On Local Error Resume Next
    Dim loopc As Long
    
    'Stop every channel being used and destroy the buffer
    For loopc = 1 To NumSoundBuffers
        Call DSBuffers(loopc).Stop
        Set DSBuffers(loopc) = Nothing
    Next loopc
    
    'Stop music
    Call Performance.Stop(Segment, SegState, 0, 0)
    
    Set DirectSound = Nothing
    
    Set Segment = Nothing
    Set Performance = Nothing
    Set Loader = Nothing
End Sub

Public Function Initialize(ByRef DX As DirectX7, ByVal hWnd As Long, ByVal sound_path As String, ByVal music_path As String) As Boolean
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Creates all DS and DM objects
'***************************************************
On Error GoTo ErrHandler

    'Validate DX object
    If DX Is Nothing Then GoTo ErrHandler
    
    'Initialize DirectSound
    Set DirectSound = DX.DirectSoundCreate("")
    
    DirectSound.SetCooperativeLevel hWnd, DSSCL_PRIORITY
    LastSoundBufferUsed = 0
    
    'Set sound path
    SoundPath = sound_path
    
    'Initialize DirectMusic
    Music = InitializedirectMusic(DX, hWnd)
    
    'Set music path
    MusicPath = music_path
    
    Initialize = True
Exit Function

ErrHandler:
End Function

Private Function InitializedirectMusic(ByRef DX As DirectX7, ByVal hWnd As Long) As Boolean
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Initialize DM. Returns false if music isn't supported by HW
'***************************************************
On Error GoTo ErrHandler
    Set Loader = DirectX.DirectMusicLoaderCreate()
    
    Set Performance = DX.DirectMusicPerformanceCreate()
    Performance.Init DirectSound, hWnd
    Performance.SetMasterAutoDownload True     'Enable auto doanload of instruments
    
    'Set port to default, 1 channel group is enough for MIDIs
    Performance.SetPort -1, 1
    
    InitializedirectMusic = True
Exit Function

ErrHandler:
End Function

Public Function PlayWave(Optional ByVal FileName As String = "", Optional ByVal LoopSound As Boolean = False) As Boolean
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Loads a new wave to memory and starts playing it
'***************************************************
On Error GoTo ErrHandler
    If FileName <> "" Then
        If Not LoadWave(FileName) Then Exit Function
    Else
        If DSBuffers(LastSoundBufferUsed) Is Nothing Then Exit Function
    End If
    
    If LoopSound Then
        DSBuffers(LastSoundBufferUsed).Play DSBPLAY_LOOPING
    Else
        DSBuffers(LastSoundBufferUsed).Play DSBPLAY_DEFAULT
    End If
    SoundPlaying = True
    
    PlayWave = True
Exit Function

ErrHandler:
End Function

Private Function LoadWave(ByVal FileName As String) As Boolean
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'loads a new wave into memory
'***************************************************
On Error GoTo ErrHandler
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    
    If Not FileExist(SoundPath & FileName, vbArchive) Then _
        Exit Function
    
    LastSoundBufferUsed = LastSoundBufferUsed + 1
    If LastSoundBufferUsed > NumSoundBuffers Then
        LastSoundBufferUsed = 1
    End If
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    
    Set DSBuffers(LastSoundBufferUsed) = DirectSound.CreateSoundBufferFromFile(SoundPath & FileName, bufferDesc, waveFormat)
    
    LoadWave = True
Exit Function

ErrHandler:
End Function

Public Sub StopWave()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Stops the current sound
'***************************************************
    If SoundPlaying Then DSBuffers(LastSoundBufferUsed).Stop
End Sub

Public Property Get PlayingSound() As Boolean
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'
'***************************************************
    PlayingSound = SoundPlaying
End Property

Public Property Let SoundVolume(ByVal Volume As Long)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the sound's volume
'***************************************************
    If SoundPlaying Then DSBuffers(LastSoundBufferUsed).SetVolume Volume
End Property

Public Property Get SoundVolume() As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the music's volume
'***************************************************
    If SoundPlaying Then SoundVolume = DSBuffers(LastSoundBufferUsed).GetVolume()
End Property

Public Function PlayMIDI(Optional ByVal file As String = "", Optional ByVal Loops As Long = -1) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Loads a midi and start to play it
'***************************************************
On Error GoTo ErrHandler
    If Not Music Then Exit Function
    
    If PlayingMusic Then StopMidi
    
    If file <> "" Then
        If Not LoadMidi(file) Then Exit Function
    Else
        'Make sure we have a loaded segment
        If Segment Is Nothing Then Exit Function
    End If
    
    'Play it
    Segment.SetRepeats Loops
    
    Set SegState = Performance.PlaySegment(Segment, 0, 0)
    
    'Set tempo to 0 and volume to max
    MusicTempo = 0
    MusicVolume = 200
    
    PlayMIDI = True
Exit Function

ErrHandler:
End Function

Private Function LoadMidi(ByVal file As String) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Loads a midi to memory
'***************************************************
On Error GoTo ErrHandler
    If Not FileExist(MusicPath & file, vbArchive) Then Exit Function
    
    If PlayingMusic Then StopMidi
    
    Set Segment = Loader.LoadSegment(MusicPath & file)
    
    If Segment Is Nothing Then
        Exit Function
    End If
    
    Segment.SetStandardMidiFile
    
    LoadMidi = True
Exit Function

ErrHandler:
End Function

Public Sub StopMidi()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Stops playing a midi
'***************************************************
On Error GoTo ErrHandler
    If PlayingMusic Then
        Call Performance.Stop(Segment, SegState, 0, 0)
    End If
Exit Sub

ErrHandler:
End Sub

Public Property Let MusicTempo(ByVal Tempo As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the music's tempo
'***************************************************
    If Not Music Then Exit Property
    
    Performance.SetMasterTempo Tempo
End Property

Public Property Let MusicVolume(ByVal Volume As Long)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the music's volume
'***************************************************
    If Not Music Then Exit Property
    
    Performance.SetMasterVolume Volume
End Property

Public Property Get MusicTempo() As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns the music's tempo
'***************************************************
    If Not Music Then Exit Property
    
    MusicTempo = Performance.GetMasterTempo
End Property

Public Property Get MusicVolume() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns the music's volume
'***************************************************
    If Not Music Then Exit Property
    
    MusicVolume = Performance.GetMasterVolume
End Property

Public Property Get PlayingMusic() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns True if a midi is being played
'***************************************************
    If Not Music Or Segment Is Nothing Then Exit Property
    
    PlayingMusic = Performance.IsPlaying(Segment, SegState)
End Property

Public Property Get MusicLenght() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns the lenght of the current midi
'***************************************************
    If Not Music Then Exit Property
    
    MusicLenght = Segment.GetLength
End Property
