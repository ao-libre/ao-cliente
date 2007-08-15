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

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

'16/03/2004 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - First Release
'
'03/03/2007 - Juan Martín Sotuyo Dodero (Maraxus) - (juansotuyo@gmail.com)
'   - Add: Sound and Music activation / deactivation from within the class.
'   - Add: Documentation
'   - Change: Volume is no longer the dB amplification / atenuation to be
'           applied but a percentage for easier interaction with user.


''
' Controls all audio using DirectSound and DirectMusic 7.
'
' @file     clsAudio.cls
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.1.0
' @date     20070303

Option Explicit

''
' Maximum number of sound buffers to hold in memory at any time.
Private Const NumSoundBuffers As Byte = 30

''
'Relative sound source, used to make fake 3d calculations
'
'@param spNone Sound is not 3D
'@param spUp Sound Source is above the user
'@param spDown Sound Source is below the user
'@param spLeft Souns Source is to the left of the user
'@param spRight Souns Source is to the right of the user
'@param spRightDown Souns Source is to the right of the user and below it
'@param spRightUp Souns Source is to the right of the user and above it
'@param spLeftDown Souns Source is to the left of the user and below it
'@param spLeftUp Souns Source is to the left of the user and above it

Public Enum eSoundPos
    spNone = 0
    spUp = 1
    spDown = 2
    spLeft = 3
    spRight = 4
    spRightDown = 5
    spRightUp = 6
    spLeftDown = 7
    spLeftUp = 8
End Enum

''
' Custom sound buffer structure.
'
' @param    FileName    The name of the file loaded in this buffer
' @param    looping     Set to true if the sound will loop until stopped
' @param    X           The X position of the sound in map coordinates
' @param    Y           The Y position of the sound in map coordinates
' @param    SoundSrc    Relative sound source for the character
' @param    Buffer      The DirectSoundBuffer object to wich the file is loaded
Private Type SoundBuffer
    FileName As String
    looping As Boolean
    x As Byte
    y As Byte
    SoundSrc As eSoundPos
    Buffer As DirectSoundBuffer
End Type

''
' States how to set a sound's looping state.
'
' @param    Default     If wave is in memory use it's looping setting, otherwise false
' @param    Disabled    Never loop
' @param    Enabled     Loop until stopped
Public Enum LoopStyle
    Default = 0
    Disabled = 1
    Enabled = 2
End Enum

''
' Set to true once the object is created to prevent multiple initializations of the same object.
Dim inited As Boolean

''
'The one and only DirectSound object
Dim DirectSound As DirectSound

''
'Array of all existing sound buffers
Dim DSBuffers(1 To NumSoundBuffers) As SoundBuffer

''
'Directory from which to load sound files by default
Dim SoundPath As String

''
'The volume for sounds (FXs)
Dim SndVolume As Long

''
'DirectMusic's Performance object
Dim Performance As DirectMusicPerformance

''
'Currently loaded segment
Dim Segment As DirectMusicSegment

''
'The one and only DirectMusic Loader
Dim Loader As DirectMusicLoader

''
'State of the currently loaded segment
Dim SegState As DirectMusicSegmentState

''
'States wether music is supported by hardware or not
Dim Music As Boolean

''
'Directory from which to load music files by default
Dim MusicPath As String

''
'Last loaded music file. Used to be able to load it when music gets activated
Dim CurrentMusicFile As String

''
'States wether music is currently activated or not
Dim mActivated As Boolean

''
'States wether sound is currently activated or not
Dim sActivated As Boolean

'%%%%FAKE 3D SOUND STUFF%%%%

''
'This indicates that the sound is not 3D
Private Const NO_3D_SOUND As Byte = 255

''
'Sound Reduction Offset Per Tile
'NOTE: Leave Constant type as it is to prevent overflows
Private Const PER_TILE_REDUCTION As Long = 25

''
'Panning Offset Per Tile
Private Const PER_TILE_PANNING As Long = 150

''
'Used to balance volume calculation.
'Big numbers allow to increase volume.
Private Const DIVIDE_MINVOLUME As Byte = 2

'%%%%FAKE 3D SOUND STUFF%%%%

''
'Constructor. Performs no action at all.

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'
'***************************************************
End Sub

''
'Destructor. Releases all created objects assuring no memory-leaks.

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
        Call DSBuffers(loopc).Buffer.Stop
        Set DSBuffers(loopc).Buffer = Nothing
    Next loopc
    
    'Stop music
    Call Performance.Stop(Segment, SegState, 0, 0)
    
    Call Performance.CloseDown
    
    Set DirectSound = Nothing
    
    Set Segment = Nothing
    Set Performance = Nothing
    Set Loader = Nothing
    
    inited = False
End Sub

''
'Initializes the clsAudio object.
'
' @param    dX          DirectX7 object to be used when creating the DirectSound and Directmusic objects.
' @param    hWnd        Handle to the window to be used by DirectMusic when playing midi files.
' @param    sound_path  Directory in wich to search for sound files.
' @param    music_path  Directory in wich to search for music files.
'
' @return   True if no errors occured, false otherwise.

Public Function Initialize(ByRef dX As DirectX7, ByVal hWnd As Long, ByVal sound_path As String, ByVal music_path As String) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Creates all DS and DM objects
'***************************************************
On Error GoTo ErrHandler
    
    ' Prevent multiple creations
    If inited Then Exit Function
    
    'Validate DX object
    If dX Is Nothing Then GoTo ErrHandler
    
    'Initialize DirectSound
    Set DirectSound = dX.DirectSoundCreate("")
    
    Call DirectSound.SetCooperativeLevel(hWnd, DSSCL_PRIORITY)
    
    'Set sound path
    SoundPath = sound_path
    
    'Set default volume
    SndVolume = 0
    
    'Initialize DirectMusic
    Music = InitializedirectMusic(dX, hWnd)
    
    'Set music path
    MusicPath = music_path
    
    mActivated = True
    sActivated = True
    
    CurrentMusicFile = vbNullString
    
    'Set tempo to 0 and volume of music
    MusicTempo = 0
    Call Performance.SetMasterVolume(200)
    
    inited = True
    
    Initialize = True
Exit Function

ErrHandler:
End Function

''
' Creates and configures all DirectMusic objects.
'
' @param    dX          DirectX7 object to be used when creating the DirectSound and Directmusic objects.
' @param    hWnd        Handle to the window to be used by DirectMusic when playing midi files.
'
' @return   True if no errors occured, false otherwise.

Private Function InitializedirectMusic(ByRef dX As DirectX7, ByVal hWnd As Long) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Initialize DM. Returns false if music isn't supported by HW
'***************************************************
On Error GoTo ErrHandler
    Set Loader = DirectX.DirectMusicLoaderCreate()
    
    Set Performance = dX.DirectMusicPerformanceCreate()
    Call Performance.Init(DirectSound, hWnd)
    Call Performance.SetMasterAutoDownload(True)        'Enable auto download of instruments
    
    'Set to port default, 1 channel group is enough for MIDIs
    Call Performance.SetPort(-1, 1)
    
    InitializedirectMusic = True
Exit Function

ErrHandler:
End Function

''
' Plays a new WAVE file.
'
' @param    FileName    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
' @param    X           The X position of the sound source, in map coordinates.
' @param    Y           The Y position of the sound source, in map coordinates.
' @param    LoopSound   The loop state for the sound.
'
' @return   The buffer index in which the sound is playing, or 0 in case of error.

Public Function PlayWave(ByVal FileName As String, ByVal x As Byte, ByVal y As Byte, ByVal SoundSrc As eSoundPos, Optional ByVal LoopSound As LoopStyle = LoopStyle.Default) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified by: Rapsodius
'New! Added Fake 3D Sound feature (using panning and volume)
'***************************************************
On Error GoTo ErrHandler
    Dim bufferIndex As Long
    
    If Not SoundActivated Then Exit Function
    
    'Get the buffer index were wave was loaded
    bufferIndex = LoadWave(FileName, LoopSound)
    If bufferIndex = 0 Then Exit Function   'If an error ocurred abort
    
    If x = NO_3D_SOUND Then
        With DSBuffers(bufferIndex)
            If .looping Then
                .Buffer.Play DSBPLAY_LOOPING
            Else
                .Buffer.Play DSBPLAY_DEFAULT
            End If
            
            'Apply volume
            Call .Buffer.SetVolume(SndVolume)
        End With
        
        PlayWave = bufferIndex
    Else
        If y <> NO_3D_SOUND Then
            Dim tmpX As Integer
            Dim tmpY As Integer
            Dim calcVol As Long
            Dim calcPanX As Long
            Dim calcPan As Integer
            
            DSBuffers(bufferIndex).x = x
            DSBuffers(bufferIndex).y = y
             
            tmpX = x
            tmpY = y
             
            Select Case SoundSrc
                Case eSoundPos.spUp
                    tmpY = -tmpY
                Case eSoundPos.spLeft
                    tmpX = -tmpX
                Case eSoundPos.spLeftUp
                    tmpX = -tmpX
                    tmpY = -tmpY
                Case eSoundPos.spRightUp
                    tmpY = -tmpY
                Case eSoundPos.spLeftDown
                    tmpX = -tmpX
            End Select
            
            If Abs(tmpX) <= 8 Then
                If Abs(tmpY) <= 6 Then
            
                    With DSBuffers(bufferIndex)
                        If .looping Then
                            .Buffer.Play DSBPLAY_LOOPING
                        Else
                            .Buffer.Play DSBPLAY_DEFAULT
                        End If
                        
                        calcVol = ((((Abs(tmpX) + Abs(tmpY)) * (DSBVOLUME_MIN \ DIVIDE_MINVOLUME)) + (14 - (Abs(tmpX) + Abs(tmpY)) * SndVolume)) \ 14 - ((Abs(tmpX) + Abs(tmpY)) * PER_TILE_REDUCTION))
                        
                        If calcVol < DSBVOLUME_MIN Then
                            calcVol = DSBVOLUME_MIN
                        End If
                        If calcVol > SndVolume Then
                            calcVol = SndVolume
                        End If
                            
                        If tmpX = 0 Then
                            calcPanX = DSBPAN_CENTER 'Avoid unnecesary calculation
                            GoTo AvoidCalcs:
                        End If
                        
                        calcPanX = ((((Abs(tmpX) * DSBPAN_RIGHT) + (8 - Abs(tmpX)) * DSBPAN_CENTER) \ 8) - (Abs(tmpX) * PER_TILE_PANNING))
                        
                        If tmpX >= 0 Then
                            calcPanX = -calcPanX
                        End If
AvoidCalcs:
                        calcPan = calcPanX \ (((2 * tmpY) + (6 * (6 - tmpY))) \ 6)
                        If calcPan < DSBPAN_LEFT Then
                            calcPan = DSBPAN_LEFT
                        End If
                        If calcPan > DSBPAN_RIGHT Then
                            calcPan = DSBPAN_RIGHT
                        End If
                        .Buffer.SetVolume calcVol
                        .Buffer.SetPan calcPan
                    End With
                    
                    PlayWave = bufferIndex
                End If
            End If
        End If
    End If
Exit Function

ErrHandler:
End Function

''
'Recalculates fake 3D effect for a specific sound or to all
'
' @param    bufferIndex Indicates which sound buffer is to be refreshed
' @param    RefreshAll  Whether to refresh all sounds or only one, if set to true, ignores the bufferIndex parameter

Public Sub RefreshWave(ByVal bufferIndex As Byte, ByVal RefreshAll As Boolean)
'**************************************************************
' RefreshWave - Recalcs panning and volume for a specific or to
' all sounds
'
' Creation Date: 08/09/07
' Last Modification Date: 08/14/07
' New volume calculation formula, bufferIndex out of range
' is checked only if RefreshAll is False
'**************************************************************
On Error Resume Next
    Dim tmpX As Integer
    Dim tmpY As Integer
    Dim calcVol As Long
    Dim calcPanX As Long
    Dim calcPan As Integer
    
    Dim i As Integer
    
    If RefreshAll Then
        For i = 1 To NumSoundBuffers
            With DSBuffers(i)
                If .x > NO_3D_SOUND Then
                    If .y > NO_3D_SOUND Then
                        tmpX = .x
                        tmpY = .y
                         
                        Select Case .SoundSrc
                            Case eSoundPos.spUp
                                tmpY = -tmpY
                            Case eSoundPos.spLeft
                                tmpX = -tmpX
                            Case eSoundPos.spLeftUp
                                tmpX = -tmpX
                                tmpY = -tmpY
                            Case eSoundPos.spRightUp
                                tmpY = -tmpY
                            Case eSoundPos.spLeftDown
                                tmpX = -tmpX
                        End Select
                        
                        If Abs(tmpX) <= 8 Then
                            If Abs(tmpY) <= 6 Then
                                calcVol = ((((Abs(tmpX) + Abs(tmpY)) * (DSBVOLUME_MIN \ DIVIDE_MINVOLUME)) + (14 - (Abs(tmpX) + Abs(tmpY)) * SndVolume)) \ 14 - ((Abs(tmpX) + Abs(tmpY)) * PER_TILE_REDUCTION))
                                
                                If calcVol < DSBVOLUME_MIN Then
                                    calcVol = DSBVOLUME_MIN
                                End If
                                If calcVol > SndVolume Then
                                    calcVol = SndVolume
                                End If
                                    
                                If tmpX = 0 Then
                                    calcPanX = DSBPAN_CENTER 'Avoid unnecesary calculation
                                    GoTo SkipCalcsLoop:
                                End If
                                
                                calcPanX = ((((Abs(tmpX) * DSBPAN_RIGHT) + (8 - Abs(tmpX)) * DSBPAN_CENTER) \ 8) - (Abs(tmpX) * PER_TILE_PANNING))
                                
                                If tmpX >= 0 Then
                                    calcPanX = -calcPanX
                                End If
SkipCalcsLoop:
                                calcPan = calcPanX \ (((2 * tmpY) + (6 * (6 - tmpY))) \ 6)
                                If calcPan < DSBPAN_LEFT Then
                                    calcPan = DSBPAN_LEFT
                                End If
                                If calcPan > DSBPAN_RIGHT Then
                                    calcPan = DSBPAN_RIGHT
                                End If
                                .Buffer.SetVolume calcVol
                                .Buffer.SetPan calcPan
                            End If
                        End If
                    End If
                End If
            End With
        Next i
    Else
        If bufferIndex <= 0 Or bufferIndex > NumSoundBuffers Then Exit Sub
        With DSBuffers(bufferIndex)
            If .x > NO_3D_SOUND Then
                If .y > NO_3D_SOUND Then
                    tmpX = .x
                    tmpY = .y
                     
                    Select Case .SoundSrc
                        Case eSoundPos.spUp
                            tmpY = -tmpY
                        Case eSoundPos.spLeft
                            tmpX = -tmpX
                        Case eSoundPos.spLeftUp
                            tmpX = -tmpX
                            tmpY = -tmpY
                        Case eSoundPos.spRightUp
                            tmpY = -tmpY
                        Case eSoundPos.spLeftDown
                            tmpX = -tmpX
                    End Select
                    
                    If Abs(tmpX) <= 8 Then
                        If Abs(tmpY) <= 6 Then
                            calcVol = ((((Abs(tmpX) + Abs(tmpY)) * (DSBVOLUME_MIN \ DIVIDE_MINVOLUME)) + (14 - (Abs(tmpX) + Abs(tmpY)) * SndVolume)) \ 14 - ((Abs(tmpX) + Abs(tmpY)) * PER_TILE_REDUCTION))
                            
                            If calcVol < DSBVOLUME_MIN Then
                                calcVol = DSBVOLUME_MIN
                            End If
                            If calcVol > SndVolume Then
                                calcVol = SndVolume
                            End If
                                
                            If tmpX = 0 Then
                                calcPanX = DSBPAN_CENTER 'Avoid unnecesary calculation
                                GoTo SkipCalcs:
                            End If
                            
                            calcPanX = ((((Abs(tmpX) * DSBPAN_RIGHT) + (8 - Abs(tmpX)) * DSBPAN_CENTER) \ 8) - (Abs(tmpX) * PER_TILE_PANNING))
                            
                            If tmpX >= 0 Then
                                calcPanX = -calcPanX
                            End If
SkipCalcs:
                            calcPan = calcPanX \ (((2 * tmpY) + (6 * (6 - tmpY))) \ 6)
                            If calcPan < DSBPAN_LEFT Then
                                calcPan = DSBPAN_LEFT
                            End If
                            If calcPan > DSBPAN_RIGHT Then
                                calcPan = DSBPAN_RIGHT
                            End If
                            .Buffer.SetVolume calcVol
                            .Buffer.SetPan calcPan
                        End If
                    End If
                End If
            End If
        End With
    End If
End Sub

''
' Loads a new WAVE file.
'
' @param    FileName    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
' @param    looping     The loop state for the sound.
'
' @return   The buffer index in which the sound is playing, or 0 in case of error.

Private Function LoadWave(ByVal FileName As String, ByVal looping As LoopStyle) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/04/06
'Loads a new wave into memory (if it's not already loaded) and returns
'the buffer index
'***************************************************
On Error GoTo ErrHandler
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    Dim i As Long
    
    FileName = UCase$(FileName)
    
    If Not FileExist(SoundPath & FileName, vbArchive) Then _
        Exit Function
    
    'Check if the buffer is in memory and not playing
    For i = 1 To NumSoundBuffers
        If DSBuffers(i).FileName = FileName Then
            If DSBuffers(i).Buffer.GetStatus = 0 Or DSBuffers(i).Buffer.GetStatus() = DSBSTATUS_BUFFERLOST Then
                'Found it!!! We just play this one :)
                LoadWave = i
                
                'Set looping if needed
                If looping <> LoopStyle.Default Then DSBuffers(i).looping = (looping = LoopStyle.Enabled)
                
                Exit Function
            End If
        End If
    Next i
    
    'Not in memory, search for an empty buffer
    For i = 1 To NumSoundBuffers
        If DSBuffers(i).Buffer Is Nothing Then
            Exit For
        End If
    Next i
    
    'No empty buffers?? Search for one not being played and use that slot
    If i > NumSoundBuffers Then
        For i = 1 To NumSoundBuffers
            If DSBuffers(i).Buffer.GetStatus() = DSBSTATUS_BUFFERLOST Or DSBuffers(i).Buffer.GetStatus() = 0 Then
                Exit For
            End If
        Next i
    End If
    
    'Still no free buffers?? Get the first non-looping sound
    If i > NumSoundBuffers Then
        For i = 1 To NumSoundBuffers
            If Not (DSBuffers(i).Buffer.GetStatus() And DSBSTATUS_LOOPING) Then
                Call DSBuffers(i).Buffer.Stop
                Exit For
            End If
        Next i
    End If
    
    'If every sound is looping, just ignore new sound unless it's looping too (in which case we set it at the first slot by default)
    If i > NumSoundBuffers Then
        If Not looping Then Exit Function
        i = 1
    End If
    
    'Create new sound
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    
    With DSBuffers(i)
        Set .Buffer = Nothing   'Get rid of any previous data
        
        .FileName = FileName
        .looping = (looping = LoopStyle.Enabled)
        
        Set .Buffer = DirectSound.CreateSoundBufferFromFile(SoundPath & FileName, bufferDesc, waveFormat)
    End With
    
    LoadWave = i
Exit Function

ErrHandler:
End Function

''
' Stops a given sound or all of them.
'
' @param    bufferIndex     The index of the sound file to stop playing. If 0, all sounds are stopped.

Public Sub StopWave(Optional ByVal bufferIndex As Long = 0)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Stops the sound in the given buffer index, or every sound if
'index is 0
'***************************************************
    If bufferIndex > 0 And bufferIndex <= NumSoundBuffers Then
        If DSBuffers(bufferIndex).Buffer.GetStatus() And DSBSTATUS_PLAYING Then
            Call DSBuffers(bufferIndex).Buffer.Stop
        End If
    ElseIf bufferIndex = 0 Then
        Dim i As Long
        For i = 1 To NumSoundBuffers
            If Not DSBuffers(i).Buffer Is Nothing Then
                If DSBuffers(i).Buffer.GetStatus() <> 0 And DSBuffers(i).Buffer.GetStatus() <> DSBSTATUS_BUFFERLOST Then
                    Call DSBuffers(i).Buffer.Stop
                End If
            End If
        Next i
    End If
End Sub

''
' Retrieves wether there are sounds currentyl playing or not.
'
' @return   true if there is at least one sound playing, false otherwise.

Public Property Get PlayingSound() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns true if there is at least one sound playing
'***************************************************
    Dim i As Long
    
    For i = 1 To NumSoundBuffers
        If DSBuffers(i).Buffer.GetStatus() And DSBSTATUS_PLAYING Then
            PlayingSound = True
            Exit Property
        End If
    Next i
End Property

''
' Sets the volume of sound.
'
' @param    Volume  New value to be set, in percentage (0 to 100).

Public Property Let SoundVolume(ByVal Volume As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the sound's volume as a percentage
'***************************************************
    Dim i As Long
    
    If Volume > 100 Then Exit Property
    
    'Take percentage to actual value
    SndVolume = Volume * (DSBVOLUME_MAX - DSBVOLUME_MIN) / 100 + DSBVOLUME_MIN
    
    For i = 1 To NumSoundBuffers
        If Not DSBuffers(i).Buffer Is Nothing Then
            If DSBuffers(i).Buffer.GetStatus And DSBSTATUS_PLAYING Then
                Call DSBuffers(i).Buffer.SetVolume(SndVolume)
            End If
        End If
    Next i
End Property

''
' Gets the volume of sound.
'
' @return   The current sound volume in percentage.

Public Property Get SoundVolume() As Byte
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the music's volume as a percentage
'***************************************************
    SoundVolume = CByte((SndVolume - DSBVOLUME_MIN) * 100 / (DSBVOLUME_MAX - DSBVOLUME_MIN))
End Property

''
' Plays a new MIDI file.
'
' @param    file    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
' @param    Loops   Number of times to loop sound. For eternal loops set it to -1.
'
' @return   True if file was properly loaded and started playing, False otherwise.

Public Function PlayMIDI(Optional ByVal file As String = "", Optional ByVal Loops As Long = -1) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Loads a midi and start to play it
'***************************************************
On Error GoTo ErrHandler
    If LenB(file) > 0 Then _
        CurrentMusicFile = file
    
    If Not MusicActivated Then Exit Function
    
    If Not Music Then Exit Function
    
    If PlayingMusic Then StopMidi
    
    If LenB(file) > 0 Then
        If Not LoadMidi(file) Then Exit Function
    Else
        'Make sure we have a loaded segment
        If Segment Is Nothing Then Exit Function
    End If
    
    'Play it
    Call Segment.SetRepeats(Loops)
    
    Set SegState = Performance.PlaySegment(Segment, 0, 0)
    
    PlayMIDI = True
Exit Function

ErrHandler:
End Function

''
' Loads a new MIDI file.
'
' @param    file    The path to the file to be loaded. Must be a relative path from the default path set at initialization.
'
' @return   True if file was properly loaded and started playing, False otherwise.

Private Function LoadMidi(ByVal file As String) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Loads a midi to memory
'***************************************************
On Error GoTo ErrHandler
    If Not FileExist(MusicPath & file, vbArchive) Then Exit Function
    
    Call StopMidi
    
    'Destroy old object
    Set Segment = Nothing
    
    Set Segment = Loader.LoadSegment(MusicPath & file)
    
    If Segment Is Nothing Then
        Exit Function
    End If
    
    Call Segment.SetStandardMidiFile
    
    LoadMidi = True
Exit Function

ErrHandler:
End Function

''
' Stops playing the currently loaded MIDI file.

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

''
' Sets the music's tempo.
'
' @param    Tempo   New Tempo value. Must be in the range of 0.25 to 2.0

Public Property Let MusicTempo(ByVal Tempo As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the music's tempo
'***************************************************
    If Not Music Then Exit Property
    
    If Tempo < 0.25 Or Tempo > 2# Then Exit Property
    
    Call Performance.SetMasterTempo(Tempo)
End Property

''
' Sets the music's volume.
'
' @param    Volume  Percentage to wich to set the music volume.

Public Property Let MusicVolume(ByVal Volume As Long)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the music's volume as a percentage
'***************************************************
    If Not Music Then Exit Property
    
    If Volume < 0 Or Volume > 100 Then Exit Property
    
    ' Volume ranges from -10000 to 10000
    Call Performance.SetMasterVolume(Volume * 200 - 10000)
End Property

''
' Gets the music's tempo.
'
' @return   Current music'tempo.

Public Property Get MusicTempo() As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns the music's tempo
'***************************************************
    If Not Music Then Exit Property
    
    MusicTempo = Performance.GetMasterTempo()
End Property

''
' Gets the music's volume.
'
' @return   Current music's volume as a percentage.

Public Property Get MusicVolume() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns the music's volume as a percentage
'***************************************************
    If Not Music Then Exit Property
    
    MusicVolume = (Performance.GetMasterVolume() + 10000) / 200
End Property

''
' Checks wether there is music currently playing.
'
' @return   True if there is a MIDI playing, false otherwise.

Public Property Get PlayingMusic() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns True if a midi is being played
'***************************************************
    If Not Music Or Segment Is Nothing Then Exit Property
    
    PlayingMusic = Performance.IsPlaying(Segment, SegState)
End Property

''
' Retrieves the music's length.
'
' @return   Current music's length.

Public Property Get MusicLenght() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns the lenght of the current midi
'***************************************************
    If Not Music Then Exit Property
    
    MusicLenght = Segment.GetLength()
End Property

''
' Retrieves wether music is activated or not.
'
' @return   True if music is activated, false otherwise.

Public Property Get MusicActivated() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/03/07
'Retrieves wether the music is activated or not
'***************************************************
    MusicActivated = mActivated
End Property

''
' Sets wether music is activated or not.
'
' @param    nValue  True if music should be activated, false otherwise.

Public Property Let MusicActivated(ByVal nValue As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/03/07
'Sets wether the music is activated or not
'***************************************************
    If mActivated = nValue Then Exit Property
    
    mActivated = nValue
    
    If nValue Then
        Call PlayMIDI(CurrentMusicFile)
    Else
        Call StopMidi
    End If
End Property

''
' Retrieves wether sound is activated or not.
'
' @return   True if sound is activated, false otherwise.

Public Property Get SoundActivated() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/03/07
'Retrieves wether the sound is activated or not
'***************************************************
    SoundActivated = sActivated
End Property

''
' Sets wether sound is activated or not.
'
' @param    nValue  True if sound should be activated, false otherwise.

Public Property Let SoundActivated(ByVal nValue As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/03/07
'Sets wether the sound is activated or not
'***************************************************
    Dim i As Long
    
    If sActivated = nValue Then Exit Property
    
    sActivated = nValue
    
    If Not nValue Then
        Call StopWave
    End If
End Property

''
'Retrieves the NO_3D_SOUND constant, in order to play normal sounds
'
'@return NO_3D_SOUND constant

Public Property Get No3DSound() As Byte
'**************************************************************
' Author: Rapsodius
' Creation Date: 08/14/07
' Retrieves the NO_3D_SOUND constant, in order to play normal sounds
'**************************************************************
    No3DSound = NO_3D_SOUND
End Property

