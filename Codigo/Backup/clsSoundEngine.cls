VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoundEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsSoundEngine - ImperiumAO - v1.3.0
'
'Sound Engine. Some functions are ORE1-Based, but modified.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

Option Explicit

Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Private mAmbienteActual As Integer
Private mLastAmbienteActual As Integer
Private mMusActual As Integer
Private WavPath As String
Private MidiPath As String
Private MP3Path As String

'Manejador de buffers
Private Buffers As CBufferMan

'El volúmen actual real (sonidos)
Private mEffectsVolume As Long

'El volúmen actual real (ambientes)
Private mEffectsVolumeAmbient As Long

'El volúmen actual real (música)
Private mMusicVolume As Long

'El volúmen máximo en la configuración (sonidos)
Private mEffectsVolumeMax As Long

'El volúmen máximo en la configuración (ambientes)
Private mAmbientVolumeMax As Long

'El volúmen máximo en la configuración (música)
Private mMusicVolumeMax As Long

'¿Estamos haciendo efecto fade?
Private mFadingStatus As Byte
Private mFadingMusicMod As Long

'¿Próximo MP3 o MIDI?
Private mNextMusic As Integer

Private INTERIOR_VOLUME As Long
Private EXTERIOR_VOLUME As Long

Private TimeElapsed As Long

Private Stream_Handle As Long

Private Sub Class_Terminate()
Engine_DeInitialize
End Sub

Private Property Let SetWavPath(ByVal inString As String)
WavPath = inString
End Property

Private Property Let SetMP3Path(ByVal inString As String)
MP3Path = inString
End Property

Private Property Let SetMidiPath(ByVal inString As String)
MidiPath = inString
End Property

Public Function Engine_DeInitialize()

'Destroy all buffers
Call Buffers.BorrarTodo
Set Buffers = Nothing

End Function

Public Function Initialize_Engine(ByVal main_hwnd As String, ByVal Set_Wav_Path As String, ByVal Set_MP3_Path As String, ByVal Set_Midi_Path As String, Optional ByVal Play_only_on_focus As Boolean = True, Optional ByVal Initialize_Sounds As Boolean = True, Optional ByVal Initialize_Music As Boolean = True, Optional ByVal sound_volume As Long, Optional ByVal music_volume As Long) As Boolean

On Error GoTo ErrHandler

If BASS_Init(1, 44100, 0, frmMain.hwnd, 0) = BASSFALSE Then GoTo ErrHandler

If Initialize_Sounds Then
    SetWavPath = Set_Wav_Path
    
    Set Buffers = New CBufferMan
    Buffers.MaxEntries = 300
    Buffers.Path = Set_Wav_Path
    Call Buffers.Init
    
    'If Play_only_on_focus Then
    '    DS.SetCooperativeLevel main_hwnd, DSSCL_NORMAL
    'Else
    '    DS.SetCooperativeLevel main_hwnd, DSSCL_PRIORITY
    'End If

    VolumenActual = sound_volume

End If

If Initialize_Music Then
    SetMidiPath = Set_Midi_Path
    SetMP3Path = Set_MP3_Path
End If

Initialize_Engine = True

Exit Function

ErrHandler:
    Initialize_Engine = False
End Function

Public Property Get AmbienteActual() As Integer
AmbienteActual = mAmbienteActual
End Property

Public Property Let AmbienteActual(ByVal vNewValue As Integer)

If Not Meteo_Engine.SecondaryStatus = 2 Then
    Call Sound_Stop(mAmbienteActual)
    mAmbienteActual = vNewValue
    Call Sound_Play(mAmbienteActual, True, mEffectsVolumeAmbient)
Else
    LastAmbienteActual = vNewValue
End If

End Property

Public Function Sound_Ambient_Play()
Call Sound_Play(mAmbienteActual, True, mEffectsVolumeAmbient)
End Function

Public Function Sound_Ambient_Stop()
Call Sound_Stop(mAmbienteActual)
End Function

Public Property Get LastAmbienteActual() As Integer
LastAmbienteActual = mLastAmbienteActual
End Property

Public Property Let LastAmbienteActual(ByVal vNewValue As Integer)
mLastAmbienteActual = vNewValue
End Property

Public Property Get MusicActual() As Integer
MusicActual = mMusActual
End Property

Public Property Let MusicActual(ByVal vNewValue As Integer)
mMusActual = vNewValue
End Property

Public Property Let Fading(ByVal vNewValue As Long)
mFadingMusicMod = vNewValue
TimeElapsed = GetTickCount
End Property

Public Property Get Fading() As Long
Fading = mFadingMusicMod
End Property

Public Property Get VolumenActual() As Long
VolumenActual = mEffectsVolumeMax
End Property

Public Property Let VolumenActual(ByVal vNewValue As Long)

If vNewValue > 0 Then vNewValue = 0
If vNewValue < -4000 Then vNewValue = -4000

mEffectsVolumeMax = vNewValue

mAmbientVolumeMax = CLng(vNewValue * 2)

EXTERIOR_VOLUME = mAmbientVolumeMax
INTERIOR_VOLUME = CLng(mAmbientVolumeMax / 1.5)

End Property

Public Property Let VolumenActualMusicMax(ByVal vNewValue As Long)
mMusicVolumeMax = vNewValue
End Property

Public Property Get VolumenActualMusicMax() As Long
VolumenActualMusicMax = mMusicVolumeMax
End Property

Public Property Get VolumenActualMusic() As Long
VolumenActualMusic = mMusicVolume
End Property

Public Property Let VolumenActualMusic(ByVal vNewValue As Long)

If vNewValue > 0 Then vNewValue = 0
If vNewValue < -4000 Then vNewValue = -4000

mMusicVolume = vNewValue

End Property

Public Property Let NextMusic(ByVal vNewValue As Integer)

If mFadingStatus <> 2 And vNewValue <> 0 Then 'Si no estamos en fade up ya...
    mNextMusic = vNewValue
Else
    mNextMusic = vNewValue
    mFadingStatus = 0
End If

End Property

Public Function Sound_Stop_All() As Boolean

Buffers.DetenerTodo
Sound_Stop_All = True

End Function

Public Function Sound_Play(ByVal Sound_File As Integer, Optional ByVal Do_Loop As Boolean = False, Optional ByVal volume As Long = 0) As Boolean

Dim Sound_Buffer As Long
Dim ch As Long

On Error GoTo Error_Handl
    
    If (Sound_File <= 0) Or (fx = 0) Then Exit Function
    
    Sound_Buffer = Buffers.GetBuffer(Sound_File)
    If Sound_Buffer = 0 Then Exit Function
    
    BASS_Start
    ch = BASS_SampleGetChannel(Sound_Buffer, BASSFALSE)
    Call BASS_ChannelSetAttributes(ch, -1, 50, Int((201 * Rnd) - 100))
    Debug.Print IIf(BASS_ChannelPlay(ch, IIf(Do_Loop, BASSTRUE, BASSFALSE)) = 0, BASS_ErrorGetCode, "OK")
        
    'If volume = 0 Then
    '    Call Sound_Buffer.SetVolume(mEffectsVolumeMax)
    'Else
    '    Call Sound_Buffer.SetVolume(volume)
    'End If
    
    'If Do_Loop Then
    '    Sound_Buffer.Play DSBPLAY_LOOPING
    'Else
    '    Sound_Buffer.Play DSBPLAY_DEFAULT
    'End If
    
    'Set Sound_Buffer = Nothing
    
    Sound_Play = True
    
    Exit Function
    
Error_Handl:
    Sound_Play = False
End Function

Public Function Sound_Stop(ByVal Sound_File As Integer) As Boolean

On Error GoTo Error_Handl
    If Sound_File <= 0 Then Exit Function
    'Buffers.GetBuffer(Sound_File).Stop
    'Buffers.GetBuffer(Sound_File).SetCurrentPosition 0
    
    Sound_Stop = True
    Exit Function
Error_Handl:
    Sound_Stop = False
End Function

Public Function Music_Empty() As Boolean
    
    On Error GoTo Error_Handl
    
    Call BASS_StreamFree(Stream_Handle)
    
    Music_Empty = True
    Exit Function

Error_Handl:
    Music_Empty = False
    End
End Function

Public Function Music_Load(ByVal file_number As Integer, Optional ByVal volume As Long = 0, Optional ByVal balance As Long = 0) As Boolean '**** Loads a MP3 *****
    
    On Error GoTo Error_Handl
        
    If Musica = CONST_MP3 Then
        If General_File_Exists(MP3Path & file_number & ".mp3", vbNormal) Then
            If Not Music_Empty() = True Then GoTo Error_Handl
            Stream_Handle = BASS_StreamCreateFile(BASSFALSE, MP3Path & file_number & ".mp3", 0, 0, 0)
        ElseIf Extract_File(Midi, MidiPath, file_number & ".mid", False) Then
            If Not Music_Empty() = True Then GoTo Error_Handl
            Stream_Handle = BASS_StreamCreateFile(BASSFALSE, Windows_Temp_Dir & file_number & ".mid", 0, 0, 0)
            Kill Windows_Temp_Dir & file_number & ".mid"
        Else
            GoTo Error_Handl
        End If
    ElseIf Musica = CONST_MIDI Then
        If Extract_File(Midi, MidiPath, file_number & ".mid", False) Then
            If Not Music_Empty() = True Then GoTo Error_Handl
            Stream_Handle = BASS_StreamCreateFile(BASSFALSE, Windows_Temp_Dir & file_number & ".mid", 0, 0, 0)
            Kill Windows_Temp_Dir & file_number & ".mid"
        End If
    Else
        GoTo Error_Handl
    End If

    mMusActual = file_number
    Music_Load = True
    Exit Function

Error_Handl:
    Music_Load = False
End Function

Public Function Music_Play() As Boolean
    
    On Error GoTo Error_Handl
    
    BASS_Start
    If BASS_ChannelPlay(Stream_Handle, BASSFALSE) = 0 Then Debug.Print BASS_ErrorGetCode
    
    Music_Play = True
    Exit Function

Error_Handl:
    Music_Play = False
End Function

Public Function Music_GetLoop() As Boolean
    On Error GoTo Error_Handl
    'If IMPos Is Nothing Then Exit Function
    'If IMPos.StopTime = IMPos.CurrentPosition Then
    '    Music_Stop
    '    Music_Play
    'End If
Error_Handl:
End Function

Public Function Music_Stop() As Boolean

    On Error GoTo Error_Handl
    
    'IMC.Stop
    'IMPos.CurrentPosition = 0
    
    Music_Stop = True
    Exit Function

Error_Handl:
    Music_Stop = False
End Function

Public Function Music_Pause() As Boolean

    On Error GoTo Error_Handl
    
    IMC.Pause
    
    Music_Pause = True
    Exit Function

Error_Handl:
    Music_Pause = False
End Function

Public Function Music_Volume_Set(ByVal sound_volume As Long) As Boolean

    On Error GoTo Error_Handl
    
    'If IBA Is Nothing Then Exit Function
    
    If sound_volume > 0 Then sound_volume = 0
    If sound_volume < -4000 Then sound_volume = -4000
    
    'IBA.volume = sound_volume

    Music_Volume_Set = True
    Exit Function

Error_Handl:
    Music_Volume_Set = False
End Function

Public Function Sound_Render()

On Error Resume Next

Dim TiempoActual As Long

Static under_roof As Boolean
Static Sound_Buffer As DirectSoundSecondaryBuffer8

    If CurrentUser.MapNum > 0 Then
        under_roof = Engine.Char_User_Roof_Get
        If CurrentUser.MapExt Then
            If Meteo_Engine.SecondaryStatus = 2 Then
                If under_roof Then
                   If mAmbienteActual <> SND_LLUVIAIN Then
                        
                        If (mAmbienteActual <> SND_LLUVIAIN) And (mAmbienteActual <> SND_LLUVIAOUT) Then
                            LastAmbienteActual = mAmbienteActual
                        End If
                                            
                        Call Sound_Stop(mAmbienteActual)
                        mEffectsVolumeAmbient = INTERIOR_VOLUME
                        mAmbienteActual = SND_LLUVIAIN
                        Call Sound_Play(SND_LLUVIAIN, True, mEffectsVolumeAmbient)
                   End If
                Else
                   If mAmbienteActual <> SND_LLUVIAOUT Then
                        
                        If (mAmbienteActual <> SND_LLUVIAIN) And (mAmbienteActual <> SND_LLUVIAOUT) Then
                            LastAmbienteActual = mAmbienteActual
                        End If
                        
                        Call Sound_Stop(mAmbienteActual)
                        mEffectsVolumeAmbient = EXTERIOR_VOLUME
                        mAmbienteActual = SND_LLUVIAOUT
                        Call Sound_Play(SND_LLUVIAOUT, True, mEffectsVolumeAmbient)
                   End If
                End If
            Else 'Acá se deberían hacer más checkeos para nublado, etc.
                If LastAmbienteActual > 0 Then
                    Call Sound_Stop(mAmbienteActual)
                    mAmbienteActual = LastAmbienteActual
                    Call Sound_Play(mAmbienteActual, True, mEffectsVolumeAmbient)
                    LastAmbienteActual = 0
                End If
                
                If under_roof Then 'En los interiores el sonido es más suave :)
                    If mEffectsVolumeAmbient <> INTERIOR_VOLUME Then
                        
                        'Set Sound_Buffer = Buffers.GetBuffer(mAmbienteActual)
                        
                        If (Not Sound_Buffer Is Nothing) Then
                            Sound_Buffer.SetVolume (INTERIOR_VOLUME)
                            mEffectsVolumeAmbient = INTERIOR_VOLUME
                        Else
                            If mAmbienteActual > 0 Then
                                mEffectsVolumeAmbient = INTERIOR_VOLUME
                                Call Sound_Play(mAmbienteActual, True, mEffectsVolumeAmbient)
                            End If
                        End If
                    End If
                Else
                    If mEffectsVolumeAmbient <> EXTERIOR_VOLUME Then
                        
                        'Set Sound_Buffer = Buffers.GetBuffer(mAmbienteActual)
                        
                        If (Not Sound_Buffer Is Nothing) Then
                            Sound_Buffer.SetVolume (EXTERIOR_VOLUME)
                            mEffectsVolumeAmbient = EXTERIOR_VOLUME
                        Else
                            If mAmbienteActual > 0 Then
                                mEffectsVolumeAmbient = EXTERIOR_VOLUME
                                Call Sound_Play(mAmbienteActual, True, mEffectsVolumeAmbient)
                            End If
                        End If
                    
                    End If
                End If
            End If
        Else
            If mEffectsVolumeAmbient <> INTERIOR_VOLUME Then
                
                'Set Sound_Buffer = Buffers.GetBuffer(mAmbienteActual)
                
                If Not Sound_Buffer Is Nothing Then
                    Sound_Buffer.SetVolume (INTERIOR_VOLUME)
                    mEffectsVolumeAmbient = INTERIOR_VOLUME
                Else
                    If mAmbienteActual > 0 Then
                        mEffectsVolumeAmbient = INTERIOR_VOLUME
                        Call Sound_Play(mAmbienteActual, True, mEffectsVolumeAmbient)
                    End If
                End If
                
            End If
        End If
    End If
    
    'Agregar FadeStatus para saber si terminar o no...
    If mFadingMusicMod <> 0 And Musica <> CONST_DESHABILITADA Then
        
        TiempoActual = GetTickCount
        
        '¿Fade up or fade down? Primero fade down, después fade up
        If TimeElapsed + 60 < TiempoActual Then
            
            TimeElapsed = TiempoActual
            
            'Fade down
            If mMusicVolume > mMusicVolumeMax Then
                
                mMusicVolume = mMusicVolumeMax
                
                If mFadingStatus = 0 Then
                    mFadingStatus = 1
                    mFadingMusicMod = mFadingMusicMod * (-1)
                Else
                    'TERMINAMOS!!!
                    mFadingMusicMod = 0
                    mFadingStatus = 0
                End If
            'Fade up
            ElseIf mMusicVolume < -4000 Then
                mFadingStatus = 2
                mMusicVolume = -4000
                mFadingMusicMod = mFadingMusicMod * (-1)
            End If
            
            mMusicVolume = mMusicVolume + mFadingMusicMod
            
            If mNextMusic > 0 And mFadingStatus = 2 Then
                If Music_Load(mNextMusic, mMusicVolume) Then
                    Music_Stop
                    Music_Play
                End If
                mNextMusic = 0
            Else
                Music_Volume_Set mMusicVolume
            End If
        End If
    Else
        mMusicVolume = mMusicVolumeMax
    End If

If Musica = CONST_MP3 Then Call Music_GetLoop

End Function

Public Function Sound_Load(ByVal file_num As Integer) As Boolean
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 7/16/2005
'**************************************************************
    
    If Buffers.CargarBuffer(file_num) Then
        Sound_Load = True
    End If

End Function

Public Sub BorraTimer()
Buffers.BorraTimerProc
End Sub
