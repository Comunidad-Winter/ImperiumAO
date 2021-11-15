VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIAO_SE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private Type SND_DB_ENTRY
    FileName As Integer
    Accesos As Long
    Cont As Long
End Type

Private mSNDDB() As SND_DB_ENTRY
Private mBufferDB() As DirectSoundSecondaryBuffer8
Private mKeys() As Integer

Private DX As DirectX8

'**** Direct Sound ********
Private DS As DirectSound8
Private DS_Enum As DirectSoundEnum8

'**** MP3 ********
Dim IMC   As IMediaControl
Dim IBA   As IBasicAudio
Dim IME   As IMediaEvent
Dim IMPos As IMediaPosition

'**** Direct Music ********
Dim DMus_Preformance As DirectMusicPerformance8
Dim DMus_Loader As DirectMusicLoader8
Dim Music_Segment As DirectMusicSegment8

Public mMaxEntries As Long
Private mContador As Long
Private mCantidadBuffers As Long
Private mAmbienteActual As Integer
Private mLastAmbienteActual As Integer
Private mMidiActual As Integer
Private WavPath As String
Private MidiPath As String

Private Property Let SetWavPath(ByVal inString As String)
WavPath = inString
End Property

Private Property Let SetMidiPath(ByVal inString As String)
MidiPath = inString
End Property

Public Function GetBuffer(ByVal FileName As Integer) As DirectSoundSecondaryBuffer8

Dim retval As Integer

If FileName = 0 Then
    Set GetBuffer = Nothing
    Exit Function
End If

If mKeys(FileName) <> 0 Then
    mSNDDB(mKeys(FileName)).Accesos = mSNDDB(mKeys(FileName)).Accesos + 1
    Set GetBuffer = mBufferDB(mKeys(FileName))
Else
    retval = BuscaVacio()
    If retval = -1 Then
        retval = 1
        retval = BorraMenosUsado()
    End If
    
    Call CrearBuffer(FileName, retval)
    mSNDDB(retval).Accesos = mSNDDB(retval).Accesos + 1
    mKeys(FileName) = retval

    Set GetBuffer = mBufferDB(retval)
End If

End Function

Public Function Init(ByVal main_hwnd As String, ByVal Set_Wav_Path As String, ByVal Set_Midi_Path As String, ByVal BufferEntries As Integer, Optional ByVal Play_only_on_focus As Boolean = True, Optional ByVal Initialize_Sounds As Boolean = True, Optional ByVal Initialize_Music As Boolean = True) As Boolean
On Error GoTo ErrHandler

If BufferEntries < 1 Then
    Init = False
    Exit Function
End If

Set DX = New DirectX8

If Initialize_Sounds Then
    SetWavPath = Set_Wav_Path
    MaxEntries = BufferEntries
    ReDim mSNDDB(0 To mMaxEntries) As SND_DB_ENTRY
    ReDim mBufferDB(0 To mMaxEntries) As DirectSoundSecondaryBuffer8
    ReDim mKeys(0 To 300) As Integer
    Set DS_Enum = DX.GetDSEnum
    Set DS = DX.DirectSoundCreate(DS_Enum.GetGuid(1))

    If Play_only_on_focus Then
        DS.SetCooperativeLevel main_hwnd, DSSCL_NORMAL
    Else
        DS.SetCooperativeLevel main_hwnd, DSSCL_PRIORITY
    End If
End If

If Initialize_Music Then
    Dim DMus_Params As DMUS_AUDIOPARAMS
    Set DMus_Preformance = DX.DirectMusicPerformanceCreate
    SetMidiPath = Set_Midi_Path
    DMus_Preformance.InitAudio main_hwnd, DMUS_AUDIOF_ALL, DMus_Params, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128
    DMus_Preformance.SetMasterAutoDownload True
    
    DMus_Preformance.SetMasterTempo 1
    DMus_Preformance.SetMasterVolume 1

    VolumenActual = VolumenInicial
    SetMidiVolume (VolumenActual)

End If


mCantidadBuffers = 0
mContador = 0
Init = True

Exit Function

ErrHandler:
    Init = False
End Function

Public Function BorrarTodo() As Long

Dim i As Integer
For i = 1 To mMaxEntries
    If mSNDDB(i).FileName > 0 Then
        Call Borrar(mSNDDB(i).FileName)
    End If
Next

BorrarTodo = 1
mCantidadBuffers = 0

End Function

Public Function Borrar(ByVal FileName As Integer) As Integer

Borrar = -1
Call ZeroMemory(mSNDDB(mKeys(FileName)), Len(mSNDDB(mKeys(FileName))))
Set mBufferDB(mKeys(FileName)) = Nothing
mKeys(FileName) = 0
Borrar = 1
mCantidadBuffers = mCantidadBuffers - 1

End Function

Private Function CrearBuffer(ByVal Archivo As Integer, ByVal bIndex As Integer) As Integer

Dim dsbd As DSBUFFERDESC, retval As Long
dsbd.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME

With mSNDDB(bIndex)
    .FileName = Archivo
    .Accesos = 0
    .Cont = mContador
End With

mContador = mContador + 1

Set mBufferDB(bIndex) = DS.CreateSoundBufferFromFile(WavPath & Archivo & ".wav", dsbd)

CrearBuffer = 1
mCantidadBuffers = mCantidadBuffers + 1

End Function

Private Sub Class_Initialize()
mMaxEntries = 0
End Sub

Private Function BuscaVacio() As Integer

Dim i As Integer
For i = 1 To mMaxEntries
    If (mBufferDB(i) Is Nothing) Then
        BuscaVacio = i
        Exit Function
    End If
Next

BuscaVacio = -1

End Function

Private Function BorraMenosUsado() As Integer

BorraMenosUsado = 1

Dim lValor As Long
Dim lCont As Long

lValor = mSNDDB(1).Accesos
lCont = mSNDDB(1).Cont

Dim i As Integer

For i = 1 To mMaxEntries
    If mSNDDB(i).Accesos < lValor And mSNDDB(i).Cont < lCont Then
        lValor = mSNDDB(i).Accesos
        lCont = mSNDDB(i).Cont
        BorraMenosUsado = i
    End If
Next
mCantidadBuffers = mCantidadBuffers - 1

mKeys(mSNDDB(BorraMenosUsado).FileName) = 0
Set mBufferDB(BorraMenosUsado) = Nothing
Call ZeroMemory(mSNDDB(BorraMenosUsado), Len(mSNDDB(BorraMenosUsado)))

End Function

Private Sub Class_Terminate()

On Error Resume Next

Dim i As Integer

For i = 1 To mMaxEntries
    Set mBufferDB(i) = Nothing
Next

End Sub

Private Property Get MaxEntries() As Long
MaxEntries = mMaxEntries
End Property

Private Property Let MaxEntries(ByVal vNewValue As Long)
mMaxEntries = vNewValue
End Property

Public Property Get CantidadBuffers() As Variant
CantidadBuffers = mCantidadBuffers
End Property

Public Property Get AmbienteActual() As Long
AmbienteActual = mAmbienteActual
End Property

Public Property Let AmbienteActual(ByVal vNewValue As Long)
mAmbienteActual = vNewValue
End Property

Public Property Get LastAmbienteActual() As Long
AmbienteActual = mLastAmbienteActual
End Property

Public Property Let LastAmbienteActual(ByVal vNewValue As Long)
mLastAmbienteActual = vNewValue
End Property

Public Property Get MidiActual() As Long
MidiActual = mMidiActual
End Property

Public Property Let MidiActual(ByVal vNewValue As Long)
mMidiActual = vNewValue
End Property

Public Function PlaySound(ByVal Sound_File As Integer, Optional ByVal Do_Loop As Boolean = False) As Boolean

On Error GoTo Error_Handl
    
    If Sound_File <= 0 Then Exit Function
    
    If Do_Loop Then
        GetBuffer(Sound_File).Play DSBPLAY_LOOPING
    Else
        GetBuffer(Sound_File).Play DSBPLAY_DEFAULT
    End If
    PlaySound = True
    Exit Function
    
Error_Handl:
    PlaySound = False
End Function

Public Function StopSound(ByVal Sound_File As Integer) As Boolean

On Error GoTo Error_Handl
    If Sound_File <= 0 Then Exit Function
    GetBuffer(Sound_File).Stop
    GetBuffer(Sound_File).SetCurrentPosition 0
    
    StopSound = True
    Exit Function
Error_Handl:
    StopSound = False
End Function

Public Function PlayMidi(ByVal Music_Number As Integer, Optional ByVal Do_Loop As Boolean = True) As Boolean

    On Error GoTo Error_Handl

    'TODO: Buffers de Midi
    Set DMus_Loader = DX.DirectMusicLoaderCreate
    DMus_Loader.SetSearchDirectory MidiPath
    
    Set Music_Segment = DMus_Loader.LoadSegment(MidiPath & Music_Number & ".mid")
    Music_Segment.SetStandardMidiFile

    If Do_Loop Then
        Music_Segment.SetRepeats -1
    Else
        Music_Segment.SetRepeats 1
    End If

    DMus_Preformance.PlaySegmentEx Music_Segment, DMUS_SEGF_DEFAULT, 0
    mMidiActual = Music_Number
    PlayMidi = True
    Exit Function

Error_Handl:
    PlayMidi = False

End Function

Public Function SetMidiVolume(ByVal volume As Long) As Boolean

    On Error GoTo Error_Handl
    If volume > 200 Then volume = 200
    If volume < -4000 Then volume = -4000

    DMus_Preformance.SetMasterVolume volume

    SetMidiVolume = True
    Exit Function

Error_Handl:
End Function

Public Function EmptyMP3() As Boolean
    
    On Error GoTo Error_Handl
    
    If ObjPtr(IMC) > 0 Then
        IMC.Stop
    End If
    
    Set IBA = Nothing
    Set IME = Nothing
    Set IMPos = Nothing
    Set IMC = Nothing
    
    EmptyMP3 = True
    Exit Function

Error_Handl:
    EmptyMP3 = False
    End
End Function

Public Function LoadMP3(ByVal file_path As String, Optional ByVal volume As Long = 0, Optional ByVal balance As Long = 0) As Boolean '**** Loads a MP3 *****
    
    On Error GoTo Error_Handl
    
    If Not EmptyMP3() = True Then GoTo Error_Handl
    
    Set IMC = New FilgraphManager
    IMC.RenderFile file_path
    
    Set IBA = IMC
    If volume < -4000 Then volume = -4000
    IBA.volume = volume
    If balance > 5000 Then balance = 5000
    If balance < -5000 Then balance = -5000
    IBA.balance = balance
    
    Set IME = IMC
    
    Set IMPos = IMC
    If ObjPtr(IMPos) Then IMPos.Rate = 1#
    IMPos.CurrentPosition = 0
    
    LoadMP3 = True
    Exit Function

Error_Handl:
    LoadMP3 = False
End Function

Public Function PlayMP3() As Boolean
    
    On Error GoTo Error_Handl
    
    IMC.Run
    
    PlayMP3 = True
    Exit Function

Error_Handl:
    PlayMP3 = False
End Function

Public Function GetLoopMP3() As Boolean
    On Error GoTo Error_Handl
    If IMPos.StopTime = IMPos.CurrentPosition Then
        StopMP3
        PlayMP3
    End If
Error_Handl:
End Function

Public Function StopMP3() As Boolean

    On Error GoTo Error_Handl
    
    IMC.Stop
    IMPos.CurrentPosition = 0
    
    StopMP3 = True
    Exit Function

Error_Handl:
    StopMP3 = False
End Function

Public Function PauseMP3() As Boolean

    On Error GoTo Error_Handl
    
    IMC.Pause
    
    PauseMP3 = True
    Exit Function

Error_Handl:
    PauseMP3 = False
End Function

Public Function SetMP3Volume(ByVal Sound_Volume As Long) As Boolean

    On Error GoTo Error_Handl
    
    If Sound_Volume > 0 Then Sound_Volume = 0
    If Sound_Volume < -4000 Then Sound_Volume = -4000
    
    IBA.volume = Sound_Volume

    SetMP3Volume = True
    Exit Function

Error_Handl:
    SetMP3Volume = False
End Function
