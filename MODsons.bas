Attribute VB_Name = "Sons"
Public DirectSound As DirectSound
Public DSDBuffer As DSBUFFERDESC
Public DSWaveFormat As WAVEFORMATEX
Public DSLine As DirectSoundBuffer
Public DSMove As DirectSoundBuffer
Public DSLevel As DirectSoundBuffer
Public Sub SetaSons()
Set DirectSound = DirectX.DirectSoundCreate("")
DirectSound.SetCooperativeLevel FRMentrada.hWnd, DSSCL_NORMAL
DSDBuffer.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
With DSWaveFormat
    .nFormatTag = WAVE_FORMAT_PCM
    .nChannels = 2
    .lSamplesPerSec = 22050
    .nBitsPerSample = 16
    .nBlockAlign = .nBitsPerSample / 8 * .nChannels
    .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
End With
Set DSLine = DirectSound.CreateSoundBufferFromFile(App.Path & "\sons\completo.wav", DSDBuffer, DSWaveFormat)
Set DSMove = DirectSound.CreateSoundBufferFromFile(App.Path & "\sons\move.wav", DSDBuffer, DSWaveFormat)
Set DSLevel = DirectSound.CreateSoundBufferFromFile(App.Path & "\sons\level.wav", DSDBuffer, DSWaveFormat)
End Sub
Public Sub DescarregaSons()
Set DSSound = Nothing
Set DirectSound = Nothing
End Sub

