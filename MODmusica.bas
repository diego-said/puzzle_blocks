Attribute VB_Name = "MusicaFundo"
Public DirectX As New DirectX7
Public DMPerformance As DirectMusicPerformance
Public DMSegment As DirectMusicSegment
Public DMSState As DirectMusicSegmentState
Public DMLoader As DirectMusicLoader
Public ini1 As Single
Public ini2 As Single
Public mus1 As Single
Public mus2 As Single
Public mus3 As Single
Public tempo As Long
Public inftempo As Long
Public infinito As Boolean
Public primeira As Boolean
Public para As Boolean
Public Sub IniciaVarMusica()
ini1 = 0: ini2 = 0: mus1 = 0: mus2 = 0: mus3 = 0: tempo = 0: inf1 = 0: inf2 = 0: inf3 = 0: infinito = False: primeira = True: para = True
End Sub
Public Sub SetaMusica()
Set DMLoader = DirectX.DirectMusicLoaderCreate()
Set DMPerformance = DirectX.DirectMusicPerformanceCreate()
Call DMPerformance.Init(Nothing, 0)
DMPerformance.SetPort -1, 16
Call DMPerformance.SetMasterAutoDownload(True)
End Sub
Public Sub ParaMusica()
If Not (DMSegment Is Nothing) And Not (DMSState Is Nothing) Then
    If DMPerformance.IsPlaying(DMSegment, DMSState) Then
        Call DMPerformance.Stop(DMSegment, DMSState, 0, 0)
    End If
End If
End Sub

Public Sub IniciaMusica()
DMLoader.SetSearchDirectory App.Path & "\musicas"
DMPerformance.SetMasterVolume (300)
tempo = tempo + 1
If tempo = 26 Then
    infinito = True
End If
inftempo = inftempo + 1
If ini1 = 0 And DMSegment Is Nothing And DMSState Is Nothing Then
    Set DMSegment = DMLoader.LoadSegment("zelda.mid")
    Set DMSState = DMPerformance.PlaySegment(DMSegment, 0, 0)
    ini1 = 1
ElseIf ini2 = 0 Then
    Set DMSegment = DMLoader.LoadSegment("inicio.mid")
    Set DMSState = DMPerformance.PlaySegment(DMSegment, 0, 0)
    ini2 = 1
    mus1 = 1
ElseIf infinito = True Then
    If primeira = True Then
        Set DMSegment = DMLoader.LoadSegment("musica1.mid")
        Set DMSState = DMPerformance.PlaySegment(DMSegment, 0, 0)
        mus1 = 0
        mus2 = 1
        primeira = False
        inftempo = 0
    Else
        If mus1 = 1 And inftempo = 165 Then
            Set DMSegment = DMLoader.LoadSegment("musica1.mid")
            Set DMSState = DMPerformance.PlaySegment(DMSegment, 0, 0)
            mus1 = 0
            mus2 = 1
            inftempo = 0
        ElseIf mus2 = 1 And inftempo = 163 Then
            Set DMSegment = DMLoader.LoadSegment("musica2.mid")
            Set DMSState = DMPerformance.PlaySegment(DMSegment, 0, 0)
            mus2 = 0
            mus3 = 1
            inftempo = 0
        ElseIf mus3 = 1 And inftempo = 290 Then
            Set DMSegment = DMLoader.LoadSegment("musica3.mid")
            Set DMSState = DMPerformance.PlaySegment(DMSegment, 0, 0)
            mus3 = 0
            mus1 = 1
            inftempo = 0
        End If
    End If
End If
End Sub
