Attribute VB_Name = "ControleJogo"
Public pont, lvla, lvl, lvl1, lvl2, lvl3, lvl4, lvl5, lvl6, lvl7, lvl8, lvl9 As Single
Public Sub Zera_Lvl()
    pont = 100
    lvla = 1
    lvl = 0
    lvl1 = 0
    lvl2 = 0
    lvl3 = 0
    lvl4 = 0
    lvl5 = 0
    lvl6 = 0
    lvl8 = 0
    lvl9 = 0
End Sub
Public Sub Controle_Lvl()
If (lvl >= 20 And lvl1 = 0) Then
    velocidade = 220
    lvl1 = 1
    pont = 100
    lvla = 1
    lvl = 0
    DSLevel.Play DSBPLAY_DEFAULT
    Unload FRMjogar
    FRMjogar.Show
ElseIf (lvl >= 20 And lvl2 = 0) Then
    velocidade = 200
    lvl2 = 1
    pont = 200
    lvla = 2
    lvl = 0
    DSLevel.Play DSBPLAY_DEFAULT
    Unload FRMjogar
    FRMjogar.Show
ElseIf (lvl >= 20 And lvl3 = 0) Then
    velocidade = 180
    lvl3 = 1
    pont = 300
    lvla = 3
    lvl = 0
    DSLevel.Play DSBPLAY_DEFAULT
    Unload FRMjogar
    FRMjogar.Show
ElseIf (lvl >= 20 And lvl4 = 0) Then
    velocidade = 160
    lvl4 = 1
    pont = 400
    lvla = 4
    lvl = 0
    DSLevel.Play DSBPLAY_DEFAULT
    Unload FRMjogar
    FRMjogar.Show
ElseIf (lvl >= 20 And lvl5 = 0) Then
    velocidade = 140
    lvl5 = 1
    pont = 500
    lvla = 5
    lvl = 0
    DSLevel.Play DSBPLAY_DEFAULT
    Unload FRMjogar
    FRMjogar.Show
ElseIf (lvl >= 20 And lvl6 = 0) Then
    velocidade = 120
    lvl6 = 1
    pont = 600
    lvla = 6
    lvl = 0
    DSLevel.Play DSBPLAY_DEFAULT
    Unload FRMjogar
    FRMjogar.Show
ElseIf (lvl >= 20 And lvl7 = 0) Then
    velocidade = 90
    lvl7 = 1
    pont = 700
    lvla = 7
    lvl = 0
    DSLevel.Play DSBPLAY_DEFAULT
    Unload FRMjogar
    FRMjogar.Show
ElseIf (lvl >= 20 And lvl8 = 0) Then
    velocidade = 70
    lvl8 = 1
    pont = 800
    lvla = 8
    lvl = 0
    DSLevel.Play DSBPLAY_DEFAULT
    Unload FRMjogar
    FRMjogar.Show
ElseIf (lvl >= 20 And lvl9 = 0) Then
    velocidade = 50
    lvl9 = 1
    pont = 900
    lvla = 9
    lvl = 0
    DSLevel.Play DSBPLAY_DEFAULT
    Unload FRMjogar
    FRMjogar.Show
ElseIf (lvl >= 20) Then
    Zera_Lvl
    DSLevel.Play DSBPLAY_DEFAULT
    Unload FRMjogar
    FRMjogar.Show
End If
End Sub

