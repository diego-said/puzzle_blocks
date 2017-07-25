Attribute VB_Name = "Recorde"
Public Sub Salva_Recorde()
Dim ind As Single
Dim aux As String
Dim cont As Single
File = FreeFile
ind = FRMrecordes.LSTrecordes.ListCount
cont = 0
Open App.Path & "\recordes.txt" For Output As #File
    Print #File, nome & ": " & Pontos
    Do While (ind - 1) > 0
        aux = FRMrecordes.LSTrecordes.List(cont)
        Print #File, aux
        cont = cont + 1
        ind = ind - 1
    Loop
Close #File
End Sub
Public Sub Carrega_Recorde()
Dim StrTemp
File = FreeFile
Open App.Path & "\recordes.txt" For Input As #File
    Do While Not EOF(File)
        Input #File, StrTemp
        FRMrecordes.LSTrecordes.AddItem StrTemp
    Loop
Close #File
End Sub
