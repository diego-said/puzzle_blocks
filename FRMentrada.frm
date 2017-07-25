VERSION 5.00
Begin VB.Form FRMentrada 
   Caption         =   "Tetris"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   2985
   Icon            =   "FRMentrada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   2985
   Begin VB.CommandButton CMDjogar 
      Caption         =   "JOGAR"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox TXTnome 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Jogador:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   0
      Picture         =   "FRMentrada.frx":030A
      Top             =   0
      Width           =   3000
   End
   Begin VB.Menu Jogo 
      Caption         =   "&Jogo"
      Begin VB.Menu Sair 
         Caption         =   "&Sair"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Ajuda 
      Caption         =   "&Opções"
      Begin VB.Menu Como_jogar 
         Caption         =   "&Como jogar"
      End
      Begin VB.Menu Recordes 
         Caption         =   "&Recordes"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu Sobre 
      Caption         =   "&Sobre"
      Begin VB.Menu Tetris 
         Caption         =   "&Tetris"
      End
      Begin VB.Menu Autor 
         Caption         =   "&Autor"
      End
   End
End
Attribute VB_Name = "FRMentrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Autor_Click()
MsgBox "Este jogo é de livre uso, podendo ser auterado por quem desejar. Ele foi desenvolvido por Diego Alves Said Simão (zephon@operamail.com) com base em outros projetos encontrados na internet.", vbOKOnly
End Sub

Private Sub CMDjogar_Click()
nome = TXTnome.Text
If Trim(nome) = "" Then
MsgBox ("Digite seu nome ou apelido!")
TXTnome.SetFocus
Else
velocidade = 220
linhas = 0
lvla = 1
lvl1 = 1
pont = 100
Unload Me
FRMjogar.Show
FRMjogar.Caption = "Jogador:" & " " & nome
End If
End Sub

Private Sub Como_jogar_Click()
FRMcontroles.Show
End Sub

Private Sub Form_Load()
Randomize Timer
IniciaVarMusica
SetaMusica
SetaSons
IniciaMusica
End Sub

Private Sub Recordes_Click()
FRMrecordes.Show
End Sub

Private Sub Sair_Click()
End
End Sub

Private Sub Tetris_Click()
MsgBox "Tetris - Versão desenvolvida em 2004", vbOKOnly
End Sub
