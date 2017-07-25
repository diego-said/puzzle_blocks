VERSION 5.00
Begin VB.Form FRMrecordes 
   Caption         =   "Tetris"
   ClientHeight    =   3255
   ClientLeft      =   3165
   ClientTop       =   345
   ClientWidth     =   4710
   Icon            =   "FRMrecordes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4710
   Begin VB.Frame Frame1 
      Caption         =   "Recordes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.ListBox LSTrecordes 
         Height          =   2790
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4455
      End
   End
End
Attribute VB_Name = "FRMrecordes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Carrega_Recorde
End Sub
