VERSION 5.00
Begin VB.Form FRMcontroles 
   Caption         =   "Tetris"
   ClientHeight    =   2940
   ClientLeft      =   3165
   ClientTop       =   345
   ClientWidth     =   3315
   Icon            =   "FRMcontroles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   3315
   Begin VB.Frame Frame1 
      Caption         =   "Como Jogar ?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.Image Image1 
         Height          =   405
         Index           =   3
         Left            =   120
         Picture         =   "FRMcontroles.frx":030A
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   405
         Index           =   2
         Left            =   120
         Picture         =   "FRMcontroles.frx":064E
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   405
         Index           =   1
         Left            =   120
         Picture         =   "FRMcontroles.frx":09A5
         Stretch         =   -1  'True
         Top             =   960
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   405
         Index           =   0
         Left            =   120
         Picture         =   "FRMcontroles.frx":0D2C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Modifica a forma da peça"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   2160
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desce a peça"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Move a peça para a direita"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Move a peça para a esquerda"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   2040
         Width           =   2550
      End
      Begin VB.Image Image2 
         Height          =   405
         Left            =   120
         Picture         =   "FRMcontroles.frx":1090
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pausa o jogo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   1
         Top             =   2520
         Width           =   1110
      End
   End
End
Attribute VB_Name = "FRMcontroles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
