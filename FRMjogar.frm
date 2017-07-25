VERSION 5.00
Begin VB.Form FRMjogar 
   BackColor       =   &H8000000A&
   Caption         =   "Tetris"
   ClientHeight    =   6270
   ClientLeft      =   2955
   ClientTop       =   1275
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMjogar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TMRlevel 
      Interval        =   1000
      Left            =   5400
      Top             =   120
   End
   Begin VB.Timer TempoMusica 
      Interval        =   1000
      Left            =   4920
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4440
      Top             =   120
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   3840
      TabIndex        =   3
      Top             =   3720
      Width           =   75
   End
   Begin VB.Image IMGpontos 
      Height          =   450
      Left            =   2880
      Picture         =   "FRMjogar.frx":030A
      Top             =   4200
      Width           =   1050
   End
   Begin VB.Image IMGlevel 
      Height          =   450
      Left            =   3000
      Picture         =   "FRMjogar.frx":0DFF
      Top             =   3600
      Width           =   675
   End
   Begin VB.Image IMGlinhas 
      Height          =   450
      Left            =   2880
      Picture         =   "FRMjogar.frx":162C
      Top             =   3000
      Width           =   900
   End
   Begin VB.Image IMGjogador 
      Height          =   450
      Left            =   2880
      Picture         =   "FRMjogar.frx":1FBF
      Top             =   2400
      Width           =   1050
   End
   Begin VB.Line Line9 
      X1              =   2880
      X2              =   5760
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Image IMGstatus 
      Height          =   450
      Left            =   3240
      Picture         =   "FRMjogar.frx":2B81
      Top             =   1560
      Width           =   2250
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   4080
      TabIndex        =   2
      Top             =   3120
      Width           =   75
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Top             =   4320
      Width           =   75
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   3000
      X2              =   3960
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   3000
      X2              =   3960
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   3000
      X2              =   3000
      Y1              =   360
      Y2              =   1080
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   3960
      X2              =   3960
      Y1              =   1080
      Y2              =   360
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   2560
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label LBLnome 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4200
      TabIndex        =   0
      Top             =   2520
      Width           =   75
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   0
      Left            =   3000
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   1
      Left            =   3240
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   2
      Left            =   3480
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   3
      Left            =   3720
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   4
      Left            =   3000
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   5
      Left            =   3240
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   6
      Left            =   3480
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   7
      Left            =   3720
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   8
      Left            =   3000
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   9
      Left            =   3240
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   10
      Left            =   3480
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Index           =   11
      Left            =   3720
      Top             =   840
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   6140
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   0
      Left            =   135
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   1
      Left            =   375
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   2
      Left            =   615
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   3
      Left            =   855
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   4
      Left            =   1095
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   5
      Left            =   1335
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   6
      Left            =   1575
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   7
      Left            =   1815
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   8
      Left            =   2055
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   9
      Left            =   2295
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   10
      Left            =   135
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   11
      Left            =   375
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   12
      Left            =   615
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   13
      Left            =   855
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   14
      Left            =   1095
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   15
      Left            =   1335
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   16
      Left            =   1575
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   17
      Left            =   1815
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   18
      Left            =   2055
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   19
      Left            =   2295
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   20
      Left            =   135
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   21
      Left            =   375
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   22
      Left            =   615
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   23
      Left            =   855
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   24
      Left            =   1095
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   25
      Left            =   1335
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   26
      Left            =   1575
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   27
      Left            =   1815
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   28
      Left            =   2055
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   29
      Left            =   2295
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   30
      Left            =   135
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   31
      Left            =   375
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   32
      Left            =   615
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   33
      Left            =   855
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   34
      Left            =   1095
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   35
      Left            =   1335
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   36
      Left            =   1575
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   37
      Left            =   1815
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   38
      Left            =   2055
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   39
      Left            =   2295
      Top             =   5160
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   40
      Left            =   135
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   41
      Left            =   375
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   42
      Left            =   615
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   43
      Left            =   855
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   44
      Left            =   1095
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   45
      Left            =   1335
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   46
      Left            =   1575
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   47
      Left            =   1815
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   48
      Left            =   2055
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   49
      Left            =   2295
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   50
      Left            =   135
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   51
      Left            =   375
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   52
      Left            =   615
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   53
      Left            =   855
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   54
      Left            =   1095
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   55
      Left            =   1335
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   56
      Left            =   1575
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   57
      Left            =   1815
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   58
      Left            =   2055
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   59
      Left            =   2295
      Top             =   4680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   60
      Left            =   135
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   61
      Left            =   375
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   62
      Left            =   615
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   63
      Left            =   855
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   64
      Left            =   1095
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   65
      Left            =   1335
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   66
      Left            =   1575
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   67
      Left            =   1815
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   68
      Left            =   2055
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   69
      Left            =   2295
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   70
      Left            =   135
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   71
      Left            =   375
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   72
      Left            =   615
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   73
      Left            =   855
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   74
      Left            =   1095
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   75
      Left            =   1335
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   76
      Left            =   1575
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   77
      Left            =   1815
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   78
      Left            =   2055
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   79
      Left            =   2295
      Top             =   4200
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   80
      Left            =   135
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   81
      Left            =   375
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   82
      Left            =   615
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   83
      Left            =   855
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   84
      Left            =   1095
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   85
      Left            =   1335
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   86
      Left            =   1575
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   87
      Left            =   1815
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   88
      Left            =   2055
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   89
      Left            =   2295
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   90
      Left            =   135
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   91
      Left            =   375
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   92
      Left            =   615
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   93
      Left            =   855
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   94
      Left            =   1095
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   95
      Left            =   1335
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   96
      Left            =   1575
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   97
      Left            =   1815
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   98
      Left            =   2055
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   99
      Left            =   2295
      Top             =   3720
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   100
      Left            =   135
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   101
      Left            =   375
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   102
      Left            =   615
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   103
      Left            =   855
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   104
      Left            =   1095
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   105
      Left            =   1335
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   106
      Left            =   1575
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   107
      Left            =   1815
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   108
      Left            =   2055
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   109
      Left            =   2295
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   110
      Left            =   135
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   111
      Left            =   375
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   112
      Left            =   615
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   113
      Left            =   855
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   114
      Left            =   1095
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   115
      Left            =   1335
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   116
      Left            =   1575
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   117
      Left            =   1815
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   118
      Left            =   2055
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   119
      Left            =   2295
      Top             =   3240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   120
      Left            =   135
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   121
      Left            =   375
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   122
      Left            =   615
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   123
      Left            =   855
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   124
      Left            =   1095
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   125
      Left            =   1335
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   126
      Left            =   1575
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   127
      Left            =   1815
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   128
      Left            =   2055
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   129
      Left            =   2295
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   130
      Left            =   135
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   131
      Left            =   375
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   132
      Left            =   615
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   133
      Left            =   855
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   134
      Left            =   1095
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   135
      Left            =   1335
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   136
      Left            =   1575
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   137
      Left            =   1815
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   138
      Left            =   2055
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   139
      Left            =   2295
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   140
      Left            =   135
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   141
      Left            =   375
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   142
      Left            =   615
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   143
      Left            =   855
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   144
      Left            =   1095
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   145
      Left            =   1335
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   146
      Left            =   1575
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   147
      Left            =   1815
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   148
      Left            =   2055
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   149
      Left            =   2295
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   150
      Left            =   135
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   151
      Left            =   375
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   152
      Left            =   615
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   153
      Left            =   855
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   154
      Left            =   1095
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   155
      Left            =   1335
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   156
      Left            =   1575
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   157
      Left            =   1815
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   158
      Left            =   2055
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   159
      Left            =   2295
      Top             =   2280
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   160
      Left            =   135
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   161
      Left            =   375
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   162
      Left            =   615
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   163
      Left            =   855
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   164
      Left            =   1095
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   165
      Left            =   1335
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   166
      Left            =   1575
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   167
      Left            =   1815
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   168
      Left            =   2055
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   169
      Left            =   2295
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   170
      Left            =   135
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   171
      Left            =   375
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   172
      Left            =   615
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   173
      Left            =   855
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   174
      Left            =   1095
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   175
      Left            =   1335
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   176
      Left            =   1575
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   177
      Left            =   1815
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   178
      Left            =   2055
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   179
      Left            =   2295
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   180
      Left            =   135
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   181
      Left            =   375
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   182
      Left            =   615
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   183
      Left            =   855
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   184
      Left            =   1095
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   185
      Left            =   1335
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   186
      Left            =   1575
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   187
      Left            =   1815
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   188
      Left            =   2055
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   189
      Left            =   2295
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   190
      Left            =   135
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   191
      Left            =   375
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   192
      Left            =   615
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   193
      Left            =   855
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   194
      Left            =   1095
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   195
      Left            =   1335
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   196
      Left            =   1575
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   197
      Left            =   1815
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   198
      Left            =   2055
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   199
      Left            =   2295
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   200
      Left            =   135
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   201
      Left            =   375
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   202
      Left            =   615
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   203
      Left            =   855
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   204
      Left            =   1095
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   205
      Left            =   1335
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   206
      Left            =   1575
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   207
      Left            =   1815
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   208
      Left            =   2055
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   209
      Left            =   2295
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   210
      Left            =   135
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   211
      Left            =   375
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   212
      Left            =   615
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   213
      Left            =   855
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   214
      Left            =   1095
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   215
      Left            =   1335
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   216
      Left            =   1575
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   217
      Left            =   1815
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   218
      Left            =   2055
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   219
      Left            =   2295
      Top             =   840
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   220
      Left            =   135
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   221
      Left            =   375
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   222
      Left            =   615
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   223
      Left            =   855
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   224
      Left            =   1095
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   225
      Left            =   1335
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   226
      Left            =   1575
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   227
      Left            =   1815
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   228
      Left            =   2055
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   229
      Left            =   2295
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   230
      Left            =   135
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   231
      Left            =   375
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   232
      Left            =   615
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   233
      Left            =   855
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   234
      Left            =   1095
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   235
      Left            =   1335
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   236
      Left            =   1575
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   237
      Left            =   1815
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   238
      Left            =   2055
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   239
      Left            =   2295
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   240
      Left            =   135
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   241
      Left            =   375
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   242
      Left            =   615
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   243
      Left            =   855
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   244
      Left            =   1095
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   245
      Left            =   1335
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   246
      Left            =   1575
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   247
      Left            =   1815
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   248
      Left            =   2055
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   249
      Left            =   2295
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   2560
      X2              =   2560
      Y1              =   120
      Y2              =   6140
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   120
      X2              =   2560
      Y1              =   6150
      Y2              =   6150
   End
End
Attribute VB_Name = "FRMjogar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Single
Dim j As Single
DSMove.Play DSBPLAY_DEFAULT
If KeyCode = 80 Then
    If parajogo = False Then
        parajogo = True
    Else
        parajogo = False
    End If
    Exit Sub
End If
If parajogo = False Then
If KeyCode = 38 And Fig01 > 1 Then
    girar
ElseIf KeyCode = 37 Then
ok = 1
For i = 1 To 4
    For j = 1 To 4
        If s(i, j) = True And X + i - 1 - 1 > 0 Then
            If n(X + i - 1 - 1, Y - j + 1) = True Then
                ok = 0
            End If
        End If
        If s(i, j) = True And X + i - 1 - 1 < 1 Then
            ok = 0
        End If
    Next j
Next i
If ok = 1 Then
    X = X - 1
    For i = 1 To 4
        For j = 1 To 4
            If s(i, j) = True Then
                If s(i + 1, j) = False Then
                    Shape1(formula(X + i, Y - j + 1)).FillColor = RGB(255, 255, 255)
                End If
            End If
            If s(i, j) = True Then
                Shape1(formula(X + i - 1, Y - j + 1)).FillColor = Fig03
            End If
        Next j
    Next i
End If
ElseIf KeyCode = 39 Then
ok = 1
For i = 1 To 4
    For j = 1 To 4
        If s(i, j) = True And X + i - 1 + 1 < 11 Then
            If n(X + i - 1 + 1, Y - j + 1) = True Then
                ok = 0
            End If
        End If
        If s(i, j) = True And X + i - 1 + 1 > 10 Then
            ok = 0
        End If
    Next j
Next i
If ok = 1 Then
    X = X + 1
    For i = 1 To 4
        For j = 1 To 4
            If s(i, j) = True Then
                If s(i - 1, j) = False Then
                    Shape1(formula(X + i - 2, Y - j + 1)).FillColor = RGB(255, 255, 255)
                End If
            End If
            If s(i, j) = True Then
                Shape1(formula(X + i - 1, Y - j + 1)).FillColor = Fig03
            End If
        Next j
    Next i
End If
ElseIf KeyCode = 40 And Fig01 > 0 Then
For i = 1 To 4
    For j = 1 To 4
        If s(i, j) = True Then
            Shape1(formula(X + i - 1, Y - j + 1)).FillColor = RGB(255, 255, 255)
        End If
    Next j
Next i
Do
    Y = Y - 1
    For i = 1 To 4
        For j = 1 To 4
            If s(i, j) = True And Y - j + 1 = 1 Then
                Fig01 = 0
            End If
            If Y - j > 0 And X + i - 1 > 0 And X + i - 1 < 11 Then
                If s(i, j) = True And n(X + i - 1, Y - j) = True Then
                    Fig01 = 0
                End If
            End If
        Next j
    Next i
Loop Until Fig01 = 0
For i = 1 To 4
    For j = 1 To 4
        If s(i, j) = True Then
            Shape1(formula(X + i - 1, Y - j + 1)).FillColor = Fig03
        End If
        If s(i, j) = True Then
            n(X + i - 1, Y - j + 1) = True
        End If
    Next j
Next i
End If
End If
End Sub
Private Sub Form_Load()
Dim i As Single
Dim j As Single
parajogo = False

For i = 1 To 4
    For j = 1 To 4
        s2(i, j) = False
    Next j
Next i

For i = 1 To 10
    For j = 1 To 25
        n(i, j) = False
    Next j
Next i


For i = 0 To 11
    Shape2(i).BorderColor = RGB(255, 255, 255)
    Shape2(i).FillStyle = 0
    Shape2(i).FillColor = RGB(255, 255, 255)
Next i

For i = 0 To 249
    Shape1(i).BorderColor = RGB(255, 255, 255)
    Shape1(i).FillStyle = 0
    Shape1(i).FillColor = RGB(255, 255, 255)
Next i

Timer1.Interval = velocidade
LBLnome.Caption = nome
Fig02 = 0
inicio
End Sub
Function formula(xx, yy)
formula = (yy - 1) * 10 + xx - 1
End Function

Private Sub TempoMusica_Timer()
If para = True Then
    ParaMusica
    para = False
End If
IniciaMusica
End Sub

Private Sub Timer1_Timer()
Dim i As Single
Dim j As Single
If parajogo = False Then
For i = 1 To 4
    For j = 1 To 4
        If s(i, j) = True And Y - j + 1 = 1 Then
            Fig01 = 0
        Else
            If Y - j > 0 And X + i - 1 > 0 And X + i - 1 < 11 Then
                If s(i, j) = True And n(X + i - 1, Y - j) = True Then
                    Fig01 = 0
                End If
            End If
        End If
    Next j
Next i

If Fig01 = 0 Then
    For i = 1 To 4
        For j = 1 To 4
            If s(i, j) = True Then
                n(X + i - 1, Y - j + 1) = True
            End If
        Next j
    Next i
    cont_pontos
    inicio
Else
    For i = 1 To 4
        For j = 1 To 4
            If s(i, j) = True And s(i, j - 1) = False Then
                Shape1(formula(X + i - 1, Y - j + 1)).FillColor = RGB(255, 255, 255)
            End If
        Next j
    Next i
End If

Y = Y - 1

For i = 1 To 4
    For j = 1 To 4
        If s(i, j) = True Then
            If Shape1(formula(X + i - 1, Y - j + 1)).FillColor <> RGB(255, 255, 255) And Shape1(formula(X + i - 1, Y - j + 1)).FillColor <> Fig03 Then
                fim
                Exit Sub
            End If
        End If
        If s(i, j) = True Then
            Shape1(formula(X + i - 1, Y - j + 1)).FillColor = Fig03
        End If
    Next j
Next i
End If
End Sub
Public Sub inicio()
Dim i As Single
Dim j As Single
Timer1.Interval = velocidade
Fig01 = Fig02
Fig03 = Fig04
Fig02 = Int(Rnd * 7) + 1
X = 4
Y = 26
Rot = 1

For i = 1 To 4
    For j = 1 To 4
        s(i, j) = s2(i, j)
        s2(i, j) = 0
    Next j
Next i

Select Case Fig02
    Case 1
        s2(2, 2) = True
        s2(3, 2) = True
        s2(2, 3) = True
        s2(3, 3) = True
        Fig04 = RGB(0, 0, 0)
    Case 2
        s2(1, 2) = True
        s2(2, 2) = True
        s2(3, 2) = True
        s2(4, 2) = True
        Fig04 = RGB(0, 200, 0)
    Case 3
        s2(2, 1) = True
        s2(3, 1) = True
        s2(3, 2) = True
        s2(3, 3) = True
        Fig04 = RGB(255, 0, 0)
    Case 4
        s2(3, 1) = True
        s2(2, 1) = True
        s2(2, 2) = True
        s2(2, 3) = True
        Fig04 = RGB(0, 0, 255)
    Case 5
        s2(3, 1) = True
        s2(3, 2) = True
        s2(3, 3) = True
        s2(2, 2) = True
        Fig04 = RGB(150, 150, 0)
    Case 6
        s2(2, 1) = True
        s2(2, 2) = True
        s2(3, 2) = True
        s2(3, 3) = True
        Fig04 = RGB(0, 150, 150)
    Case 7
        s2(3, 1) = True
        s2(3, 2) = True
        s2(2, 2) = True
        s2(2, 3) = True
        Fig04 = RGB(150, 0, 150)
End Select

For i = 1 To 4
    For j = 1 To 3
        Shape2((j - 1) * 4 + i - 1).FillColor = &HE0E0E0
        If s2(i, j) = True Then
            Shape2((j - 1) * 4 + i - 1).FillColor = Fig04
        End If
    Next j
Next i

End Sub

Public Sub girar()
Rot2 = Rot + 1
If Rot2 = 5 Then Rot2 = 1
If (Fig01 = 2 Or Fig01 > 5) And Rot2 = 3 Then Rot2 = 1
For i = 1 To 4: For i2 = 1 To 4: s3(i, i2) = 0: Next i2: Next i
Select Case Fig01
Case 2
Select Case Rot2
Case 1
s3(1, 2) = True
s3(2, 2) = True
s3(3, 2) = True
s3(4, 2) = True
Case 2
s3(2, 1) = True
s3(2, 2) = True
s3(2, 3) = True
s3(2, 4) = True
End Select
Case 3
Select Case Rot2
Case 1
s3(2, 1) = True
s3(3, 1) = True
s3(3, 2) = True
s3(3, 3) = True
Case 2
s3(4, 1) = True
s3(4, 2) = True
s3(3, 2) = True
s3(2, 2) = True
Case 3
s3(3, 3) = True
s3(2, 3) = True
s3(2, 2) = True
s3(2, 1) = True
Case 4
s3(2, 2) = True
s3(2, 1) = True
s3(3, 1) = True
s3(4, 1) = True
End Select
Case 4
Select Case Rot2
Case 1
s3(3, 1) = True
s3(2, 1) = True
s3(2, 2) = True
s3(2, 3) = True
Case 2
s3(2, 1) = True
s3(3, 1) = True
s3(4, 1) = True
s3(4, 2) = True
Case 3
s3(3, 1) = True
s3(3, 2) = True
s3(3, 3) = True
s3(2, 3) = True
Case 4
s3(2, 1) = True
s3(2, 2) = True
s3(3, 2) = True
s3(4, 2) = True
End Select
Case 5
Select Case Rot2
Case 1
s3(3, 1) = True
s3(3, 2) = True
s3(3, 3) = True
s3(2, 2) = True
Case 2
s3(3, 1) = True
s3(2, 2) = True
s3(3, 2) = True
s3(4, 2) = True
Case 3
s3(2, 1) = True
s3(2, 2) = True
s3(2, 3) = True
s3(3, 2) = True
Case 4
s3(2, 1) = True
s3(3, 1) = True
s3(4, 1) = True
s3(3, 2) = True
End Select
Case 6
Select Case Rot2
Case 1
s3(2, 1) = True
s3(2, 2) = True
s3(3, 2) = True
s3(3, 3) = True
Case 2
s3(2, 2) = True
s3(3, 2) = True
s3(3, 1) = True
s3(4, 1) = True
End Select
Case 7
Select Case Rot2
Case 1
s3(3, 1) = True
s3(3, 2) = True
s3(2, 2) = True
s3(2, 3) = True
Case 2
s3(2, 1) = True
s3(3, 1) = True
s3(3, 2) = True
s3(4, 2) = True
End Select
End Select
ok = 1
For i = 1 To 4: For i2 = 1 To 4
If s3(i, i2) = True Then
If X + i - 1 < 1 Or X + i - 1 > 10 Or Y - i2 + 1 < 1 Then ok = 0
If ok = 1 Then
If n(X + i - 1, Y - i2 + 1) = True Then ok = 0
End If
End If
Next i2: Next i
If ok = 0 Then Exit Sub
Rot = Rot2
For i = 1 To 4: For i2 = 1 To 4
If s3(i, i2) = True And s(i, i2) = False Then Shape1(formula(X + i - 1, Y - i2 + 1)).FillColor = Fig03
If s3(i, i2) = False And s(i, i2) = True Then Shape1(formula(X + i - 1, Y - i2 + 1)).FillColor = RGB(255, 255, 255)
s(i, i2) = s3(i, i2)
Next i2: Next i
End Sub
Public Sub cont_pontos()
Dim i As Single
Dim j As Single
Dim k As Single
For i2 = 25 To 1 Step -1
ok = 1
For i = 1 To 10
If n(i, i2) = False Then ok = 0
Next i
If ok = 1 Then
FRMjogar.SetFocus
Pontos = Pontos + pont
DSLine.Play DSBPLAY_DEFAULT
linhas = linhas + 1
lvl = lvl + 1
For i = 1 To 10
For i3 = i2 To 24
n(i, i3) = n(i, i3 + 1)
Shape1(formula(i, i3)).FillColor = Shape1(formula(i, i3 + 1)).FillColor
Next i3
Next i
End If
Next i2
End Sub

Public Sub fim()
MsgBox "Parabéns" & " " & nome & " " & "sua pontuação foi de" & " " & Pontos
FRMrecordes.LSTrecordes.AddItem nome & ": " & Pontos
Salva_Recorde
Unload Me
FRMentrada.Show
End Sub

Private Sub TMRlevel_Timer()
Label6(0).Caption = lvla
Label6(1).Caption = Pontos
Label6(2).Caption = linhas
Controle_Lvl
End Sub
