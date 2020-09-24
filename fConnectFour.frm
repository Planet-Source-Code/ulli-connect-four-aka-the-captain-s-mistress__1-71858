VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fConnectFour 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'Kein
   Caption         =   "Connect Four"
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   ForeColor       =   &H00000000&
   Icon            =   "fConnectFour.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   7845
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btReset 
      BackColor       =   &H0080FF80&
      Caption         =   "&New Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Grafisch
      TabIndex        =   22
      ToolTipText     =   "Play again?"
      Top             =   7155
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   195
      Top             =   7665
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton btExit 
      BackColor       =   &H008080FF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6615
      Style           =   1  'Grafisch
      TabIndex        =   13
      ToolTipText     =   "Play again?"
      Top             =   7155
      Width           =   900
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   180
      Picture         =   "fConnectFour.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbHint 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Show  Hint "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   285
      Left            =   6600
      TabIndex        =   27
      ToolTipText     =   "Hints might not be perfect"
      Top             =   1500
      Width           =   1005
   End
   Begin VB.Label lbTakeBack 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Take Back"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   285
      Left            =   6600
      TabIndex        =   26
      ToolTipText     =   "Takes back last computer move|and your last move"
      Top             =   1905
      Width           =   1005
   End
   Begin VB.Label lbReplay 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Replay"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   6600
      TabIndex        =   25
      ToolTipText     =   "Replays this game"
      Top             =   2310
      Width           =   1005
   End
   Begin VB.Label lbLoad 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Load Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6600
      TabIndex        =   24
      ToolTipText     =   "Loads a game from file"
      Top             =   2715
      Width           =   1005
   End
   Begin VB.Label lbSave 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Save Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6600
      TabIndex        =   23
      ToolTipText     =   "Saves the current game status"
      Top             =   3120
      Width           =   1005
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Index           =   6
      Left            =   5835
      TabIndex        =   21
      Top             =   1545
      Width           =   105
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Index           =   5
      Left            =   5040
      TabIndex        =   20
      Top             =   1530
      Width           =   105
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Index           =   4
      Left            =   4245
      TabIndex        =   19
      Top             =   1530
      Width           =   105
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Index           =   3
      Left            =   3450
      TabIndex        =   18
      Top             =   1530
      Width           =   105
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Index           =   2
      Left            =   2655
      TabIndex        =   17
      Top             =   1530
      Width           =   105
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Index           =   1
      Left            =   1860
      TabIndex        =   16
      Top             =   1530
      Width           =   105
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Index           =   0
      Left            =   1065
      TabIndex        =   15
      Top             =   1530
      Width           =   105
   End
   Begin VB.Label lbTimeCheck 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "tc"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   165
      Left            =   7050
      TabIndex        =   14
      Top             =   5790
      Width           =   105
   End
   Begin VB.Image imgUMGEDV 
      Height          =   630
      Left            =   6765
      Picture         =   "fConnectFour.frx":1194
      ToolTipText     =   "Author's e-mail address:|                                  |umgedv@yahoo.com|                               "
      Top             =   6225
      Width           =   675
   End
   Begin VB.Line lnQuad 
      BorderColor     =   &H002418ED&
      BorderWidth     =   11
      DrawMode        =   15  'Stift und inverse Anzeige mischen
      Visible         =   0   'False
      X1              =   2070
      X2              =   2400
      Y1              =   7785
      Y2              =   7785
   End
   Begin VB.Image imgSmile 
      Height          =   240
      Left            =   270
      Picture         =   "fConnectFour.frx":2826
      Top             =   7290
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape shp 
      BorderColor     =   &H00C08000&
      BorderWidth     =   2
      Height          =   360
      Index           =   2
      Left            =   7305
      Shape           =   3  'Kreis
      Top             =   180
      Width           =   360
   End
   Begin VB.Shape shp 
      BorderColor     =   &H00C08000&
      BorderWidth     =   2
      Height          =   360
      Index           =   1
      Left            =   6930
      Shape           =   3  'Kreis
      Top             =   180
      Width           =   360
   End
   Begin VB.Label lbMini 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "â"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6975
      TabIndex        =   12
      ToolTipText     =   "Minimize"
      Top             =   255
      Width           =   255
   End
   Begin VB.Image imgThink 
      Height          =   270
      Left            =   240
      Picture         =   "fConnectFour.frx":2B68
      Top             =   7290
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lbExit 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "û"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   7365
      TabIndex        =   11
      ToolTipText     =   "Exit"
      Top             =   150
      Width           =   240
   End
   Begin VB.Shape shp 
      BorderColor     =   &H00C08000&
      BorderWidth     =   2
      Height          =   360
      Index           =   0
      Left            =   6555
      Shape           =   3  'Kreis
      Top             =   180
      Width           =   360
   End
   Begin VB.Label lbVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C08000&
      Height          =   165
      Left            =   270
      TabIndex        =   10
      Top             =   6900
      Width           =   90
   End
   Begin VB.Label lbHelp 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   270
      Left            =   6585
      TabIndex        =   9
      ToolTipText     =   $"fConnectFour.frx":2F9A
      Top             =   195
      Width           =   300
   End
   Begin VB.Line ln 
      BorderColor     =   &H00C08000&
      BorderWidth     =   3
      X1              =   270
      X2              =   7500
      Y1              =   6855
      Y2              =   6855
   End
   Begin VB.Label lbOpt 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Champion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   3
      Left            =   6600
      TabIndex        =   8
      ToolTipText     =   "Playing Level"
      Top             =   5385
      Width           =   1005
   End
   Begin VB.Label lbOpt 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Expert"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   2
      Left            =   6600
      TabIndex        =   7
      ToolTipText     =   "Playing Level"
      Top             =   4980
      Width           =   1005
   End
   Begin VB.Label lbOpt 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   1
      Left            =   6600
      TabIndex        =   6
      ToolTipText     =   "Playing Level"
      Top             =   4575
      Width           =   1005
   End
   Begin VB.Label lbOpt 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Beginner"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Index           =   0
      Left            =   6600
      TabIndex        =   5
      ToolTipText     =   "Playing Level"
      Top             =   4170
      Width           =   1005
   End
   Begin VB.Label lbAuthor 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   3480
      TabIndex        =   4
      Top             =   6375
      Width           =   60
   End
   Begin VB.Label lbWinIn 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   3465
      TabIndex        =   3
      Top             =   6630
      Width           =   75
   End
   Begin VB.Image imgFoot 
      Height          =   345
      Index           =   1
      Left            =   4920
      Picture         =   "fConnectFour.frx":317B
      Top             =   6495
      Width           =   675
   End
   Begin VB.Image imgFoot 
      Height          =   345
      Index           =   0
      Left            =   1410
      Picture         =   "fConnectFour.frx":3DF5
      Top             =   6495
      Width           =   675
   End
   Begin VB.Label lbWins 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   585
      TabIndex        =   2
      Top             =   7290
      Width           =   60
   End
   Begin VB.Label lbTitle 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C08000&
      Height          =   405
      Left            =   3390
      TabIndex        =   1
      Top             =   180
      Width           =   255
   End
   Begin VB.Image imgDropB 
      Height          =   615
      Index           =   0
      Left            =   1305
      Picture         =   "fConnectFour.frx":4A6F
      Top             =   7710
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgDropW 
      Height          =   615
      Index           =   0
      Left            =   675
      Picture         =   "fConnectFour.frx":5E8D
      Top             =   7710
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgWhite 
      Height          =   615
      Index           =   6
      Left            =   5580
      MouseIcon       =   "fConnectFour.frx":72AB
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":7B75
      Top             =   690
      Width           =   615
   End
   Begin VB.Image imgWhite 
      Height          =   615
      Index           =   5
      Left            =   4785
      MouseIcon       =   "fConnectFour.frx":8F93
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":985D
      Top             =   690
      Width           =   615
   End
   Begin VB.Image imgWhite 
      Height          =   615
      Index           =   4
      Left            =   3990
      MouseIcon       =   "fConnectFour.frx":AC7B
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":B545
      Top             =   690
      Width           =   615
   End
   Begin VB.Image imgWhite 
      Height          =   615
      Index           =   3
      Left            =   3195
      MouseIcon       =   "fConnectFour.frx":C963
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":D22D
      Top             =   690
      Width           =   615
   End
   Begin VB.Image imgWhite 
      Height          =   615
      Index           =   2
      Left            =   2400
      MouseIcon       =   "fConnectFour.frx":E64B
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":EF15
      Top             =   690
      Width           =   615
   End
   Begin VB.Image imgWhite 
      Height          =   615
      Index           =   1
      Left            =   1605
      MouseIcon       =   "fConnectFour.frx":10333
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":10BFD
      Top             =   690
      Width           =   615
   End
   Begin VB.Image imgWhite 
      Height          =   615
      Index           =   0
      Left            =   810
      MouseIcon       =   "fConnectFour.frx":1201B
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":128E5
      Top             =   690
      Width           =   615
   End
   Begin VB.Image imgBlack 
      Height          =   615
      Index           =   6
      Left            =   5580
      MouseIcon       =   "fConnectFour.frx":13D03
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":145CD
      Top             =   690
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBlack 
      Height          =   615
      Index           =   5
      Left            =   4785
      MouseIcon       =   "fConnectFour.frx":159EB
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":162B5
      Top             =   690
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBlack 
      Height          =   615
      Index           =   4
      Left            =   3990
      MouseIcon       =   "fConnectFour.frx":176D3
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":17F9D
      Top             =   690
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBlack 
      Height          =   615
      Index           =   3
      Left            =   3195
      MouseIcon       =   "fConnectFour.frx":193BB
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":19C85
      Top             =   690
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBlack 
      Height          =   615
      Index           =   2
      Left            =   2400
      MouseIcon       =   "fConnectFour.frx":1B0A3
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":1B96D
      Top             =   690
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBlack 
      Height          =   615
      Index           =   1
      Left            =   1605
      MouseIcon       =   "fConnectFour.frx":1CD8B
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":1D655
      Top             =   690
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBlack 
      Height          =   615
      Index           =   0
      Left            =   810
      MouseIcon       =   "fConnectFour.frx":1EA73
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fConnectFour.frx":1F33D
      Top             =   690
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgBoard 
      Height          =   5190
      Left            =   510
      Picture         =   "fConnectFour.frx":2075B
      Top             =   1440
      Width           =   5955
   End
End
Attribute VB_Name = "fConnectFour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************
'   Connect Four aka The Captain's Mistress
'********************************************

'During his long sea voyages, Captain Cook was apparently often absent in the evenings and eventually
'the crew began to joke that he must have a mistress in his cabin. When they discovered that the Captain
'had simply been playing this game with the ship's scientists, the game was christened "The Captain's Mistress"

'Uses bitmaps for the board, iterative search deepening, alpha-beta pruning, and principal variation search.
'No positional evaluation is made at the leaves, it simply relies on search depth to find winning combinations.
'Search depth is about 10 to 12 with a 0.1 seconds time check (beginners level).

'The board looks like this using two orthogonal and two diagonal bitmaps for each color, and a bit is set for each
'piece at that location. This arrangement limits the number of (non-loop!)-tests for a connected quad to 13 or less.

'                                 1   2   4   8  16  32  64

'                               +---+---+---+---+---+---+---+
'                               |   |   |   | o |   |   |   |  32
'                               +---+---+---+---+---+---+---+
'                               |   |   |   | o |   |   |   |  16
'                               +---+---+---+---+---+---+---+
'                               |   |   |   | o |   |   |   |   8
'                 west-east --> +---+---+---+---+---+---+---+
'                               |   |   |   | o |   |   |   |   4
'                               +---+---+---+---+---+---+---+
'                               |   |   |   |   |   |   |   |   2
'                               +---+---+---+---+---+---+---+
'                               |   |   |   |   |   |   |   |   1
'                               +---+---+---+---+---+---+---+
'                             /    ^   ^   ^   ^   ^   ^   ^  \
'                            /           south-north           \
'                  southwest-northeast                    southeast-northwest

'for example:
'The white quad shown above will show up in byte SNW(3) as b00111100 (d60) and can be tested with an appropriate AND-mask.
'Only those bitmaps are tested which were affected by the last move; this makes the search pretty fast.

'The Hint option shows the move which the computer has worked out for you during HER search, no search is made for YOU.

'+++++++++++++++++++++
' Development History
'+++++++++++++++++++++

'14Mar2009 - some speed improvements by eliminating duplicate calculations and using API functions for moving data
'            and by arranging that moved data start at dword boundaries (7% improvement)
'            the search has also been improved by altering the move ordering (16% improvement)
'            eliminated a few dead variables
'            taken care of different CPU clock frequencies and of running in IDE

'26Mar2009 - using GetTickCount instead of Timer
'            some pimping up of the UI
'            added balloon tooltips
'            remove chip on click so you can't click twice
'            changed timing control
'            fixed bug with prediction and Draw detection
'            changed chip dropping routine
'            added flashing
'            added quad marker
'            added keyboard operation
'            added TakeBack
'            added Replay
'            eliminated deadwood

Private Declare Function Beeper Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Sub CopyBoard Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, Optional ByVal Length As Long = 84)
Private Declare Sub CopyMoves Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, Optional ByVal Length As Long = 28)
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As tPoint) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function InflateRect Lib "user32" (lpRect As tRECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32" () 'a manifest file, if present, may influence the popup balloons
Private Declare Function MessageBeep Lib "user32" (ByVal wType As VbMsgBoxStyle) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As tRECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type tPoint
    X                       As Long
    Y                       As Long
End Type
Private CP                  As tPoint   'cursor position

Private Type tRECT
    LT                      As tPoint   'top left
    RB                      As tPoint   'bot rite
End Type
Private WindowRect          As tRECT    'for drawing the frame

'+++++++++++++++++
'The Board
'+++++++++++++++++
Private Type tBoard                     'is moved rather frequently - so it is kept as short as possible (84 bytes)
    Filled(0 To 6)          As Byte     'column fill
    SNW(0 To 6)             As Byte     'south -> north white
    SNB(0 To 6)             As Byte     'south -> north black
    WEW(0 To 5)             As Byte     'east -> west white
    WEB(0 To 5)             As Byte     'east -> west black
    SwNeW(-3 To 8)          As Byte     'south-west -> north-east white
    SwNeB(-3 To 8)          As Byte     'south-west -> north-east black
    SeNwW(-3 To 8)          As Byte     'south-east -> north-west white
    SeNwB(-3 To 8)          As Byte     'south-east -> north-west black
    SideToMove              As Byte
    Extra1                  As Byte     'bring it up to 84 bytes so it will hopefully use dword moves only
    Extra2                  As Byte
End Type

Private Board               As tBoard   'the board bitmaps
Private OrdMovs(0 To 6)     As Long

Private Enum eConsts
    Infinity = 99999                    'value outside normal range
    DepthLimit = 30                     'max ply searchable
    HTCAPTION = 2                       'API
    WM_NCLBUTTONDOWN = 161              'API
    CS_DROPSHADOW = &H20000             'API
    GCL_STYLE = -26                     'API
    OneStep = 795                       'displayed board raster pitch in twips
    TwoSteps = OneStep * 2
    ThreeSteps = TwoSteps + OneStep
    FourSteps = TwoSteps * 2
    FiveSteps = FourSteps + OneStep
End Enum
#If False Then ':) Line inserted by Formatter
Private Infinity, DepthLimit, HTCAPTION, WM_NCLBUTTONDOWN, CS_DROPSHADOW, GCL_STYLE, OneStep, TwoSteps, ThreeSteps, FourSteps, FiveSteps ':) Line inserted by Formatter
#End If ':) Line inserted by Formatter

Private PV(0 To DepthLimit, _
           0 To DepthLimit) As Long     'Principal Variation

Private TimeElapsed         As Single   'during CPU speed assessment
Private HSCFreq             As Currency
Private TenMicrosecs        As Currency 'HscCount for 10 µsecs
Private Result              As Long     'The Search Result
Private Posns               As Long     '# of nodes in search tree visited
Private Cutoffs             As Long     '# of cutoffs
Private TimeStart           As Long
Private TimeUp              As Long     'in millisecs
Private IterDepth           As Long     'controls iterative deepening
Private IndexW              As Long     'controls the white drop-chips
Private IndexB              As Long     'same for black
Private Origin              As tPoint   'top left corner of matrix
Private Half                As Long     'half width of chip image
Private CurrColumn          As Long     'the current position of the mouse cursor
Private Tooltip             As cTooltip
Private LastCtl             As Control
Private InPV                As Boolean  'true if search is in PV
Private SearchInProgress    As Boolean  'what it says
Private Replay              As Boolean  'true when replaying
Private Const White         As Byte = 1 'black is 0
Private Const Signature     As String = "TCM"

Private Sub btExit_Click()

    Unload Me

End Sub

Private Sub btExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip btExit

End Sub

Private Sub btReset_Click()

  Dim i As Long
  Dim j As Long

    With Board

        For i = 0 To 5 'reset the board
            .SNW(i) = 0
            .SNB(i) = 0
            .WEW(i) = 0
            .WEB(i) = 0
            .SwNeB(i) = 0
            .SwNeW(i) = 0
            .SeNwB(i) = 0
            .SeNwW(i) = 0
            .Filled(i) = 0
            imgWhite(i).Visible = True
            imgBlack(i).Visible = False
        Next i
        .SNW(i) = 0 'one more for horizontal
        .SNB(i) = 0
        .Filled(i) = 0
        lnQuad.Visible = False
        imgWhite(i).Visible = True
        imgBlack(i).Visible = False
        lbHint.Enabled = True
        lbTakeBack.Enabled = True

        For i = 1 To 21
            imgDropW(i).Visible = False
            imgDropB(i).Visible = False
        Next i
        .SideToMove = White

    End With 'BOARD
    For i = 0 To DepthLimit
        For j = 0 To DepthLimit
            PV(i, j) = -1
    Next j, i
    lbWins.ForeColor = vbBlack
    IndexW = 0
    IndexB = 0

    lbWins = vbNullString
    imgSmile.Visible = False
    lbWinIn = vbNullString
    CurrColumn = 2
    Form_KeyDown vbKeyRight, 0
    pic.SetFocus

End Sub

Private Sub btReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip btReset

End Sub

Private Function CheckForQuad(ByVal Bits As Long) As Boolean

    CheckForQuad = ((Bits And 15) = 15) Or ((Bits And 30) = 30) Or ((Bits And 60) = 60)
    If CheckForQuad Then
        RemoveButtons
    End If

End Function

Private Function CheckSeNwB(Board As tBoard, ByVal Which As Long) As Boolean

    CheckSeNwB = CheckForQuad(Board.SeNwB(Which))

End Function

Private Function CheckSeNwW(Board As tBoard, ByVal Which As Long) As Boolean

    CheckSeNwW = CheckForQuad(Board.SeNwW(Which))

End Function

Private Function CheckSNB(Board As tBoard, ByVal Which As Long) As Boolean

    CheckSNB = CheckForQuad(Board.SNB(Which))

End Function

Private Function CheckSNW(Board As tBoard, ByVal Which As Long) As Boolean

    CheckSNW = CheckForQuad(Board.SNW(Which))

End Function

Private Function CheckSwNeB(Board As tBoard, ByVal Which As Long) As Boolean

    CheckSwNeB = CheckForQuad(Board.SwNeB(Which))

End Function

Private Function CheckSwNeW(Board As tBoard, ByVal Which As Long) As Boolean

    CheckSwNeW = CheckForQuad(Board.SwNeW(Which))

End Function

Private Function CheckWEB(Board As tBoard, ByVal Which As Long) As Boolean

  Dim Bits  As Long

    Bits = Board.WEB(Which)
    CheckWEB = CheckForQuad(Bits) Or (Bits And 120) = 120 'one extra check because WE has 7 bits
    If CheckWEB Then
        RemoveButtons
    End If

End Function

Private Function CheckWEW(Board As tBoard, ByVal Which As Long) As Boolean

  Dim Bits  As Long

    Bits = Board.WEW(Which)
    CheckWEW = CheckForQuad(Bits) Or (Bits And 120) = 120 'one extra check because WE has 7 bits
    If CheckWEW Then
        RemoveButtons
    End If

End Function

Private Sub ComputerMove()

  Dim CMResult  As Long
  Dim i         As Long
  Dim Col       As Long

    If Replay Then
        Col = Val(imgDropB(IndexB + 1).Tag)
        MoveCursor Col
        Wait 2000
        Flash imgBlack(Col)
        IterDepth = 1
        imgBlack_Click (Col)
        IterDepth = 0
      Else 'REPLAY = FALSE/0
        Screen.MousePointer = vbHourglass
        Enabled = False
        lbWins = "Thinking..."
        imgThink.Visible = True
        DoEvents

        TimeStart = GetTickCount
        InPV = False
        Posns = 0
        Cutoffs = 0

        SearchInProgress = True
        Do 'iterative search deepening
            IterDepth = IterDepth + 1
            CMResult = Search(Board, 0, IterDepth, -Infinity, Infinity)
            InPV = True 'the search will have returned a principal variation
            DoEvents
        Loop While GetTickCount < (TimeStart + TimeUp) And IterDepth < DepthLimit And Abs(CMResult) = 0
        SearchInProgress = False

        imgThink.Visible = False
        If PV(0, 0) < 0 Then 'no move for black
            For i = 0 To 6
                imgBlack(i).Visible = False
                imgWhite(i).Visible = False
            Next i
            lbHint.Enabled = False
            lbHint.BorderStyle = 0
            lbWins.ForeColor = vbRed
            lbWins = "Red has lost"
            MessageBeep vbInformation
          Else 'NOT PV(0,...
            i = Infinity - CMResult
            If i Then
                lbWinIn = IIf(CMResult > 0 And (CMResult And 1), "Red wins in " & IIf(i > 2, "max ", vbNullString) & i & " moves", vbNullString)
              Else 'I = FALSE/0
                lbWinIn = vbNullString
            End If
            lbWins = "P=" & Format$(Posns, "#,0") & " C=" & Format$(Cutoffs, "#,0") & " D=" & IterDepth & " T=" & Format$((GetTickCount - TimeStart) / 1000, "#0.000")
            Flash imgBlack(PV(0, 0))
            imgBlack_Click (PV(0, 0))

            If (CMResult And 1) = 0 And (IndexW + IndexB = 42) Then
                lbWins = "Draw"
                MessageBeep vbInformation
                lbHint.Enabled = False
            End If
        End If
        IterDepth = 0
        Screen.MousePointer = vbNormal
        If lbHint.BorderStyle = 1 Then
            If CMResult > 0 Then
                SetCursorPos (Left + lbHint.Left + lbHint.Width / 2) / 15, (Top + lbHint.Top + lbHint.Height / 2) / 15
                lbHint.Enabled = False
                lbHint.BorderStyle = 0
              Else 'NOT CMRESULT...
                DoHint PV(0, 1)
            End If
        End If
    End If
    Enabled = True

End Sub

Private Sub CreateTooltip(Ctl As Control)

    If Not LastCtl Is Ctl Then
        On Error Resume Next
            LastCtl.ToolTipText = Tooltip.InitialText
        On Error GoTo 0
        Set LastCtl = Ctl
        Set Tooltip = New cTooltip
        If Tooltip.Create(Ctl, Ctl.ToolTipText, , , TTIconInfo, " ", &HFFFFC0, &H809020, 500) Then 'created tooltip?
            Tooltip.SubstituteFont "Arial", 9, True, True
        End If
        Ctl.ToolTipText = vbNullString      '...and erase it so we don't get two tips
    End If

End Sub

Private Sub DoHint(Idx As Long)

    Flash imgWhite(PV(0, 1)), 8
    MoveCursor PV(0, 1)

End Sub

Private Sub Flash(Chip As Image, Optional ByVal HowOften As Long = 7)

  Dim i As Long

    For i = 1 To HowOften
        Chip.Visible = Not Chip.Visible
        Wait 600
    Next i

End Sub

Private Sub Form_Initialize() ':) Line inserted by Formatter

    InitCommonControls ':) Line inserted by Formatter

End Sub ':) Line inserted by Formatter

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim i As Long
  Dim j As Long

    If Board.SideToMove = White Then
        Select Case KeyCode
          Case vbKeyMenu, vbKeyControl ' alt key, ctl key
            'do nothing
          Case vbKeyX
            btExit_Click
          Case vbKeyN
            btReset_Click
          Case vbKeyReturn, vbKeyDown
            imgWhite_Click CInt(CurrColumn)
            If imgWhite(CurrColumn).Visible = False Then 'not visible after move
                Form_KeyDown vbKeyRight, 0 'recursion to re-position cursor
            End If
          Case Else
            Select Case KeyCode
              Case vbKeyLeft
                j = 1
              Case vbKeyUp
                j = 2
              Case vbKeyRight
                j = 3
              Case vbKey1 To vbKey7 'numeric key
                If imgWhite((KeyCode And 7) - 1).Visible Then
                    CurrColumn = (KeyCode And 7) - 1
                    Form_KeyDown vbKeyReturn, 0
                  Else 'NOT IMGWHITE((KEYCODE...
                    Beeper 3000, 30
                End If
              Case Else
                Beeper 300, 50
            End Select
            If j Then
                For i = 1 To 7
                    CurrColumn = CurrColumn + j - 2
                    Select Case CurrColumn
                      Case Is < 0
                        CurrColumn = 6
                      Case Is > 6
                        CurrColumn = 0
                    End Select
                    If imgWhite(CurrColumn).Visible Then
                        MoveCursor CurrColumn
                        Exit For 'loop varying i
                    End If
                Next i
            End If
        End Select
    End If

End Sub

Private Sub Form_Load()

  Dim i         As Long
  Dim Username  As String
  Dim tmpBoard  As tBoard
  Dim From      As Currency
  Dim Till      As Currency

    SetClassLong hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Or CS_DROPSHADOW

    'evaluate CPU speed
    DoEvents
    QueryPerformanceCounter From
    For i = 1 To 10970
        CopyBoard tmpBoard, Board
        Search tmpBoard, 0, 1, -Infinity, Infinity
    Next i
    QueryPerformanceCounter Till
    Till = Till - From
    QueryPerformanceFrequency HSCFreq
    TimeElapsed = Till / HSCFreq
    'comes out as 0.1 secs on my CPU - which I now proudly declare the 100 % standard, the reason being that I
    'have no other and it doesn't matter anyway because if your CPU is faster she's just given less time to think

    TenMicrosecs = HSCFreq / 10000
    For i = 1 To 21             'create the chips for dropping
        Load imgDropW(i)
        Load imgDropB(i)
    Next i
    lbTitle = Caption
    Half = imgWhite(0).Width / 2
    Origin.X = imgWhite(0).Left + Half
    Origin.Y = imgWhite(0).Top + 1020 + Half
    Half = Half - 15

    'ordered moves, order is 3, 2, 4, 5, 1, 0, 6
    OrdMovs(0) = 3              'for move ordering
    OrdMovs(1) = 2
    OrdMovs(2) = 4
    OrdMovs(3) = 5
    OrdMovs(4) = 1
    OrdMovs(5) = 0
    OrdMovs(6) = 6

    lbOpt_MouseDown 0, 0, 0, 0, 0 'beginner

    Set Tooltip = New cTooltip    'instantiate the tooltip class
    Username = String$(256, 0)
    i = 256
    GetUserName Username, i
    Username = Left(Username, i - 1)
    lbExit.ToolTipText = "Good bye, " & Username
    btExit.ToolTipText = lbExit.ToolTipText
    With App
        lbAuthor = .LegalCopyright
        lbVersion = "Version " & .Major & "." & .Minor & "." & .Revision & "      CPU " & Format$(10 / TimeElapsed, "#0.0\%")
        pic.ToolTipText = Caption & " - " & lbVersion & "||" & lbAuthor
        If .PrevInstance Then
            MsgBox "One Instance only, please!", vbCritical, Caption
            Unload Me
          Else '.PREVINSTANCE = FALSE/0
            Show
            DoEvents
            btReset_Click
        End If
    End With 'APP

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MousePointer = vbSizeAll
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0& 'grab form
    Form_Resize 'to repaint the frame

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MousePointer = vbDefault

    If Not LastCtl Is Nothing Then 'reset control's ttt and kill tooltip object
        LastCtl.ToolTipText = Tooltip.InitialText
        Set LastCtl = Nothing
        Set Tooltip = Nothing
    End If

End Sub

Private Sub Form_Resize()

  Dim Colr  As Long
  Dim i     As Long

    If WindowState = vbNormal Then
        DoEvents

        'draw a frame
        SetRect WindowRect, 0, 0, ScaleX(Width, ScaleMode, vbPixels), ScaleY(Height, ScaleMode, vbPixels)
        For i = 0 To 255 Step 22
            Colr = 255 - Abs(128 - i)
            ForeColor = RGB(Colr - 64, Colr, Colr - 32)
            With WindowRect
                Rectangle hDC, .LT.X, .LT.Y, .RB.X, .RB.Y
            End With 'WINDOWRECT
            InflateRect WindowRect, -1, -1
        Next i
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Tooltip = Nothing

End Sub

Private Sub imgBlack_Click(Index As Integer)

  Dim i As Long

    imgBlack(Index).Visible = False
    IndexB = IndexB + 1
    imgDropB(IndexB).Left = imgBlack(Index).Left
    imgDropB(IndexB).Tag = Index
    Slide imgDropB(IndexB), 5 - Board.Filled(Index)
    MakeMove Board, Index
    For i = 0 To 6
        imgWhite(i).Visible = (Board.Filled(i) < 6)
        imgBlack(i).Visible = False
    Next i
    Board.SideToMove = 1 - Board.SideToMove
    If IterDepth = 0 Then
        ComputerMove
    End If
    GetCursorPos CP
    SetCursorPos CP.X + 1, CP.Y 'move a little to reset cursor to hand
    DoEvents
    SetCursorPos CP.X, CP.Y

End Sub

Private Sub imgBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub imgBoard_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseMove Button, Shift, X, Y

End Sub

Private Sub imgIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub imgIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseMove Button, Shift, X, Y

End Sub

Private Sub imgUMGEDV_Click()

    Clipboard.SetText "UMGEDV@Yahoo.com"

End Sub

Private Sub imgUMGEDV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip imgUMGEDV

End Sub

Private Sub imgWhite_Click(Index As Integer)

  Dim i As Long

    If imgWhite(Index).Visible Then
        imgWhite(Index).Visible = False
        IndexW = IndexW + 1
        imgDropW(IndexW).Left = imgWhite(Index).Left
        imgDropW(IndexW).Tag = Index
        Slide imgDropW(IndexW), 5 - Board.Filled(Index)
        MakeMove Board, Index
        For i = 0 To 6
            imgBlack(i).Visible = (Board.Filled(i) < 6)
            imgWhite(i).Visible = False
        Next i
        Board.SideToMove = 1 - Board.SideToMove
        If IterDepth = 0 Then
            ComputerMove
        End If
        GetCursorPos CP
        SetCursorPos CP.X + 1, CP.Y 'move a little to reset cursor to hand
        DoEvents
        SetCursorPos CP.X, CP.Y
    End If

End Sub

Private Sub imgWhite_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    CurrColumn = Index

End Sub

Private Sub lbExit_Click()

    Unload Me

End Sub

Private Sub lbExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip lbExit

End Sub

Private Sub lbHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip lbHelp

End Sub

Private Sub lbHint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'this works like a pen - latch on first Down and release on second Up

    If lbHint.Tag = vbNullString Then
        lbHint.BorderStyle = 1
        If lbWinIn = vbNullString Then
            If PV(0, 1) > 0 Then
                DoHint PV(0, 1)
            End If
        End If
    End If
    lbHint.Tag = lbHint.Tag & "A"

End Sub

Private Sub lbHint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip lbHint

End Sub

Private Sub lbHint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If lbHint.Tag = "AA" Then
        lbHint.Tag = vbNullString
        lbHint.BorderStyle = 0
    End If

End Sub

Private Sub lbLoad_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lbLoad.BorderStyle = 1

End Sub

Private Sub lbLoad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip lbLoad

End Sub

Private Sub lbLoad_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim Moves As String
  Dim i     As Long

    lbLoad.BorderStyle = 0
    With cDlg
        .InitDir = App.Path & "\Games"
        .DialogTitle = "Enter/Select file to load from..."
        .FileName = vbNullString
        .DefaultExt = ".cn4"
        .Filter = "Connect4(*.CN4)|*.CN4|All Files(*.*)|*.*"
        .Flags = cdlOFNPathMustExist Or cdlOFNLongNames
        On Error Resume Next
            .ShowOpen
            If Err = 0 Then
                i = FreeFile
                Open .FileName For Input As i
                Moves = Input(LOF(i), i)
                Close i
            End If
        On Error GoTo 0
        If Len(Moves) > 2 Then
            If Left$(Moves, Len(Signature)) = Signature Then
                Moves = Mid$(Moves, Len(Signature) + 1)
                IndexB = 0
                For i = 0 To Len(Moves) Step 2
                    If Not IsNumeric(Mid$(Moves, i + 1, 1)) Then
                        Exit For 'loop varying i
                    End If
                    imgDropW(i / 2 + 1).Tag = Mid$(Moves, i + 1, 1)
                    imgDropB(i / 2 + 1).Tag = Mid$(Moves, i + 2, 1)
                    IndexB = IndexB + 1
                Next i
                lbWinIn = Mid$(Moves, i + 1)
                lbReplay_MouseUp 0, 0, 0, 0
              Else 'NOT LEFT$(MOVES,...
                MsgBox Replace$(.FileName, "\", " > ") & vbCrLf & vbCrLf & "is not a valid " & Caption & " File", vbCritical, Caption
            End If
        End If
    End With 'CDLG

End Sub

Private Sub lbMini_Click()

    WindowState = vbMinimized

End Sub

Private Sub lbMini_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip lbMini

End Sub

Private Sub lbOpt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'TimeUp works as follows: The program repeats Searches with increasing depth limits and when a Search
  'returns and finds that time is up it will stop iterating the Search. Since shallow Searches are quite fast
  'and normally return with a principal variation set up (which the next iteration will use to trigger
  'cutoffs early, making the deeper Searches faster) the net result is more Search plies within a given time
  'than without iteration. Also TimeUp does not interupt a Search while it is under way, making Search control
  'easier.

  Dim i As Long

    For i = 0 To 3
        lbOpt(i).BorderStyle = Abs(i = Index)
    Next i
    TimeUp = ((Index / 2 + 0.1) * TimeElapsed * 1005) 'bigger TimeElapsed means slower CPU - so give her more time, and vice versa
    TimeUp = TimeUp * 10
    lbTimeCheck = "Time Up" & vbCrLf & Round(TimeUp, 0) & " mSecs"

End Sub

Private Sub lbOpt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip lbOpt(Index)

End Sub

Private Sub lbReplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lbReplay.BorderStyle = 1

End Sub

Private Sub lbReplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip lbReplay

End Sub

Private Sub lbReplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim HowMany   As Long
  Dim Col       As Long
  Dim tmpWinIn  As String
  Const RPiP    As String = "Replay in progress"

    If IndexB = 0 Then
        lbWins = "Nothing  to replay"
        MessageBeep vbCritical
      Else 'NOT INDEXB...
        HowMany = IndexB
        tmpWinIn = lbWinIn
        btReset_Click
        Enabled = False
        Replay = True
        lbWins = RPiP
        DoEvents
        Do
            Col = Val((imgDropW(IndexW + 1).Tag))
            MoveCursor Col
            Wait 2000
            Flash imgWhite(Col), 8
            imgWhite_Click (Col)
        Loop Until IndexB = HowMany
        Replay = False
        If lbWins = RPiP Then
            lbWins = vbNullString
        End If
        Enabled = True
        lbWinIn = tmpWinIn
    End If
    lbReplay.BorderStyle = 0

End Sub

Private Sub lbSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lbSave.BorderStyle = 1

End Sub

Private Sub lbSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip lbSave

End Sub

Private Sub lbSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim i     As Long
  Dim hFile As Long

    If IndexB Then
        lbWins.ForeColor = vbBlack
        imgSmile.Visible = False
        With cDlg
            .InitDir = App.Path & "\Games"
            .DialogTitle = "Enter/Select file to save to..."
            .FileName = vbNullString
            .DefaultExt = ".cn4"
            .Filter = "Connect4(*.CN4)|*.CN4|All Files(*.*)|*.*"
            .Flags = cdlOFNLongNames Or cdlOFNOverwritePrompt
            On Error Resume Next
                .ShowSave
                If Err = 0 Then
                    hFile = FreeFile
                    Open .FileName For Output As hFile
                    Print #hFile, Signature;
                    For i = 0 To IndexB
                        Print #hFile, imgDropW(i).Tag; imgDropB(i).Tag;
                    Next i
                    Print #hFile, lbWinIn;
                    Close hFile
                    lbWins = "Game Status saved"
                  Else 'NOT ERR...
                    lbWins = "Not saved"
                End If
            On Error GoTo 0
        End With 'CDLG
    End If
    lbSave.BorderStyle = 0

End Sub

Private Sub lbTakeBack_Click()

  Dim i     As Long
  Dim Col   As Long
  Dim Row   As Long
  Dim Bit   As Long

    lbWins = vbNullString
    lbWinIn = vbNullString

    If IndexB + IndexW Then
        With Board

            Col = Val(imgDropB(IndexB).Tag)
            Slide imgDropB(IndexB), -1
            IndexB = IndexB - 1

            .Filled(Col) = .Filled(Col) - 1
            Row = .Filled(Col)
            .WEB(Row) = .WEB(Row) And Not 2 ^ Col
            Bit = 2 ^ Row
            .SNB(Col) = .SNB(Col) And Not Bit
            i = Col - Row + 2
            .SwNeB(i) = .SwNeB(i) And Not Bit
            i = Row + Col - 3
            .SeNwB(i) = .SeNwB(i) And Not Bit

            Col = Val(imgDropW(IndexW).Tag)
            Slide imgDropW(IndexW), -1
            IndexW = IndexW - 1

            .Filled(Col) = .Filled(Col) - 1
            Row = .Filled(Col)
            .WEW(Row) = .WEW(Row) And Not 2 ^ Col
            Bit = 2 ^ Row
            .SNW(Col) = .SNW(Col) And Not Bit
            i = Col - Row + 2
            .SwNeW(i) = .SwNeW(i) And Not Bit
            i = Row + Col - 3
            .SeNwW(i) = .SeNwW(i) And Not Bit

            For i = 0 To 6
                imgWhite(i).Visible = (.Filled(i) < 6)
                imgBlack(i).Visible = False
            Next i

        End With 'BOARD

        lbWins = "Two moves have been taken back"
      Else 'NOT INDEXB...
        lbWins = "No moves to take back"
        MessageBeep vbCritical
    End If
    lbTakeBack.BorderStyle = 0

End Sub

Private Sub lbTakeBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lbTakeBack.BorderStyle = 1

End Sub

Private Sub lbTakeBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip lbTakeBack

End Sub

Private Sub lbTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub lbTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseMove Button, Shift, X, Y

End Sub

Private Function MakeMove(Board As tBoard, ByVal Column As Long) As Boolean

  'column is 0 to 6
  'color is white or black (corresponding to yellow and red pieces)

  Dim Bit   As Long
  Dim Row   As Long
  Dim i     As Long

    With Board
        Row = .Filled(Column)
        If .SideToMove = White Then
            .WEW(Row) = .WEW(Row) Or 2 ^ Column
            If CheckWEW(Board, Row) = False Then
                Bit = 2 ^ Row
                .SNW(Column) = .SNW(Column) Or Bit
                If CheckSNW(Board, Column) = False Then
                    i = Column - Row + 2
                    .SwNeW(i) = .SwNeW(i) Or Bit
                    If CheckSwNeW(Board, i) = False Then
                        i = Row + Column - 3
                        .SeNwW(i) = .SeNwW(i) Or Bit
                        MakeMove = CheckSeNwW(Board, i)
                      Else 'NOT CHECKSWNEW(BOARD,...
                        MakeMove = True
                    End If
                  Else 'NOT CHECKSNW(BOARD,...
                    MakeMove = True
                End If
              Else 'NOT CHECKWEW(BOARD,...
                MakeMove = True
            End If
          Else 'NOT .SIDETOMOVE...
            .WEB(Row) = .WEB(Row) Or 2 ^ Column
            If CheckWEB(Board, Row) = False Then
                Bit = 2 ^ Row
                .SNB(Column) = .SNB(Column) Or Bit
                If CheckSNB(Board, Column) = False Then
                    i = Column - Row + 2
                    .SwNeB(i) = .SwNeB(i) Or Bit
                    If CheckSwNeB(Board, i) = False Then
                        i = Row + Column - 3
                        .SeNwB(i) = .SeNwB(i) Or Bit
                        MakeMove = CheckSeNwB(Board, i)
                      Else 'NOT CHECKSWNEB(BOARD,...
                        MakeMove = True
                    End If
                  Else 'NOT CHECKSNB(BOARD,...
                    MakeMove = True
                End If
              Else 'NOT CHECKWEB(BOARD,...
                MakeMove = True
            End If
        End If
        .Filled(Column) = .Filled(Column) + 1
    End With 'BOARD

End Function

Private Sub MarkQuad()

  'finds the completed quad and marks it

  Dim Row       As Long
  Dim Col       As Long
  Dim cx        As Long
  Dim cy        As Long
  Dim Bitmap    As Long

    cx = imgDropB(IndexB).Left + Half
    cy = imgDropB(IndexB).Top + Half
    Col = (cx - Origin.X) / OneStep
    Row = 5 - (cy - Origin.Y) / OneStep

    With lnQuad
        .ZOrder 0

        'west --> east
        Bitmap = Board.WEB(Row) '15, 30, 60, 120
        If (Bitmap And 15) = 15 Or (Bitmap And 30) = 30 Or (Bitmap And 60) = 60 Or (Bitmap And 120) = 120 Then
            .Y1 = cy
            .Y2 = cy
            Select Case True
              Case (Bitmap And 15) = 15
                .X1 = Origin.X
              Case (Bitmap And 30) = 30
                .X1 = Origin.X + OneStep
              Case (Bitmap And 60) = 60
                .X1 = Origin.X + TwoSteps
              Case (Bitmap And 120) = 120
                .X1 = Origin.X + ThreeSteps
            End Select
            .X2 = .X1 + ThreeSteps
            .Visible = True
        End If

        'sout --> north
        Bitmap = Board.SNB(Col) '15, 30, 60
        If (Bitmap And 15) = 15 Or (Bitmap And 30) = 30 Or (Bitmap And 60) = 60 Then
            .X1 = cx    'because it can only be the top chip
            .X2 = cx
            .Y1 = cy
            .Y2 = cy + ThreeSteps
            .Visible = True
        End If

        'southwest --> northeast
        Bitmap = Board.SwNeB(Col - Row + 2) '15, 30, 60
        If (Bitmap And 15) = 15 Or (Bitmap And 30) = 30 Or (Bitmap And 60) = 60 Then
            Select Case True
              Case (Bitmap And 15) = 15
                .X1 = ThreeSteps
                .Y1 = TwoSteps
              Case (Bitmap And 30) = 30
                .X1 = FourSteps
                .Y1 = OneStep
              Case (Bitmap And 60) = 60
                .X1 = FiveSteps
                .Y1 = 0
            End Select
            .X1 = .X1 + Origin.X + (Col - Row) * OneStep
            .X2 = .X1 - ThreeSteps
            .Y1 = .Y1 + Origin.Y
            .Y2 = .Y1 + ThreeSteps
            .Visible = True
        End If

        'southeast --> northwest
        Bitmap = Board.SeNwB(Row + Col - 3) '15, 30, 60
        If (Bitmap And 15) = 15 Or (Bitmap And 30) = 30 Or (Bitmap And 60) = 60 Then
            Select Case True
              Case (Bitmap And 15) = 15
                .X1 = -ThreeSteps
                .Y1 = TwoSteps
              Case (Bitmap And 30) = 30
                .X1 = -FourSteps
                .Y1 = OneStep
              Case (Bitmap And 60) = 60
                .X1 = -FiveSteps
                .Y1 = 0
            End Select
            .X1 = .X1 + Origin.X + (Row + Col) * OneStep
            .X2 = .X1 + ThreeSteps
            .Y1 = .Y1 + Origin.Y
            .Y2 = .Y1 + ThreeSteps
            .Visible = True
        End If

    End With 'LNQUAD

End Sub

Private Sub MoveCursor(ByVal Col As Long)

    SetCursorPos (Left + imgWhite(Col).Left + Half) / 15 + 1, (Top + imgWhite(0).Top + Half) / 15
    DoEvents
    SetCursorPos (Left + imgWhite(Col).Left + Half) / 15, (Top + imgWhite(0).Top + Half) / 15

End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CreateTooltip pic

End Sub

Private Sub RemoveButtons()

  Dim i As Long

    If SearchInProgress = False Then 'not searching
        For i = 0 To 6
            Board.Filled(i) = 6
        Next i
        lbWins.ForeColor = vbRed
        MessageBeep vbExclamation
        lbWins = "Red wins"
        imgSmile.Visible = True
        MarkQuad
        lbHint.Enabled = False
        lbTakeBack.Enabled = False
    End If

End Sub

Private Function RunningInIDE(Optional c As Boolean = False) As Boolean

  Static b  As Boolean

    b = c
    If b = False Then
        Debug.Assert RunningInIDE(True)
    End If
    RunningInIDE = b

End Function

Private Function Search(Board As tBoard, ByVal Depth As Long, ByVal MinDepth As Long, ByVal Alpha As Long, ByVal Beta As Long) As Long

  'negamax search

  Dim TempBoard     As tBoard
  Dim OMs(0 To 6)   As Long 'ordered moves, order is 3, 2, 4, 5, 1, 0, 6
  Dim Move          As Long                       '  0, 1, 2, 3, 4, 5, 6
  Dim i             As Long

    If Depth < MinDepth And Depth <> DepthLimit Then         'min depth not yet searched and not at depth limit
        Posns = Posns + 1               'count positions

        CopyMoves OMs(0), OrdMovs(0)    'get default move order

        If InPV Then
            If PV(0, Depth) >= 0 Then   'there is a PV move for this depth

                'move ordering
                OMs(0) = PV(0, Depth)   'analyse the most promising move (the PV move) first
                '                        this ensures that a: the Principal Variation is established as soon as possible or
                '                                          b: cutoff is triggered as early as possible
                Select Case OMs(0)
                  Case 3
                    'do nothing
                  Case 2                'OMs(0) was 3 and is now 2
                    OMs(1) = 3          'so 3 has to go where 2 was
                  Case 4                'analog to above
                    OMs(2) = 3
                  Case 5                'analog to above
                    OMs(3) = 3
                  Case 1                'analog to above
                    OMs(4) = 3
                  Case 0                'analog to above
                    OMs(5) = 3
                  Case Else '6          'analog to above
                    OMs(6) = 3
                End Select

              Else 'NOT PV(0,...
                InPV = False            'is out of PV now
            End If
        End If

        PV(Depth, Depth) = -1           'clear Principal Variation at this depth

        Search = Depth - Infinity       'depth added to find shortest winning line

        With Board

            For Move = 0 To 6                                   'the 7 possible moves

                If .Filled(OMs(Move)) < 6 Then                  'column not yet filled
                    CopyBoard TempBoard, Board

                    If MakeMove(TempBoard, OMs(Move)) Then      'this created a quad
                        Result = Infinity - Depth
                      Else 'no quad created 'NOT MAKEMOVE(TEMPBOARD,...

                        With TempBoard
                            .SideToMove = White - .SideToMove                               'toggle side to move
                            Result = -Search(TempBoard, Depth + 1, MinDepth, -Beta, -Alpha) 'minimax (negamax) recursion
                            .SideToMove = White - .SideToMove                               'back to this side to move
                        End With 'TEMPBOARD

                    End If 'move created quad

                    If Result > Search Then                     'this was a better move
                        Search = Result
                        For i = Depth + 1 To DepthLimit - 1     'create principal variation
                            PV(Depth, i) = PV(Depth + 1, i)     'by copying all best moves after this best move to current depth
                            If PV(Depth, i) = -1 Then           'end of PV
                                Exit For 'loop varying i
                            End If
                        Next i
                        PV(Depth, Depth) = OMs(Move)            'and enter this best move into PV

                        Select Case Search
                          Case Is >= Beta
                            Cutoffs = Cutoffs + 1
                            Exit For        'this is bad enough - dont wanna know if there are any worse moves 'loop varying move
                          Case Is > Alpha
                            Alpha = Search  'limit low
                        End Select

                    End If 'better move
                End If 'column filled

            Next Move

        End With 'BOARD
    End If

End Function

Private Sub Slide(Chip As Image, ByVal HowDeep As Long)

  'drops chip into a slot and bounces a little or sildes it back out

  Dim i As Long
  Dim j As Single

    j = 16

    With Chip

        If HowDeep < 0 Then
            'lift it
            Do
                Wait 12
                .Top = .Top - 15
            Loop Until .Top = 900
            .Visible = False
          Else 'NOT HOWDEEP...
            'position it and make it visible
            .Top = 900
            .Visible = True
            .ZOrder 0

            'drop it
            Do
                Wait 8
                .Top = .Top + 15
            Loop Until .Top = 1710 + OneStep * HowDeep

            'bounce it
            Do
                For i = 1 To j
                    Wait 40
                    .Top = .Top - 15
                Next i
                For i = 1 To j
                    Wait 40
                    .Top = .Top + 15
                Next i
                j = j / 1.4
            Loop Until j < 3
        End If
    End With 'CHIP
    DoEvents

End Sub

Private Sub Wait(ByVal Interval As Long) 'one interval is 10 µsecs

  Dim From  As Currency
  Dim Till  As Currency

    QueryPerformanceCounter From
    Till = From + Interval * TenMicrosecs
    Do
        DoEvents
        QueryPerformanceCounter From
    Loop While From < Till

End Sub

':) Ulli's VB Code Formatter V2.24.21 (2009-Apr-01 09:11)  Decl: 160  Code: 1204  Total: 1364 Lines
':) CommentOnly: 88 (6,5%)  Commented: 141 (10,3%)  Filled: 1068 (78,3%)  Empty: 296 (21,7%)  Max Logic Depth: 8
