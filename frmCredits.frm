VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H0044342E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Odyssey Realms [Credits]"
   ClientHeight    =   9000
   ClientLeft      =   105
   ClientTop       =   330
   ClientWidth     =   9015
   ControlBox      =   0   'False
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblOptima 
      BackColor       =   &H0044342E&
      Caption         =   "Optima"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   38
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lblVelius 
      BackColor       =   &H0044342E&
      Caption         =   "Velius"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   5280
      Width           =   615
   End
   Begin VB.Label lblClassicQuests 
      BackColor       =   &H0044342E&
      Caption         =   "Classic Quests"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   36
      Top             =   4680
      Width           =   1965
   End
   Begin VB.Label lblDeadKitty 
      BackColor       =   &H0044342E&
      Caption         =   "Dead Kitty Registry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   35
      Top             =   3360
      Width           =   2445
   End
   Begin VB.Label lblBorfshwitz 
      BackColor       =   &H0044342E&
      Caption         =   "Patrick Bukowski (Borfshwitz)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label lblUnknown 
      BackColor       =   &H0044342E&
      Caption         =   "Raymond Cox (Unknown)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   33
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label lblRebornSource 
      BackColor       =   &H0044342E&
      Caption         =   "Reborn Source"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   32
      Top             =   1680
      Width           =   2325
   End
   Begin VB.Label lblTheZeus 
      BackColor       =   &H0044342E&
      Caption         =   "The Zeus"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   31
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblGolbez 
      BackColor       =   &H0044342E&
      Caption         =   "Golbez"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   30
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblRza 
      BackColor       =   &H0044342E&
      Caption         =   "Jason McDermott (Rza)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   29
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label lblSlasher 
      BackColor       =   &H0044342E&
      Caption         =   "Jeremy McDermott (Slasher)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   28
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label lblOdysseyRealms 
      BackColor       =   &H0044342E&
      Caption         =   "Odyssey Realms"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   3120
      TabIndex        =   27
      Top             =   4680
      Width           =   2085
   End
   Begin VB.Label lblPure 
      BackColor       =   &H0044342E&
      Caption         =   "Jesse Gottschalk (Pure)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   26
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblPengwy 
      BackColor       =   &H0044342E&
      Caption         =   "Dave Tu (Pengwy)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   25
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblSteverino 
      BackColor       =   &H0044342E&
      Caption         =   "Steve Harris (Steverino)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   24
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label lblMarco 
      BackColor       =   &H0044342E&
      Caption         =   "Marco Pelloni (Captain Marco)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   23
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label lblJames 
      BackColor       =   &H0044342E&
      Caption         =   "James Serine (DarkOne)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   22
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label lblFank 
      BackColor       =   &H0044342E&
      Caption         =   "Fankadore"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   21
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblJaron 
      BackColor       =   &H0044342E&
      Caption         =   "Jaron Leavitt (Llamaboy)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   20
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblProgramming 
      BackColor       =   &H0044342E&
      Caption         =   "Misc Development"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   19
      Top             =   3360
      Width           =   2325
   End
   Begin VB.Label lblA201Source 
      BackColor       =   &H0044342E&
      Caption         =   "A201 Source"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Top             =   1680
      Width           =   1605
   End
   Begin VB.Label lblGecky 
      BackColor       =   &H0044342E&
      Caption         =   "Judy Shmidt (Gecky)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblArchbane 
      BackColor       =   &H0044342E&
      Caption         =   "Greg Dorando (Archbane)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label lblOriginalArt 
      BackColor       =   &H0044342E&
      Caption         =   "Art"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   1605
   End
   Begin VB.Label lblLighting 
      BackColor       =   &H0044342E&
      Caption         =   "Lighting and Weather"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "The Odyssey Online Classic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label lblHistory 
      BackColor       =   &H0044342E&
      Caption         =   $"frmCredits.frx":1CFA
      ForeColor       =   &H009AADC2&
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   8655
   End
   Begin VB.Label lblOriginalGame 
      BackColor       =   &H0044342E&
      Caption         =   "Original Game"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1965
   End
   Begin VB.Label lblBaD 
      BackColor       =   &H0044342E&
      Caption         =   "Christopher Lowenthal (BaD)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   10
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      BackColor       =   &H0044342E&
      Caption         =   "Odyssey Realms Version B4"
      ForeColor       =   &H009AADC2&
      Height          =   195
      Left            =   6720
      TabIndex        =   9
      Top             =   8640
      Width           =   2100
   End
   Begin VB.Label lblThe4On 
      BackColor       =   &H0044342E&
      Caption         =   "Jamie Ryan (The 4on)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   8
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lblJay 
      BackColor       =   &H0044342E&
      Caption         =   "Jay Manley (Xtreme/Carrera)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label lblQBcrusher 
      BackColor       =   &H0044342E&
      Caption         =   "Dante Pellicciotti (QBcrusher)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   6
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label lblRemote 
      BackColor       =   &H0044342E&
      Caption         =   "James Chambers (Remote)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      Top             =   2040
      Width           =   2265
   End
   Begin VB.Label lblVivi 
      BackColor       =   &H0044342E&
      Caption         =   "Vivi"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label lblSmithy 
      BackColor       =   &H0044342E&
      Caption         =   "Clay Rance (Smithy)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label btnOk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   8280
      Width           =   2775
   End
   Begin VB.Label lblBugaboo 
      BackColor       =   &H0044342E&
      Caption         =   "Justin E. Schumacher (Bugaboo)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lblSpecialThanks 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Special thanks to all those who donated their time or money to make this game possible over the years."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   1095
      Left            =   1320
      TabIndex        =   2
      Top             =   6840
      Width           =   6135
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnOk_Click()
    Unload Me
    frmMenu.Show
End Sub

Private Sub Form_Load()
    lblVer.Caption = "Version B" + CStr(ClientVer)
End Sub

