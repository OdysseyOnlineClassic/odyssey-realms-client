VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H0044342E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Credits]"
   ClientHeight    =   7095
   ClientLeft      =   105
   ClientTop       =   330
   ClientWidth     =   5850
   ControlBox      =   0   'False
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblVer 
      AutoSize        =   -1  'True
      BackColor       =   &H0044342E&
      Caption         =   "Version B1"
      ForeColor       =   &H009AADC2&
      Height          =   195
      Left            =   4920
      TabIndex        =   10
      Top             =   6840
      Width           =   765
   End
   Begin VB.Label lblQBcrusher 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Scripting Assistance by Jamie Ryan (The 4on)"
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
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   5535
   End
   Begin VB.Label lblJay 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Programming and scripting by Jay Manley (Xtreme/Carrera)"
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
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   5535
   End
   Begin VB.Label lblQBcrusher 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Odyssey Realms Registry && Odyssey Classic History Book maintenance by Dante Pellicciotti (QBcrusher)"
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
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   5535
   End
   Begin VB.Label lblRemote 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Source code A201 by James Chambers (Remote)"
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
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Label lblArt 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Art by Greg Dorando (Archbane), Judy Shmidt (Gecky), Vivi"
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Label lblSmithy 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Original lighting and weather effects by Clay Rance (Smithy)"
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
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   5535
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
      Left            =   1560
      TabIndex        =   5
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label lblBugaboo 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Original game by Justin E. Schumacher (Bugaboo)"
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
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label lblSpecialThanks 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Special thanks to all those who donated their time or money to make this game possible."
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
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0044342E&
      Caption         =   "Odyssey Online Classic"
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
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
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
