VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H0061514B&
   BorderStyle     =   0  'None
   Caption         =   "Odyssey Realms"
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "More Servers"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   420
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   4260
   End
   Begin VB.Label lblCurrentServer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   2880
      TabIndex        =   0
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Facebook"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   300
      Index           =   14
      Left            =   2280
      TabIndex        =   12
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Youtube"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   300
      Index           =   13
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reddit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   300
      Index           =   12
      Left            =   1200
      TabIndex        =   10
      Top             =   2760
      Width           =   950
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Server List"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   420
      Index           =   9
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Discord"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   300
      Index           =   8
      Left            =   3360
      TabIndex        =   7
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Website"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   420
      Index           =   5
      Left            =   2280
      TabIndex        =   6
      Top             =   1800
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   420
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New Account"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   420
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   420
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   420
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2100
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   420
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   4260
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_GotFocus()
    lblCurrentServer = ServerDescription
End Sub

Private Sub Form_Load()
    gHW = Me.hwnd


    Dim File As String
    Dim FileByteArray() As Byte

    File = GetGfxPath + "menu.rsc"
    FileByteArray() = StrConv(File, vbFromUnicode)
    ReDim Preserve FileByteArray(UBound(FileByteArray) + 1)

    frmMenu.Picture = LoadPicture(File)

    lblCurrentServer = ServerDescription
End Sub

Private Sub lblMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMenu(Index).BackColor = &H61514B
End Sub

Private Sub lblMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim uTopic As String
Dim uFile As String
Dim uParams As String
Dim uDirectory As String
    lblMenu(Index).BackColor = &H44342E
    If X >= 0 And X <= lblMenu(Index).Width And Y >= 0 And Y <= lblMenu(Index).Height Then
        Select Case Index
        Case 0    'Login
            frmLogin.Show
            Me.Hide
        Case 1    'New Account
            frmAccount.Show
            Me.Hide
        Case 2    'Options
            frmOptions.Show
            Me.Hide
        Case 3    'Credits
            frmCredits.Show
            Me.Hide
        Case 4    'Quit
            blnEnd = True
        Case 5    'Website
            uTopic = "Open"
            uFile = TheWebSite
            uParams = 0&
            uDirectory = 0&
            RunShellExecute uTopic, uFile, uParams, uDirectory, 1
        Case 6    'Registry
            uTopic = "Open"
            uFile = "registry.exe"
            uParams = 0&
            uDirectory = 0&
            RunShellExecute uTopic, uFile, uParams, uDirectory, 1
        Case 8    'Discord
            uTopic = "Open"
            uFile = TheDiscord
            uParams = 0&
            uDirectory = 0&
            RunShellExecute uTopic, uFile, uParams, uDirectory, 1
        Case 9    'Server List
            Me.Hide
            Load frmServerList
            frmServerList.Show
            frmServerList.Refresh
        Case 12    'Reddit
            uTopic = "Open"
            uFile = TheSubReddit
            uParams = 0&
            uDirectory = 0&
            RunShellExecute uTopic, uFile, uParams, uDirectory, 1
        Case 13    'Youtube
            uTopic = "Open"
            uFile = TheYoutubeChannel
            uParams = 0&
            uDirectory = 0&
            RunShellExecute uTopic, uFile, uParams, uDirectory, 1
        Case 14    'Facebook Page
            uTopic = "Open"
            uFile = TheFacebookPage
            uParams = 0&
            uDirectory = 0&
            RunShellExecute uTopic, uFile, uParams, uDirectory, 1
        End Select
    End If
End Sub

Sub RunShellExecute(sTopic As String, sFile As Variant, _
                    sParams As Variant, sDirectory As Variant, _
                    nShowCmd As Long)

    Dim hWndDesk As Long
    Dim SUCCESS As Long

    'the desktop will be the
    'default for error messages
    hWndDesk = GetDesktopWindow()

    'execute the passed operation
    SUCCESS = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

    'This is optional. comment the three lines
    'below to not have the "Open With.." dialog appear
    'when the ShellExecute API call fails
    If SUCCESS = 31 Then
        MsgBox "Couldn't load the default application"    'shouldn't happen
        'but if it does, try to get the user to make an association...
        Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
    End If
End Sub
