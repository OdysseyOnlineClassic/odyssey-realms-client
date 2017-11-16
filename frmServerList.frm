VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServerList 
   BackColor       =   &H0061514B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Odyssey - Select Server"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   Icon            =   "frmServerList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sckPing 
      Index           =   6
      Left            =   840
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckPing 
      Index           =   5
      Left            =   720
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckPing 
      Index           =   4
      Left            =   600
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckPing 
      Index           =   3
      Left            =   480
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckPing 
      Index           =   2
      Left            =   360
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckPing 
      Index           =   1
      Left            =   240
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckPing 
      Index           =   0
      Left            =   120
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstServers 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   2730
      ItemData        =   "frmServerList.frx":0E42
      Left            =   120
      List            =   "frmServerList.frx":0E49
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label btnPlay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Play!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   3120
      Width           =   1695
   End
End
Attribute VB_Name = "frmServerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnPlay_Click()
    
    Select Case lstServers.ListIndex

        Case 0 'Classic
            ServerDescription = "Odyssey Realms - Classic"
            CacheDirectory = App.Path + "\classic"
            ServerIP = "libertyarchives.info"
            ServerPort = 5756
        Case 1 'PK Island
            ServerDescription = "PK Island"
            CacheDirectory = App.Path + "\pkisland"
            ServerIP = "libertyarchives.info"
            ServerPort = 5758
        Case 2 'Ethia
            ServerDescription = "Ethia"
            CacheDirectory = App.Path + "\ethia"
            ServerIP = "libertyarchives.info"
            ServerPort = 5750
        Case 3 'God Sandbox
            ServerDescription = "God Sandbox"
            CacheDirectory = App.Path + "\sandbox"
            ServerIP = "libertyarchives.info"
            ServerPort = 5752
        Case 4 'Condemned
            ServerDescription = "Condemned"
            CacheDirectory = App.Path + "\condemned"
            ServerIP = "libertyarchives.info"
            ServerPort = 5753
        Case 5 'Fankenstein
            ServerDescription = "Fankenstein"
            CacheDirectory = App.Path + "\fankenstein"
            ServerIP = "libertyarchives.info"
            ServerPort = 5751
    End Select
    
    On Error Resume Next
    sckPing(0).Close
    sckPing(1).Close
    sckPing(2).Close
    sckPing(3).Close
    sckPing(4).Close
    sckPing(5).Close
    On Error GoTo 0
    
    Unload Me
    InitializeGame
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        ServerDescription = "localhost"
        ServerIP = "127.0.0.1"
        ServerPort = 5750
        CacheDirectory = App.Path + "\localhost"
        Unload Me
        InitializeGame
    End If
End Sub

Private Sub Form_Load()

    'Override Code 'Skip server list for now
    'ServerDescription = "Classic"
    'CacheDirectory = "classic"
    'ServerIP = "208.79.77.39"
    'ServerPort = 5756
    'Unload Me
    'InitializeGame
    'End Override Code
    
    lstServers.Clear
    
    lstServers.AddItem "Classic"
    lstServers.AddItem "PK Island"
    lstServers.AddItem "Ethia"
    lstServers.AddItem "God Sandbox"
    lstServers.AddItem "Condemned"
    lstServers.AddItem "Fankenstein"
       
       
    'Classic
    sckPing(0).RemoteHost = "libertyarchives.info"
    sckPing(0).RemotePort = 5756
    sckPing(0).connect
       
    'PK Island
    sckPing(1).RemoteHost = "libertyarchives.info"
    sckPing(1).RemotePort = 5758
    sckPing(1).connect
    
    'Ethia
    sckPing(2).RemoteHost = "libertyarchives.info"
    sckPing(2).RemotePort = 5750
    sckPing(2).connect
    
    'God Sandbox
    sckPing(3).RemoteHost = "libertyarchives.info"
    sckPing(3).RemotePort = 5752
    sckPing(3).connect
    
    'Condemned
    sckPing(4).RemoteHost = "libertyarchives.info"
    sckPing(4).RemotePort = 5753
    sckPing(4).connect
    
    'Fankenstein
    sckPing(5).RemoteHost = "libertyarchives.info"
    sckPing(5).RemotePort = 5751
    sckPing(5).connect
    
    lstServers.ListIndex = 0
End Sub

Private Sub sckPing_Connect(index As Integer)
    Dim St As String, send As String
    St = Chr$(35)
    sckPing(index).SendData DoubleChar$(Len(St)) + Chr$(CheckSum(St) * 20 Mod 194) + Chr$(0) + St
End Sub

Private Sub sckPing_DataArrival(index As Integer, ByVal bytesTotal As Long)
    Dim Receive As String
    sckPing(index).GetData Receive, vbString, bytesTotal
    lstServers.List(index) = lstServers.List(index) + " (" + Receive + ")"
    sckPing(index).Close
End Sub

