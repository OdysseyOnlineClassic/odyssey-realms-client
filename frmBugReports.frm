VERSION 5.00
Begin VB.Form frmBugReports 
   BackColor       =   &H0061514B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Odyssey Online Classic [Bug Reports]"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstReports 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   1980
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Width           =   5175
   End
   Begin VB.OptionButton optStatus 
      BackColor       =   &H0061514B&
      Caption         =   "Closed"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   12
      Top             =   5760
      Width           =   1215
   End
   Begin VB.OptionButton optStatus 
      BackColor       =   &H0061514B&
      Caption         =   "Resolving"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   11
      Top             =   5760
      Width           =   1215
   End
   Begin VB.OptionButton optStatus 
      BackColor       =   &H0061514B&
      Caption         =   "Open"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtIP 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   3000
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox txtPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   3000
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   2055
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3360
      Width           =   5175
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Bug resolve status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   5520
      Width           =   5175
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   7
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Bug description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   5295
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Submitted by Player - User:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Submitted bug reports requiring attention:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Update Report"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   345
      Index           =   1
      Left            =   2880
      TabIndex        =   2
      Top             =   6240
      Width           =   2475
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009AADC2&
      Height          =   345
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   6240
      Width           =   2475
   End
   Begin VB.Menu lstOptions 
      Caption         =   "List Options"
      Visible         =   0   'False
      Begin VB.Menu lstRemove 
         Caption         =   "Remove Bug Report"
      End
   End
End
Attribute VB_Name = "frmBugReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button_Click(Index As Integer)
    Select Case Index
        Case 0 'Cancel
            Unload Me
        Case 1 'Update
            
    End Select
End Sub

Private Sub lstRemove_Click()
Dim A As Long
    If MsgBox("Are you sure you wish to delete this bug report?", vbYesNo + vbCritical) = vbYes Then
        With lstReports
            SendSocket Chr$(92) + DoubleChar(.ItemData(.ListIndex))
            .RemoveItem .ListIndex
            txtPlayer = vbNullString
            txtIP = vbNullString
            txtDescription = vbNullString
            For A = 0 To 2
                optStatus(A).Enabled = False
                optStatus(A).value = False
            Next A
        End With
    End If
End Sub

Private Sub lstReports_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim A As Long
    With lstReports
        If .ListCount > 0 Then
            A = .TopIndex + Int(Y / (.Height / 10))
            If A < .ListCount Then
                .ListIndex = A
            Else
                .ListIndex = -1
            End If
        End If
    End With
End Sub

Private Sub lstReports_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim A As Long, B As Long
Dim St As String
    With lstReports
        If .ListCount > 0 Then
            If .ListIndex >= 0 Then
                A = .ItemData(.ListIndex)
                With Bug(A)
                    txtPlayer = .PlayerName + " - (" + .PlayerUser + ")"
                    txtIP = .PlayerIP
                    txtDescription = .Description
                    If .status < 3 Then
                        For B = 0 To 2
                            optStatus(B).Enabled = True
                        Next B
                        optStatus(.status - 1).value = True
                    Else
                        optStatus(2).value = True
                    End If
                End With
                If Button = 2 Then
                    Me.PopupMenu lstOptions, , X + 300, Y + 500
                End If
            Else
                txtPlayer = vbNullString
                txtIP = vbNullString
                txtDescription = vbNullString
                For B = 0 To 2
                    optStatus(B).Enabled = False
                    optStatus(B).value = False
                Next B
            End If
        End If
    End With
End Sub
