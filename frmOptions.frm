VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H0044342E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Odyssey Realms [Options]"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4935
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H0061514B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox chkFrameRate 
         BackColor       =   &H0061514B&
         Caption         =   "High"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   1680
         TabIndex        =   19
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox chkFrameRate 
         BackColor       =   &H0061514B&
         Caption         =   "Medium"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   1680
         TabIndex        =   18
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox chkFrameRate 
         BackColor       =   &H0061514B&
         Caption         =   "Low"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   1680
         TabIndex        =   17
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CheckBox chkDisableLighting 
         BackColor       =   &H0061514B&
         Caption         =   "Disable Lighting and Weather"
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
         Height          =   420
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chkDisablePlayerLights 
         BackColor       =   &H0061514B&
         Caption         =   "Disable Other Player's Lights"
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
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CheckBox chkHighPriority 
         BackColor       =   &H0061514B&
         Caption         =   "High Priority"
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
         Height          =   300
         Left            =   2640
         TabIndex        =   12
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optLighting 
         BackColor       =   &H0061514B&
         Caption         =   "Low"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   1455
      End
      Begin VB.OptionButton optLighting 
         BackColor       =   &H0061514B&
         Caption         =   "Medium"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   1455
      End
      Begin VB.OptionButton optLighting 
         BackColor       =   &H0061514B&
         Caption         =   "High"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox chkBroadcasts 
         BackColor       =   &H0061514B&
         Caption         =   "Display Broadcasts"
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
         Height          =   300
         Left            =   2640
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
      Begin VB.CheckBox chkWAV 
         BackColor       =   &H0061514B&
         Caption         =   "WAV Sound Effects"
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
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox chkMidi 
         BackColor       =   &H0061514B&
         Caption         =   "MIDI Music"
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
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1935
      End
      Begin VB.CheckBox chkWindowed 
         BackColor       =   &H0061514B&
         Caption         =   "Windowed Mode"
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
         Height          =   300
         Left            =   2640
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label cmdReportBug 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Report Bug"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   3000
         TabIndex        =   20
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackColor       =   &H0061514B&
         Caption         =   "Frame Rate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label cmdMacros 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Macros"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   1680
         Width           =   1635
      End
      Begin VB.Label cmdChangePassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   1320
         Width           =   1620
      End
      Begin VB.Label lblLightingQuality 
         BackColor       =   &H0061514B&
         Caption         =   "Lighting Quality:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label btnCancel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   350
         Left            =   1800
         TabIndex        =   6
         Top             =   3960
         Width           =   1275
      End
      Begin VB.Label btnOk 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0044342E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009AADC2&
         Height          =   350
         Left            =   3240
         TabIndex        =   5
         Top             =   3960
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Unload Me
    If blnPlaying = False Then frmMenu.Show
End Sub


Private Sub btnOk_Click()
Dim A As Long
    With Options
        If chkMidi = 1 Then
            .MIDI = True
        Else
            If .MIDI = True Then
                StopMidi
            End If
            .MIDI = False
        End If
        If chkWAV = 1 Then
            .Wav = True
        Else
            .Wav = False
        End If
        If chkBroadcasts = 1 Then
            .Broadcasts = True
        Else
            .Broadcasts = False
        End If
        If chkWindowed = 1 Then
            .Windowed = True
        Else
            .Windowed = False
        End If
        If chkHighPriority = 1 Then
            .HighPriority = True
            SetPriority HIGH_PRIORITY_CLASS
        Else
            .HighPriority = False
            SetPriority NORMAL_PRIORITY_CLASS
        End If
        If optLighting(0) = True Then
            .LightingQuality = 0
        ElseIf optLighting(1) = True Then
            .LightingQuality = 1
        ElseIf optLighting(2) = True Then
            .LightingQuality = 2
        End If
        For A = 0 To 2
            If chkFrameRate(A) = 1 Then
                .FrameRate = A
                Exit For
            End If
        Next A
        If chkDisablePlayerLights = 1 Then
            .DisablePlayerLights = True
        Else
            .DisablePlayerLights = False
        End If
        If chkDisableLighting = 1 Then
            .DisableLighting = True
        Else
            .DisableLighting = False
        End If
    End With
    SaveOptions
    If blnPlaying = True Then
        RedrawMap = True
    Else
        frmMenu.Show
    End If
    Unload Me
End Sub

Private Sub chkFrameRate_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim A As Long
    For A = 0 To 2
        If A <> Index Then chkFrameRate(A) = 0
    Next A
End Sub

Private Sub cmdChangePassword_Click()
    frmNewPass.Show
    Unload Me
End Sub

Private Sub cmdMacros_Click()
    frmMacros.Show
End Sub

Private Sub cmdReportBug_Click()
    frmReportBug.Show
End Sub

Private Sub Form_Load()
Dim A As Long
    With Options
        If .MIDI = True Then
            chkMidi = 1
        Else
            chkMidi = 0
        End If
        If .Wav = True Then
            chkWAV = 1
        Else
            chkWAV = 0
        End If
        If .Broadcasts = True Then
            chkBroadcasts = 1
        Else
            chkBroadcasts = 0
        End If
        If .Windowed = True Then
            chkWindowed = 1
        Else
            chkWindowed = 0
        End If
        If .HighPriority = True Then
            chkHighPriority = 1
        Else
            chkHighPriority = 0
        End If
        If .LightingQuality = 0 Then
            optLighting(0) = True
        ElseIf .LightingQuality = 1 Then
            optLighting(1) = True
        ElseIf .LightingQuality = 2 Then
            optLighting(2) = True
        End If
        For A = 0 To 2
            If A = .FrameRate Then
                chkFrameRate(A) = 1
            Else
                chkFrameRate(A) = 0
            End If
        Next A
        If .DisablePlayerLights = True Then
            chkDisablePlayerLights = 1
        Else
            chkDisablePlayerLights = 0
        End If
        If .DisableLighting = True Then
            chkDisableLighting = 1
        Else
            chkDisableLighting = 0
        End If
    End With
    frmOptions_Loaded = True
    If blnPlaying = False Then
        cmdChangePassword.Visible = False
        cmdReportBug.Visible = False
    Else
        cmdChangePassword.Visible = True
        cmdReportBug.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmOptions_Loaded = False
End Sub

