VERSION 5.00
Begin VB.Form frmReportBug 
   BackColor       =   &H0061514B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Odyssey Realms [Submit Bug Report]"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBugTitle 
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
      MaxLength       =   30
      TabIndex        =   8
      Top             =   2040
      Width           =   5895
   End
   Begin VB.TextBox txtBugReport 
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
      MaxLength       =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2640
      Width           =   5895
   End
   Begin VB.Label Label5 
      BackColor       =   &H0061514B&
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
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   5895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0061514B&
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter a short descriptive bug title below: (30 Characters Max)"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   5895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0061514B&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmReportBug.frx":0000
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
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Button 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0044342E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Submit"
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
      Left            =   4200
      TabIndex        =   5
      Top             =   5880
      Width           =   1980
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
      Left            =   2160
      TabIndex        =   4
      Top             =   5880
      Width           =   1980
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0061514B&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmReportBug.frx":01DC
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   5040
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackColor       =   &H0061514B&
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Characters Required: 30"
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
      Left            =   240
      TabIndex        =   2
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label lblCharCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0061514B&
      BackStyle       =   0  'Transparent
      Caption         =   "0/3000"
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
      Left            =   4560
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
End
Attribute VB_Name = "frmReportBug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button_Click(Index As Integer)
    Select Case Index
        Case 0 'Cancel
            Unload Me
        Case 1 'Submit
            If Len(txtBugTitle) >= 3 And Len(txtBugReport) >= 30 Then
                If MsgBox("Are you sure you wish to submit this bug report?", vbQuestion + vbYesNo) = vbYes Then
                    SendSocket Chr$(23) + txtBugTitle + Chr$(0) + txtBugReport
                    MsgBox "Your bug report has been submitted succesfully!", vbInformation + vbOKOnly
                    Unload Me
                End If
            Else
                MsgBox "Error: The title must be at least 3 characters long and the description must be at least 30 characters long.", vbCritical + vbOKOnly
            End If
    End Select
End Sub

Private Sub txtBugReport_Change()
    lblCharCount = CStr(Len(txtBugReport)) + "/3000"
    If Len(txtBugTitle) >= 3 And Len(txtBugReport) >= 30 Then
        Button(1).Enabled = True
    Else
        Button(1).Enabled = False
    End If
End Sub

Private Sub txtBugTitle_Change()
    If Len(txtBugTitle) >= 3 And Len(txtBugReport) >= 30 Then
        Button(1).Enabled = True
    Else
        Button(1).Enabled = False
    End If
End Sub
