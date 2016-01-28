VERSION 4.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AgDRIFT Information"
   ClientHeight    =   4980
   ClientLeft      =   2055
   ClientTop       =   1605
   ClientWidth     =   5640
   ForeColor       =   &H80000008&
   Height          =   5385
   Icon            =   "INFO.frx":0000
   Left            =   1995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5640
   Top             =   1260
   Width           =   5760
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1560
      ScaleHeight     =   735
      ScaleWidth      =   2655
      TabIndex        =   4
      Top             =   360
      Width           =   2655
      Begin VB.Label lblTM 
         AutoSize        =   -1  'True
         Caption         =   "TM"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   -1  'True
            strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.Line linLogo 
         BorderColor     =   &H000000FF&
         BorderWidth     =   4
         X1              =   720
         X2              =   2400
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblLogo 
         Caption         =   "AgDRIFT"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   24
            underline       =   0   'False
            italic          =   -1  'True
            strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1441
      Left            =   4680
      TabIndex        =   0
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   $"INFO.frx":030A
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Spray Drift Task Force Spray Model"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   5445
   End
   Begin VB.Label Label2 
      Caption         =   $"INFO.frx":041D
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "lblVersion"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   5415
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' $Id: info.frm,v 1.3 2000/03/06 21:29:22 tom Exp $

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  'Center the form on the screen
  CenterForm Me
  'display the version string
  lblVersion.Caption = "Version " & Format$(AGDRIFTVERSION)
End Sub
