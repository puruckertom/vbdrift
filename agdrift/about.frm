VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About AgDRIFT"
   ClientHeight    =   5400
   ClientLeft      =   2985
   ClientTop       =   1680
   ClientWidth     =   7245
   ForeColor       =   &H80000008&
   HelpContextID   =   1005
   Icon            =   "ABOUT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5400
   ScaleWidth      =   7245
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2160
      ScaleHeight     =   735
      ScaleWidth      =   2655
      TabIndex        =   4
      Top             =   120
      Width           =   2655
      Begin VB.Label lblTM 
         AutoSize        =   -1  'True
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   7
         Top             =   0
         Width           =   195
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
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdAboutOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1005
      Left            =   3360
      TabIndex        =   0
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   $"ABOUT.frx":030A
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   6975
   End
   Begin VB.Label Label5 
      Caption         =   $"ABOUT.frx":05B4
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   $"ABOUT.frx":0721
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   6975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Spray Drift Task Force Spray Software"
      Height          =   195
      Left            =   2070
      TabIndex        =   3
      Top             =   960
      Width           =   2745
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "lblVersion"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   5415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: about.frm,v 1.11 2011/12/27 17:47:00 tom Exp $

Private Sub cmdAboutOK_Click()
  Unload Me
End Sub

Private Sub Form_Click()
  Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  'Exit when any key is released
  '
  'We do it this way (rather than KeyDown) so that the user
  'may capture the about screen with Alt-PrtSc. The only
  'side effect of this approach is that if the program is
  'by pressing a key, e.g. F5 in the development environment,
  'this form appears only fleetingly.
  Unload Me
End Sub

Private Sub Form_Load()
  'Center the form on the screen
  CenterForm Me
  'display the version string
  lblVersion.Caption = "Version " & GetVersionString(AGDRIFTVERSION)
End Sub

Private Sub Label1_Click()
  Unload Me
End Sub

Private Sub Label2_Click()
  Unload Me
End Sub

Private Sub Label3_Click()
  Unload Me
End Sub

Private Sub Label4_Click()
  Unload Me
End Sub

Private Sub Label5_Click()
  Unload Me
End Sub

Private Sub lblLogo_Click()
  Unload Me
End Sub

Private Sub lblTM_DblClick()
  'display the Easter Egg Screen
  frmEasterEgg.Show vbModal
End Sub

Private Sub lblVersion_Click()
  Unload Me
End Sub

Private Sub picLogo_Click()
  Unload Me
End Sub
