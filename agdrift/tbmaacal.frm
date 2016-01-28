VERSION 5.00
Begin VB.Form frmTBMAACalc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multiple Application Assessment Calculations"
   ClientHeight    =   4875
   ClientLeft      =   1575
   ClientTop       =   1860
   ClientWidth     =   6150
   ForeColor       =   &H80000008&
   Icon            =   "TBMAACAL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4875
   ScaleWidth      =   6150
   Begin VB.Frame fraMessage 
      Caption         =   "Messages"
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5895
      Begin VB.ListBox lstCalcStat 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         HelpContextID   =   1451
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   375
      HelpContextID   =   1451
      Left            =   2640
      TabIndex        =   0
      Top             =   4440
      Width           =   855
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   5895
      Begin VB.PictureBox picTherm 
         AutoRedraw      =   -1  'True
         DrawMode        =   14  'Copy Pen
         Height          =   255
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   4995
         TabIndex        =   4
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label lblStatusMessage 
         Alignment       =   2  'Center
         Caption         =   "Status message 2"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   5655
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblStatusMessage 
         Alignment       =   2  'Center
         Caption         =   "Status message 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmTBMAACalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tbmaacal.frm,v 1.4 2001/04/26 16:22:02 tom Exp $
'Calculations form

Private Sub cmdStop_Click()
  UI.OkToDoCalcs = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'halt any calculations
  UI.OkToDoCalcs = False
End Sub

Private Sub Timer1_Timer()
  lblStatusMessage(1).Caption = "Elapsed Time: " & Format$(CDbl(Now) - CDbl(StartDate), "hh:mm:ss")
End Sub



