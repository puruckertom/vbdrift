VERSION 5.00
Begin VB.Form frmTier1air 
   BorderStyle     =   0  'None
   Caption         =   "Tier I Aerial Agricultural Input"
   ClientHeight    =   6795
   ClientLeft      =   1185
   ClientTop       =   1740
   ClientWidth     =   9480
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   HelpContextID   =   1280
   Icon            =   "TIER1AIR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Tag             =   "tier1"
   Begin VB.Frame fraDropSize 
      Caption         =   "Drop Size Distribution"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   4095
      Begin VB.OptionButton optDropSize 
         Caption         =   "optDropSize"
         Height          =   255
         HelpContextID   =   1280
         Index           =   8
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   3615
      End
      Begin VB.OptionButton optDropSize 
         Caption         =   "optDropSize"
         Height          =   255
         HelpContextID   =   1280
         Index           =   6
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   3615
      End
      Begin VB.OptionButton optDropSize 
         Caption         =   "optDropSize"
         Height          =   255
         HelpContextID   =   1280
         Index           =   4
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   3615
      End
      Begin VB.OptionButton optDropSize 
         Caption         =   "optDropSize"
         Height          =   255
         HelpContextID   =   1280
         Index           =   2
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.PictureBox picLogo 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5280
      ScaleHeight     =   735
      ScaleWidth      =   4095
      TabIndex        =   7
      Top             =   6000
      Width           =   4095
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
         TabIndex        =   9
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
      Begin VB.Label lblTier 
         Caption         =   "Tier I Aerial Agricultural"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   10
         Top             =   120
         Width           =   1485
      End
   End
   Begin VB.Frame fraRunTitle 
      Caption         =   "Title"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   9255
      Begin VB.TextBox txtRunTitle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   1300
         Left            =   120
         TabIndex        =   0
         Text            =   "Untitled"
         Top             =   240
         Width           =   9015
      End
   End
End
Attribute VB_Name = "frmTier1air"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tier1air.frm,v 1.6 2002/02/06 16:06:28 tom Exp $

Private Sub Form_Load()
'Initialize the data entry form
  Dim SaveDC As Integer
  Dim c As Control
  
  Me.Caption = FormCaption
  
  For Each c In optDropSize()
    c.Caption = GetBasicNameDSD(c.Index)
  Next
  optDropSize(4).Caption = optDropSize(4).Caption + " (default)"
  
  'Save the current state of DataChanged.
  'We need to do this because by loading a new form
  'and updating its controls, DataChanged will be
  'set.
  SaveDC = UI.DataChanged
  
  'Transfer User data to form controls
  txtRunTitle.Text = UD.Title               'Title
  'set the DSD option button the safe way, because not all BasicTypes
  'are supported. This should be taken care of in the tier change
  'routines, but... :-)
  For Each c In optDropSize()
    If c.Index = UD.DSD(0).BasicType Then
      c.Value = True
      Exit For
    End If
  Next
  
  UpdateDataChangedFlag SaveDC 'restore DataChanged
End Sub

Private Sub Form_Resize()
'relocate controls when the form is resized
  'position agdrift logo
  'the top must not go above the DSD frame
  toplimit = fraDropSize.Top + fraDropSize.Height + 300
  leftlimit = 300
  logotop = Me.ScaleHeight - picLogo.Height - 300
  logoleft = Me.ScaleWidth - picLogo.Width - 300
  If logotop < toplimit Then logotop = toplimit
  If logoleft < leftlimit Then logoleft = leftlimit
  picLogo.Top = logotop
  picLogo.Left = logoleft
  
  'position the title frame and text box
  'it must not get narrower than the DSD frame
  widlimit = fraDropSize.Left + fraDropSize.Width
  titlewidth = Me.ScaleWidth - fraRunTitle.Left - 100
  If titlewidth < widlimit Then titlewidth = widlimit
  fraRunTitle.Width = titlewidth
  'text box
  txtRunTitle.Width = fraRunTitle.Width - txtRunTitle.Left - 120
End Sub

Private Sub optDropSize_Click(Index As Integer)
  UD.DSD(0).BasicType = Index
  LoadTier1Data UD, UC
  UpdateDataChangedFlag True 'Data was changed
End Sub

Private Sub txtRunTitle_Change()
  UD.Title = txtRunTitle.Text
  UpdateDataChangedFlag True 'Data was changed
End Sub

