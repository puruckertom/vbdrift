VERSION 5.00
Begin VB.Form frmTier1orc 
   BorderStyle     =   0  'None
   Caption         =   "Tier I Orchard/Airblast Input"
   ClientHeight    =   6795
   ClientLeft      =   2265
   ClientTop       =   1830
   ClientWidth     =   9480
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   HelpContextID   =   1458
   Icon            =   "TIER1ORC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Tag             =   "tier1"
   Begin VB.PictureBox picLogo 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   4440
      ScaleHeight     =   735
      ScaleWidth      =   4935
      TabIndex        =   22
      Top             =   6000
      Width           =   4935
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
         TabIndex        =   30
         Top             =   0
         Width           =   195
      End
      Begin VB.Label lblTier 
         AutoSize        =   -1  'True
         Caption         =   "Tier I Orchard/Airblast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2520
         TabIndex        =   23
         Top             =   240
         Width           =   2415
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
         TabIndex        =   20
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame fraRunTitle 
      Caption         =   "Title"
      Height          =   735
      Left            =   120
      TabIndex        =   21
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
   Begin VB.Frame fraOrc 
      Caption         =   "Combination Orchards"
      Height          =   1575
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   3615
      Begin VB.OptionButton optBoomHeight 
         Caption         =   "optBoomHeight"
         Height          =   255
         HelpContextID   =   1458
         Index           =   15
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   3375
      End
      Begin VB.OptionButton optBoomHeight 
         Caption         =   "optBoomHeight"
         Height          =   255
         HelpContextID   =   1458
         Index           =   14
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   3375
      End
      Begin VB.OptionButton optBoomHeight 
         Caption         =   "optBoomHeight"
         Height          =   255
         HelpContextID   =   1458
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton optBoomHeight 
         Caption         =   "optBoomHeight"
         Height          =   255
         HelpContextID   =   1458
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin VB.OptionButton optBoomHeight 
         Caption         =   "optBoomHeight"
         Height          =   255
         HelpContextID   =   1458
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
   End
   Begin VB.Frame fraExtended 
      Caption         =   "Extended Settings"
      Height          =   3495
      Left            =   3840
      TabIndex        =   25
      Top             =   720
      Width           =   5535
      Begin VB.Frame fraSwathRange 
         Caption         =   "Swath Range"
         Height          =   1215
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   2895
         Begin VB.TextBox txtStartSwath 
            Height          =   285
            HelpContextID   =   1458
            Left            =   1560
            TabIndex        =   7
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtEndSwath 
            Height          =   285
            HelpContextID   =   1458
            Left            =   1560
            TabIndex        =   8
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblStartSwath 
            AutoSize        =   -1  'True
            Caption         =   "Starting Tree Row:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblEndSwath 
            AutoSize        =   -1  'True
            Caption         =   "Ending Tree Row:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1290
         End
      End
      Begin VB.CheckBox chkExtended 
         Caption         =   "&Access Extended Settings"
         Height          =   255
         HelpContextID   =   1458
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.Frame fraIndividual 
         Caption         =   "Individual Orchards"
         Height          =   1815
         Left            =   120
         TabIndex        =   29
         Top             =   1560
         Width           =   5295
         Begin VB.OptionButton optBoomHeight 
            Caption         =   "optBoomHeight"
            Height          =   255
            HelpContextID   =   1458
            Index           =   3
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   2415
         End
         Begin VB.OptionButton optBoomHeight 
            Caption         =   "optBoomHeight"
            Height          =   255
            HelpContextID   =   1458
            Index           =   13
            Left            =   2640
            TabIndex        =   19
            Top             =   1200
            Width           =   2415
         End
         Begin VB.OptionButton optBoomHeight 
            Caption         =   "optBoomHeight"
            Height          =   255
            HelpContextID   =   1458
            Index           =   12
            Left            =   2640
            TabIndex        =   18
            Top             =   960
            Width           =   2415
         End
         Begin VB.OptionButton optBoomHeight 
            Caption         =   "optBoomHeight"
            Height          =   255
            HelpContextID   =   1458
            Index           =   11
            Left            =   2640
            TabIndex        =   17
            Top             =   720
            Width           =   2415
         End
         Begin VB.OptionButton optBoomHeight 
            Caption         =   "optBoomHeight"
            Height          =   255
            HelpContextID   =   1458
            Index           =   10
            Left            =   2640
            TabIndex        =   16
            Top             =   480
            Width           =   2415
         End
         Begin VB.OptionButton optBoomHeight 
            Caption         =   "optBoomHeight"
            Height          =   255
            HelpContextID   =   1458
            Index           =   9
            Left            =   2640
            TabIndex        =   15
            Top             =   240
            Width           =   2415
         End
         Begin VB.OptionButton optBoomHeight 
            Caption         =   "optBoomHeight"
            Height          =   255
            HelpContextID   =   1458
            Index           =   8
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   2415
         End
         Begin VB.OptionButton optBoomHeight 
            Caption         =   "optBoomHeight"
            Height          =   255
            HelpContextID   =   1458
            Index           =   7
            Left            =   240
            TabIndex        =   13
            Top             =   1200
            Width           =   2415
         End
         Begin VB.OptionButton optBoomHeight 
            Caption         =   "optBoomHeight"
            Height          =   255
            HelpContextID   =   1458
            Index           =   6
            Left            =   240
            TabIndex        =   12
            Top             =   960
            Width           =   2415
         End
         Begin VB.OptionButton optBoomHeight 
            Caption         =   "optBoomHeight"
            Height          =   255
            HelpContextID   =   1458
            Index           =   5
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   2415
         End
         Begin VB.OptionButton optBoomHeight 
            Caption         =   "optBoomHeight"
            Height          =   255
            HelpContextID   =   1458
            Index           =   4
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   2415
         End
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Information"
      Height          =   1935
      Left            =   120
      TabIndex        =   31
      Top             =   2280
      Width           =   3615
      Begin VB.Label lblInfo 
         Caption         =   "lblInfo"
         Height          =   1575
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmTier1orc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tier1orc.frm,v 1.10 2002/02/06 16:06:29 tom Exp $
Private PropTakeAction As Boolean

Private Sub chkExtended_Click()
  Dim xUD As UserData
  
  If PropTakeAction Then
    PropTakeAction = False
    If chkExtended.Value = 1 Then 'checked
      fraSwathRange.Enabled = True
      lblStartSwath.Enabled = True
      txtStartSwath.Enabled = True
      txtStartSwath.Text = AGFormat$(UD.OA.BegTrow)
      lblEndSwath.Enabled = True
      txtEndSwath.Enabled = True
      txtEndSwath.Text = AGFormat$(UD.OA.EndTrow)
      'Individual orchards are hidden from World users
      If UI.HasConfidentialData Then
        fraIndividual.Enabled = True
        For i = 3 To 13
          optBoomHeight(i).Enabled = True
        Next
      End If
    Else                          'unchecked
      UserDataDefault xUD 'get a set of default data
      UD.OA.BegTrow = xUD.OA.BegTrow 'reset to defaults
      UD.OA.EndTrow = xUD.OA.EndTrow 'reset to defaults
      LoadTier1Data UD, UC
      fraSwathRange.Enabled = False
      lblStartSwath.Enabled = False
      txtStartSwath.Enabled = False
      txtStartSwath.Text = ""
      lblEndSwath.Enabled = False
      txtEndSwath.Enabled = False
      txtEndSwath.Text = ""
      'Individual orchards are hidden from World users
      If UI.HasConfidentialData Then
        fraIndividual.Enabled = False
        For i = 3 To 13
          optBoomHeight(i).Enabled = False
          'can't have an individual orchard
          If optBoomHeight(i).Value Then
            PropTakeAction = True
            optBoomHeight(0).Value = True
            PropTakeAction = False
          End If
        Next
      End If
    End If
    PropTakeAction = True
  End If
End Sub

Private Sub Form_Load()
'Initialize the data entry form
  Dim SaveDC As Integer
  Dim c As Control
  
  Me.Caption = FormCaption
  
  'Save the current state of DataChanged.
  'We need to do this because by loading a new form
  'and updating its controls, DataChanged will be
  'set.
  SaveDC = UI.DataChanged
  PropTakeAction = False
  
  'Label the option buttons
  For Each c In optBoomHeight()
    c.Caption = "&" + GetBasicNameOA(c.Index)
  Next
  
  'Hide extra combinations orchards from the public
  If Not AGDRIFTREGULATORY Then
    optBoomHeight(14).Visible = False
    optBoomHeight(15).Visible = False
  End If
  
  'Hide individual orchards from World users
  If Not UI.HasConfidentialData Then
    fraIndividual.Visible = False
  End If
  
  'Transfer User data to form controls
  txtRunTitle.Text = UD.Title               'Title
  optBoomHeight(UD.OA.BasicType) = True
  PropTakeAction = True
  
  lblInfo.Caption = GetTier1Info(UD.ApplMethod, UD.OA.BasicType)
  chkExtended_Click    'Updates Extended controls
  
  UpdateDataChangedFlag SaveDC 'restore DataChanged
End Sub

Private Sub Form_Resize()
'relocate controls when the form is resized
  'position agdrift logo
  'the top must not go above the Extended frame
  Const MRGN = 300
  toplimit = fraExtended.Top + fraExtended.Height + MRGN
  leftlimit = MRGN
  logotop = Me.ScaleHeight - picLogo.Height - MRGN
  logoleft = Me.ScaleWidth - picLogo.Width - MRGN
  If logotop < toplimit Then logotop = toplimit
  If logoleft < leftlimit Then logoleft = leftlimit
  picLogo.Top = logotop
  picLogo.Left = logoleft
  
  'position the title frame and text box
  'it must not get narrower than the Orchard frame
  widlimit = fraOrc.Left + fraOrc.Width
  titlewidth = Me.ScaleWidth - fraRunTitle.Left - 100
  If titlewidth < widlimit Then titlewidth = widlimit
  fraRunTitle.Width = titlewidth
  'text box
  txtRunTitle.Width = fraRunTitle.Width - txtRunTitle.Left - 120
End Sub

Private Sub optBoomHeight_Click(Index As Integer)
  Dim c As Control
  If PropTakeAction Then
    PropTakeAction = False
    'turn off buttons in other frames
    For Each c In optBoomHeight
      If c.Index <> Index Then c.Value = False
    Next
    PropTakeAction = True
    UD.OA.BasicType = Index
    LoadTier1Data UD, UC
    lblInfo.Caption = GetTier1Info(UD.ApplMethod, UD.OA.BasicType)
    UpdateDataChangedFlag True 'Data was changed
  End If
End Sub

Private Sub txtRunTitle_Change()
  If PropTakeAction Then
    UD.Title = txtRunTitle.Text
    UpdateDataChangedFlag True 'Data was changed
  End If
End Sub

Public Sub txtStartSwath_LostFocus()
'note that this routine is public so that
'T1GndOrchKluge can call it
  If PropTakeAction Then
    UD.OA.BegTrow = Val(txtStartSwath.Text)
    'Clamp the new value
    If UD.OA.BegTrow < 1 Then UD.OA.BegTrow = 1
    If UD.OA.BegTrow > 20 Then UD.OA.BegTrow = 20
    If UD.OA.BegTrow > UD.OA.EndTrow Then UD.OA.BegTrow = UD.OA.EndTrow
    'Update the text in case the value changed
    PropTakeAction = False
    txtStartSwath.Text = Format$(UD.OA.BegTrow)
    PropTakeAction = True
    'reload the data
    LoadTier1Data UD, UC
    UpdateDataChangedFlag True 'Data was changed
  End If
End Sub

Private Sub txtStartSwath_KeyPress(KeyAscii As Integer)
  If PropTakeAction Then
    If KeyAscii = Asc(vbCr) Then
      txtStartSwath_LostFocus
      KeyAscii = 0
    End If
  End If
End Sub

Public Sub txtEndSwath_LostFocus()
'note that this routine is public so that
'T1GndOrchKluge can call it
  If PropTakeAction Then
    UD.OA.EndTrow = Val(txtEndSwath.Text)
    'Clamp the new value
    If UD.OA.EndTrow < 1 Then UD.OA.EndTrow = 1
    If UD.OA.EndTrow > 20 Then UD.OA.EndTrow = 20
    If UD.OA.EndTrow < UD.OA.BegTrow Then UD.OA.EndTrow = UD.OA.BegTrow
    'Update the text in case the value changed
    PropTakeAction = False
    txtEndSwath.Text = Format$(UD.OA.EndTrow)
    PropTakeAction = True
    'reload the data
    LoadTier1Data UD, UC
    UpdateDataChangedFlag True 'Data was changed
  End If
End Sub

Private Sub txtEndSwath_KeyPress(KeyAscii As Integer)
  If PropTakeAction Then
    If KeyAscii = Asc(vbCr) Then
      txtEndSwath_LostFocus
      KeyAscii = 0
    End If
  End If
End Sub

