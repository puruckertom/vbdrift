VERSION 5.00
Begin VB.Form frmTier1Gnd 
   BorderStyle     =   0  'None
   Caption         =   "Tier I Ground Input"
   ClientHeight    =   6795
   ClientLeft      =   1650
   ClientTop       =   3570
   ClientWidth     =   9480
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   HelpContextID   =   1457
   Icon            =   "TIER1GND.frx":0000
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
      Left            =   5400
      ScaleHeight     =   735
      ScaleWidth      =   3975
      TabIndex        =   11
      Top             =   6000
      Width           =   3975
      Begin VB.Label lblTier 
         AutoSize        =   -1  'True
         Caption         =   "Tier I Ground"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1470
      End
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
         TabIndex        =   17
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
         TabIndex        =   9
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame fraRunTitle 
      Caption         =   "Title"
      Height          =   735
      Left            =   120
      TabIndex        =   10
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
   Begin VB.Frame fraExtended 
      Caption         =   "Extended Settings"
      Height          =   1575
      Left            =   3840
      TabIndex        =   15
      Top             =   720
      Width           =   5535
      Begin VB.TextBox txtSwaths 
         Height          =   285
         HelpContextID   =   1457
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chkExtended 
         Caption         =   "&Access Extended Settings"
         Height          =   255
         HelpContextID   =   1457
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblSwaths 
         Caption         =   "Number of Swaths:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame fraBoomHeight 
      Caption         =   "Boom Height"
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   3615
      Begin VB.OptionButton optBoomHeight 
         Caption         =   "&Low Boom"
         Height          =   255
         HelpContextID   =   1457
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optBoomHeight 
         Caption         =   "&High Boom"
         Height          =   255
         HelpContextID   =   1457
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame fraDropSize 
      Caption         =   "Drop Size Distribution"
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   3615
      Begin VB.OptionButton optDropSize 
         Caption         =   "ASAE Fine to Medium/Coarse"
         Height          =   255
         HelpContextID   =   1457
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton optDropSize 
         Caption         =   "ASAE Very Fine to Fine"
         Height          =   255
         HelpContextID   =   1457
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Information"
      Height          =   1575
      Left            =   3840
      TabIndex        =   18
      Top             =   2280
      Width           =   5535
      Begin VB.Label lblInfo 
         Caption         =   "lblInfo"
         Height          =   1215
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   5295
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraPercentile 
      Caption         =   "Data Percentile"
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   3615
      Begin VB.OptionButton optPercentile 
         Caption         =   "50th Percentile"
         Height          =   255
         HelpContextID   =   1457
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   2655
      End
      Begin VB.OptionButton optPercentile 
         Caption         =   "90th Percentile"
         Height          =   255
         HelpContextID   =   1457
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmTier1Gnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tier1gnd.frm,v 1.10 2002/02/06 16:06:28 tom Exp $
Private PropTakeAction As Boolean

Private Sub chkExtended_Click()
  Dim xUD As UserData
  If PropTakeAction Then
    PropTakeAction = False
    If chkExtended.Value = 1 Then 'checked
      lblSwaths.Enabled = True
      txtSwaths.Enabled = True
      txtSwaths.Text = AGFormat$(UD.GA.NumSwaths)
    Else                          'unchecked
      UserDataDefault xUD
      UD.GA.NumSwaths = xUD.GA.NumSwaths 'reset to default
      LoadTier1Data UD, UC
      lblSwaths.Enabled = False
      txtSwaths.Enabled = False
      txtSwaths.Text = ""
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
  PropTakeAction = True
  
  'Adjust form for Public/Regulatory versions
  If Not AGDRIFTREGULATORY Then
    fraPercentile.Visible = False
    For Each c In optPercentile
      c.Visible = False
    Next
  End If
  
  'Transfer User data to form controls
  txtRunTitle.Text = UD.Title               'Title
  optBoomHeight(0).Value = True
  optDropSize(0).Value = True
  If Not AGDRIFTREGULATORY Then
    optPercentile(0).Value = True  'Public uses this one only
  Else
    optPercentile(1).Value = True  'Regulatory default
  End If
  
  chkExtended_Click    'Updates Extended controls
  
  UpdateDataChangedFlag SaveDC 'restore DataChanged
End Sub

Private Sub Form_Resize()
'relocate controls when the form is resized
  'position agdrift logo
  'the top must not go above the extended frame
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
  widlimit = fraDropSize.Left + fraDropSize.Width
  titlewidth = Me.ScaleWidth - fraRunTitle.Left - 100
  If titlewidth < widlimit Then titlewidth = widlimit
  fraRunTitle.Width = titlewidth
  'text box
  txtRunTitle.Width = fraRunTitle.Width - txtRunTitle.Left - 120
End Sub

Private Sub optBoomHeight_Click(Index As Integer)
  Dim iBOOM As Integer
  Dim iDSD As Integer
  Dim iPCT As Integer
  Dim c As Control
  If PropTakeAction Then
    iBOOM = Index
    For Each c In optDropSize
      If c.Value Then iDSD = c.Index
    Next
    For Each c In optPercentile
      If c.Value Then iPCT = c.Index
    Next
    UD.GA.BasicType = (4 * iPCT) + (2 * iBOOM) + iDSD
    LoadTier1Data UD, UC
    lblInfo.Caption = GetTier1Info(UD.ApplMethod, UD.GA.BasicType)
    UpdateDataChangedFlag True 'Data was changed
  End If
End Sub

Private Sub optDropSize_Click(Index As Integer)
  Dim iBOOM As Integer
  Dim iDSD As Integer
  Dim iPCT As Integer
  Dim c As Control
  If PropTakeAction Then
    iDSD = Index
    For Each c In optBoomHeight
      If c.Value Then iBOOM = c.Index
    Next
    For Each c In optPercentile
      If c.Value Then iPCT = c.Index
    Next
    UD.GA.BasicType = (4 * iPCT) + (2 * iBOOM) + iDSD
    LoadTier1Data UD, UC
    lblInfo.Caption = GetTier1Info(UD.ApplMethod, UD.GA.BasicType)
    UpdateDataChangedFlag True 'Data was changed
  End If
End Sub

Private Sub optPercentile_Click(Index As Integer)
  Dim iBOOM As Integer
  Dim iDSD As Integer
  Dim iPCT As Integer
  Dim c As Control
  If PropTakeAction Then
    iPCT = Index
    For Each c In optBoomHeight
      If c.Value Then iBOOM = c.Index
    Next
    For Each c In optDropSize
      If c.Value Then iDSD = c.Index
    Next
    UD.GA.BasicType = (4 * iPCT) + (2 * iBOOM) + iDSD
    LoadTier1Data UD, UC
    lblInfo.Caption = GetTier1Info(UD.ApplMethod, UD.GA.BasicType)
    UpdateDataChangedFlag True 'Data was changed
  End If
End Sub

Private Sub txtRunTitle_Change()
  If PropTakeAction Then
    UD.Title = txtRunTitle.Text
    UpdateDataChangedFlag True 'Data was changed
  End If
End Sub

Public Sub txtSwaths_LostFocus()
'note how this sub is Public so that T1GndOrchKluge can call it
  If PropTakeAction Then
    UD.GA.NumSwaths = Val(txtSwaths.Text)
    'Clamp the new value
    If UD.GA.NumSwaths < 1 Then UD.GA.NumSwaths = 1
    If UD.GA.NumSwaths > 20 Then UD.GA.NumSwaths = 20
    'Update the text in case the value changed
    PropTakeAction = False
    txtSwaths.Text = Format$(UD.GA.NumSwaths)
    PropTakeAction = True
    'reload the data
    LoadTier1Data UD, UC
    UpdateDataChangedFlag True 'Data was changed
  End If
End Sub

Private Sub txtSwaths_KeyPress(KeyAscii As Integer)
  If PropTakeAction Then
    If KeyAscii = Asc(vbCr) Then
      txtSwaths_LostFocus
      KeyAscii = 0
    End If
  End If
End Sub
