VERSION 5.00
Begin VB.Form frmTier1Lib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tier 1 Library"
   ClientHeight    =   4965
   ClientLeft      =   1770
   ClientTop       =   2040
   ClientWidth     =   9450
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   HelpContextID   =   1285
   Icon            =   "TIER1LIB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4965
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Tag             =   "tier1"
   Begin VB.Frame fraApplMeth 
      Caption         =   "Application Method"
      Height          =   615
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   9375
      Begin VB.OptionButton optApplMeth 
         Caption         =   "Orchard/Airblast"
         Height          =   255
         HelpContextID   =   1285
         Index           =   2
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optApplMeth 
         Caption         =   "Ground"
         Height          =   255
         HelpContextID   =   1285
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optApplMeth 
         Caption         =   "Aerial"
         Height          =   255
         HelpContextID   =   1285
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraOrchard 
      Caption         =   "Orchard Airblast"
      Height          =   3855
      Left            =   1440
      TabIndex        =   42
      Top             =   600
      Width           =   9375
      Begin VB.Frame fraOrc 
         Caption         =   "Combination Orchards"
         Height          =   1575
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   3735
         Begin VB.OptionButton optOrchard 
            Caption         =   "optOrchard"
            Height          =   255
            HelpContextID   =   1458
            Index           =   14
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   3495
         End
         Begin VB.OptionButton optOrchard 
            Caption         =   "optOrchard"
            Height          =   255
            HelpContextID   =   1458
            Index           =   15
            Left            =   120
            TabIndex        =   27
            Top             =   1200
            Width           =   3495
         End
         Begin VB.OptionButton optOrchard 
            Caption         =   "optOrchard"
            Height          =   255
            HelpContextID   =   1458
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   3495
         End
         Begin VB.OptionButton optOrchard 
            Caption         =   "optOrchard"
            Height          =   255
            HelpContextID   =   1458
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   3495
         End
         Begin VB.OptionButton optOrchard 
            Caption         =   "optOrchard"
            Height          =   255
            HelpContextID   =   1458
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   3495
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Extended Settings"
         Height          =   3495
         Left            =   3960
         TabIndex        =   54
         Top             =   240
         Width           =   5295
         Begin VB.CheckBox chkExtendedOrch 
            Caption         =   "&Access Extended Settings"
            Height          =   255
            HelpContextID   =   1458
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   2175
         End
         Begin VB.Frame fraSwathRange 
            Caption         =   "Swath Range"
            Height          =   1335
            Left            =   2520
            TabIndex        =   55
            Top             =   240
            Width           =   2655
            Begin VB.TextBox txtEndSwath 
               Height          =   285
               HelpContextID   =   1458
               Left            =   1560
               TabIndex        =   30
               Top             =   720
               Width           =   735
            End
            Begin VB.TextBox txtStartSwath 
               Height          =   285
               HelpContextID   =   1458
               Left            =   1560
               TabIndex        =   29
               Top             =   360
               Width           =   735
            End
            Begin VB.Label lblEndSwath 
               AutoSize        =   -1  'True
               Caption         =   "Ending Tree Row:"
               Height          =   195
               Left            =   120
               TabIndex        =   57
               Top             =   720
               Width           =   1290
            End
            Begin VB.Label lblStartSwath 
               AutoSize        =   -1  'True
               Caption         =   "Starting Tree Row:"
               Height          =   195
               Left            =   120
               TabIndex        =   56
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.Frame fraIndividual 
            Caption         =   "Individual Orchards"
            Height          =   1815
            Left            =   120
            TabIndex        =   61
            Top             =   1560
            Width           =   5055
            Begin VB.OptionButton optOrchard 
               Caption         =   "optOrchard"
               Height          =   255
               HelpContextID   =   1458
               Index           =   4
               Left            =   120
               TabIndex        =   32
               Top             =   480
               Width           =   2415
            End
            Begin VB.OptionButton optOrchard 
               Caption         =   "optOrchard"
               Height          =   255
               HelpContextID   =   1458
               Index           =   5
               Left            =   120
               TabIndex        =   33
               Top             =   720
               Width           =   2415
            End
            Begin VB.OptionButton optOrchard 
               Caption         =   "optOrchard"
               Height          =   255
               HelpContextID   =   1458
               Index           =   6
               Left            =   120
               TabIndex        =   34
               Top             =   960
               Width           =   2415
            End
            Begin VB.OptionButton optOrchard 
               Caption         =   "optOrchard"
               Height          =   255
               HelpContextID   =   1458
               Index           =   7
               Left            =   120
               TabIndex        =   36
               Top             =   1200
               Width           =   2415
            End
            Begin VB.OptionButton optOrchard 
               Caption         =   "optOrchard"
               Height          =   255
               HelpContextID   =   1458
               Index           =   8
               Left            =   120
               TabIndex        =   62
               Top             =   1440
               Width           =   2415
            End
            Begin VB.OptionButton optOrchard 
               Caption         =   "optOrchard"
               Height          =   255
               HelpContextID   =   1458
               Index           =   9
               Left            =   2520
               TabIndex        =   37
               Top             =   240
               Width           =   2415
            End
            Begin VB.OptionButton optOrchard 
               Caption         =   "optOrchard"
               Height          =   255
               HelpContextID   =   1458
               Index           =   10
               Left            =   2520
               TabIndex        =   38
               Top             =   480
               Width           =   2415
            End
            Begin VB.OptionButton optOrchard 
               Caption         =   "optOrchard"
               Height          =   255
               HelpContextID   =   1458
               Index           =   11
               Left            =   2520
               TabIndex        =   39
               Top             =   720
               Width           =   2415
            End
            Begin VB.OptionButton optOrchard 
               Caption         =   "optOrchard"
               Height          =   255
               HelpContextID   =   1458
               Index           =   12
               Left            =   2520
               TabIndex        =   40
               Top             =   960
               Width           =   2415
            End
            Begin VB.OptionButton optOrchard 
               Caption         =   "optOrchard"
               Height          =   255
               HelpContextID   =   1458
               Index           =   13
               Left            =   2520
               TabIndex        =   41
               Top             =   1200
               Width           =   2415
            End
            Begin VB.OptionButton optOrchard 
               Caption         =   "optOrchard"
               Height          =   255
               HelpContextID   =   1458
               Index           =   3
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   2415
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Information"
         Height          =   1935
         Left            =   120
         TabIndex        =   52
         Top             =   1800
         Width           =   3735
         Begin VB.Label lblInfoOrch 
            Caption         =   "lblInfo"
            Height          =   1575
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   3495
         End
      End
   End
   Begin VB.Frame fraGround 
      Caption         =   "Ground Sprayer"
      Height          =   3735
      Left            =   960
      TabIndex        =   43
      Top             =   600
      Width           =   9375
      Begin VB.Frame fraExtended 
         Caption         =   "Extended Settings"
         Height          =   1215
         Left            =   3480
         TabIndex        =   47
         Top             =   240
         Width           =   5775
         Begin VB.CheckBox chkExtendedGnd 
            Caption         =   "&Access Extended Settings"
            Height          =   255
            HelpContextID   =   1457
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtSwathsGnd 
            Height          =   285
            HelpContextID   =   1457
            Left            =   1680
            TabIndex        =   22
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblSwathsGnd 
            Caption         =   "Number of Swaths:"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraBoomHeight 
         Caption         =   "Boom Height"
         Height          =   1095
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   3255
         Begin VB.OptionButton optBoomHeightGnd 
            Caption         =   "&High Boom"
            Height          =   255
            HelpContextID   =   1457
            Index           =   1
            Left            =   360
            TabIndex        =   16
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optBoomHeightGnd 
            Caption         =   "&Low Boom"
            Height          =   255
            HelpContextID   =   1457
            Index           =   0
            Left            =   360
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Drop Size Distribution"
         Height          =   1215
         Left            =   120
         TabIndex        =   49
         Top             =   1320
         Width           =   3255
         Begin VB.OptionButton optDropSizeGnd 
            Caption         =   "ASAE Very Fine to Fine"
            Height          =   255
            HelpContextID   =   1457
            Index           =   0
            Left            =   360
            TabIndex        =   17
            Top             =   360
            Width           =   2535
         End
         Begin VB.OptionButton optDropSizeGnd 
            Caption         =   "ASAE Fine to Medium/Coarse"
            Height          =   255
            HelpContextID   =   1457
            Index           =   1
            Left            =   360
            TabIndex        =   18
            Top             =   720
            Width           =   2535
         End
      End
      Begin VB.Frame fraInfo 
         Caption         =   "Information"
         Height          =   2175
         Left            =   3480
         TabIndex        =   50
         Top             =   1440
         Width           =   5775
         Begin VB.Label lblInfoGnd 
            Caption         =   "lblInfoGnd"
            Height          =   1815
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   5535
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraPercentile 
         Caption         =   "Data Percentile"
         Height          =   1095
         Left            =   120
         TabIndex        =   63
         Top             =   2520
         Width           =   3255
         Begin VB.OptionButton optPercentile 
            Caption         =   "90th Percentile"
            Height          =   255
            HelpContextID   =   1457
            Index           =   1
            Left            =   360
            TabIndex        =   20
            Top             =   720
            Width           =   2775
         End
         Begin VB.OptionButton optPercentile 
            Caption         =   "50th Percentile"
            Height          =   255
            HelpContextID   =   1457
            Index           =   0
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Width           =   2655
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1285
      Left            =   7560
      TabIndex        =   0
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1285
      Left            =   8520
      TabIndex        =   1
      Top             =   4560
      Width           =   855
   End
   Begin VB.Frame fraAerialFS 
      Caption         =   "Aerial"
      Height          =   2775
      Left            =   480
      TabIndex        =   44
      Top             =   600
      Width           =   4335
      Begin VB.Frame fraDropSizeFS 
         Caption         =   "Drop Size Distribution"
         Height          =   2415
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   4095
         Begin VB.OptionButton optDropSizeFS 
            Caption         =   "optDropSizeFS"
            Height          =   255
            HelpContextID   =   1280
            Index           =   0
            Left            =   360
            TabIndex        =   9
            Top             =   240
            Width           =   3615
         End
         Begin VB.OptionButton optDropSizeFS 
            Caption         =   "optDropSizeFS"
            Height          =   255
            HelpContextID   =   1280
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   600
            Width           =   3615
         End
         Begin VB.OptionButton optDropSizeFS 
            Caption         =   "optDropSizeFS"
            Height          =   255
            HelpContextID   =   1280
            Index           =   4
            Left            =   360
            TabIndex        =   11
            Top             =   960
            Width           =   3615
         End
         Begin VB.OptionButton optDropSizeFS 
            Caption         =   "optDropSizeFS"
            Height          =   255
            HelpContextID   =   1280
            Index           =   6
            Left            =   360
            TabIndex        =   12
            Top             =   1320
            Width           =   3615
         End
         Begin VB.OptionButton optDropSizeFS 
            Caption         =   "optDropSizeFS"
            Height          =   255
            HelpContextID   =   1280
            Index           =   8
            Left            =   360
            TabIndex        =   13
            Top             =   1680
            Width           =   3615
         End
         Begin VB.OptionButton optDropSizeFS 
            Caption         =   "optDropSizeFS"
            Height          =   255
            HelpContextID   =   1280
            Index           =   10
            Left            =   360
            TabIndex        =   14
            Top             =   2040
            Width           =   3615
         End
      End
   End
   Begin VB.Frame fraAerial 
      Caption         =   "Aerial"
      Height          =   2055
      Left            =   0
      TabIndex        =   35
      Top             =   600
      Width           =   4335
      Begin VB.Frame fraDropSize 
         Caption         =   "Drop Size Distribution"
         Height          =   1695
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   4095
         Begin VB.OptionButton optDropSize 
            Caption         =   "optDropSize"
            Height          =   255
            HelpContextID   =   1280
            Index           =   2
            Left            =   360
            TabIndex        =   5
            Top             =   240
            Width           =   3615
         End
         Begin VB.OptionButton optDropSize 
            Caption         =   "optDropSize"
            Height          =   255
            HelpContextID   =   1280
            Index           =   4
            Left            =   360
            TabIndex        =   6
            Top             =   600
            Width           =   3615
         End
         Begin VB.OptionButton optDropSize 
            Caption         =   "optDropSize"
            Height          =   255
            HelpContextID   =   1280
            Index           =   6
            Left            =   360
            TabIndex        =   7
            Top             =   960
            Width           =   3615
         End
         Begin VB.OptionButton optDropSize 
            Caption         =   "optDropSize"
            Height          =   255
            HelpContextID   =   1280
            Index           =   8
            Left            =   360
            TabIndex        =   8
            Top             =   1320
            Width           =   3615
         End
      End
   End
End
Attribute VB_Name = "frmTier1Lib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tier1lib.frm,v 1.11 2001/08/13 17:40:07 tom Exp $
'This form generates a "key" string that tells the plot
'data generation routines what Tier 1 data to retrieve.
Option Explicit

Public AerialOnly As Boolean 'Set to True to restrict selection

Public DataSource As String
Public DataTitle As String

Public ApplMethod As Integer  '0=Aerial 1=Ground/Orchard
Public AerialType As Integer
Public GroundType As Integer
Public GroundSwaths As Integer
Public OrchardType As Integer
Public OrchardStartSwath As Integer
Public OrchardEndSwath As Integer

Private PropTakeAction As Boolean

Private Sub chkExtendedGnd_Click()
  If PropTakeAction Then
    If chkExtendedGnd.Value = 1 Then 'checked
      lblSwathsGnd.Enabled = True
      txtSwathsGnd.Enabled = True
      txtSwathsGnd.Text = AGFormat$(GroundSwaths)
    Else
      lblSwathsGnd.Enabled = False
      txtSwathsGnd.Enabled = False
      txtSwathsGnd.Text = ""
      GroundSwaths = 20
    End If
  End If
End Sub

Private Sub chkExtendedOrch_Click()
  Dim i As Integer
  If PropTakeAction Then
    PropTakeAction = False 'turn off controls for most of this
    If chkExtendedOrch.Value = 1 Then 'checked
      fraSwathRange.Enabled = True
      lblStartSwath.Enabled = True
      txtStartSwath.Enabled = True
      txtStartSwath.Text = AGFormat$(OrchardStartSwath)
      lblEndSwath.Enabled = True
      txtEndSwath.Enabled = True
      txtEndSwath.Text = AGFormat$(OrchardEndSwath)
      'Individual orchards are hidden from World users
      If UI.HasConfidentialData Then
        fraIndividual.Enabled = True
        For i = 3 To 13
          optOrchard(i).Enabled = True
        Next
      End If
    Else
      OrchardStartSwath = 1
      OrchardEndSwath = 20
      fraSwathRange.Enabled = False
      lblStartSwath.Enabled = False
      txtStartSwath.Enabled = False
      txtStartSwath.Text = ""
      lblEndSwath.Enabled = False
      txtEndSwath.Enabled = False
      txtEndSwath.Text = ""
      'Individual orchards are hidden from World users
      PropTakeAction = True 'need controls on for this part
      If UI.HasConfidentialData Then
        fraIndividual.Enabled = False
        For i = 3 To 13
          optOrchard(i).Enabled = False
          'can't have an individual orchard
          If optOrchard(i).Value Then
            optOrchard(i).Value = False
            optOrchard(0).Value = True
          End If
        Next
      End If
    End If
    PropTakeAction = True 'turn controls back on
  End If
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOk_Click()
  Dim Src As String
  Dim Title As String
  
  Src = "Tier1Lib: " + Format$(ApplMethod)
  Title = "Tier I Lib: " + GetBasicNameAM(ApplMethod)
  Select Case ApplMethod
  Case AM_AERIAL
    Src = Src + ", " + Format$(AerialType)
    Title = Title + ", " + GetBasicNameDSD(AerialType)
  Case AM_GROUND
    Src = Src + ", " + Format$(GroundType)
    Src = Src + ", " + Format$(GroundSwaths)
    Title = Title + ", " + GetBasicNameGA(GroundType)
    Title = Title + ", " + Format$(GroundSwaths)
  Case AM_ORCHARD
    Src = Src + ", " + Format$(OrchardType)
    Src = Src + ", " + Format$(OrchardStartSwath)
    Src = Src + ", " + Format$(OrchardEndSwath)
    Title = Title + ", " + GetBasicNameOA(OrchardType)
    Title = Title + ", " + Format$(OrchardStartSwath)
    Title = Title + ", " + Format$(OrchardEndSwath)
  End Select
  'Save these strings in the form's output fields
  DataSource = Src
  DataTitle = Title
  Me.Hide
End Sub

Private Sub Form_Activate()
'Initialize the data entry form
'All this stuff goes in the Activate Event
'so that form variables can be initialized

'tbc change this to use public subs
  Dim c As Control
  
  CenterForm Me
  
  'Enable control responses
  PropTakeAction = True
  
  'Hide Aerial for Public use version.
  'Obviously, setting AerialOnly to true in the
  'Public use version would produce a form that
  'is not very useful.
  If Not AGDRIFTREGULATORY Then
    optApplMeth(0).Visible = False
    optApplMeth(1).Value = True 'ground
  Else
    optApplMeth(0).Value = True 'Aerial
  End If
  
  'Application Method
  'disable for AerialOnly flag
  optApplMeth(1).Enabled = Not AerialOnly
  optApplMeth(2).Enabled = Not AerialOnly
  
  'Aerial
  optDropSize(4).Value = True 'ASAE Fine to Medium
  optDropSizeFS(4).Value = True 'ASAE Fine to Medium
  
  'Ground
  If Not AGDRIFTREGULATORY Then
    fraPercentile.Visible = False
    For Each c In optPercentile
      c.Visible = False
    Next
  End If
  optBoomHeightGnd(0).Value = True
  optDropSizeGnd(0).Value = True
  If Not AGDRIFTREGULATORY Then
    optPercentile(0).Value = True  'Public uses this one only
  Else
    optPercentile(1).Value = True  'Regulatory default
  End If
  chkExtendedGnd_Click
  
  'Orchard/Airblast
  For Each c In optOrchard()
    c.Caption = "&" + GetBasicNameOA(c.Index)
  Next
  optOrchard(0).Value = True
  chkExtendedOrch_Click
  'Hide extra combinations orchards from the public
  If Not AGDRIFTREGULATORY Then
    optOrchard(14).Visible = False
    optOrchard(15).Visible = False
  End If
  'turn off individual orchards for World users
  If Not UI.HasConfidentialData Then
    fraIndividual.Visible = False
  End If
  
  Me.Tag = ""  'default return value
End Sub

Private Sub Form_Load()
  Dim c As Control
  
  'Aerial (SDTF)
  For Each c In optDropSize()
    c.Caption = GetBasicNameDSD(c.Index)
  Next
  optDropSize(4).Caption = optDropSize(4).Caption + " (default)"

  For Each c In optDropSizeFS()
    c.Caption = GetBasicNameDSD(c.Index)
  Next
  optDropSizeFS(4).Caption = optDropSizeFS(4).Caption + " (default)"

  'initialize return variables
  DataSource = ""
  DataTitle = ""
  
  'initialize some variables
  GroundSwaths = 20
  OrchardStartSwath = 1
  OrchardEndSwath = 20
End Sub

Private Sub optApplMeth_Click(Index As Integer)
  Dim f As Frame
 
  If PropTakeAction Then
    'Set form variable
    ApplMethod = Index
  
    'Start with a known display state
    fraAerial.Visible = False
    fraAerialFS.Visible = False
    fraGround.Visible = False
    fraOrchard.Visible = False
  
    'Select the frame to display
    Select Case ApplMethod
    Case AM_AERIAL
      Select Case UD.Smokey
      Case AUD_SDTF
        Set f = fraAerial
      Case AUD_FS
        Set f = fraAerialFS
      End Select
    Case AM_GROUND
        Set f = fraGround
    Case AM_ORCHARD
        Set f = fraOrchard
    End Select
  
    'Position and display it
    f.Left = fraApplMeth.Left
    f.Top = fraApplMeth.Top + fraApplMeth.Height
    f.Visible = True
  End If
End Sub

Private Sub optBoomHeightGnd_Click(Index As Integer)
  Dim iDSD As Integer
  Dim iBOOM As Integer
  Dim iPCT As Integer
  Dim c As Control
  If PropTakeAction Then
    'record boom selection
    For Each c In optBoomHeightGnd()
      If c.Value Then iBOOM = c.Index: Exit For
    Next
    'record DSD selection
    For Each c In optDropSizeGnd()
      If c.Value Then iDSD = c.Index: Exit For
    Next
    'record percentile selection
    For Each c In optPercentile()
      If c.Value Then iPCT = c.Index: Exit For
    Next
    'figure out the BasicType
    GroundType = (4 * iPCT) + (2 * iBOOM) + iDSD
    'display info
    lblInfoGnd.Caption = GetTier1Info(AM_GROUND, GroundType)
  End If
End Sub

Private Sub optDropSize_Click(Index As Integer)
  If PropTakeAction Then
    AerialType = Index
  End If
End Sub

Private Sub optDropSizeFS_Click(Index As Integer)
  If PropTakeAction Then
    AerialType = Index
  End If
End Sub

Private Sub optDropSizeGnd_Click(Index As Integer)
  If PropTakeAction Then
    'Do the same thing as the boom height
    optBoomHeightGnd_Click Index
  End If
End Sub

Private Sub optOrchard_Click(Index As Integer)
  Dim i As Integer
  If PropTakeAction Then
    OrchardType = Index
    Select Case OrchardType
    Case 0 To 2 'combo orchards
      For i = 3 To 13: optOrchard(i).Value = False: Next
    Case 3 To 13 'combo orchards
      For i = 0 To 2: optOrchard(i).Value = False: Next
    End Select
    lblInfoOrch.Caption = GetTier1Info(AM_ORCHARD, OrchardType)
  End If
End Sub

Private Sub optPercentile_Click(Index As Integer)
  If PropTakeAction Then
    'Do the same thing as the boom height
    optBoomHeightGnd_Click Index
  End If
End Sub

Private Sub txtEndSwath_LostFocus()
  If PropTakeAction Then
    OrchardEndSwath = Val(txtEndSwath.Text)
    'Clamp the new value
    If OrchardEndSwath < 1 Then OrchardEndSwath = 1
    If OrchardEndSwath > 20 Then OrchardEndSwath = 20
    If OrchardEndSwath < OrchardStartSwath Then OrchardEndSwath = OrchardStartSwath
    'Update the text in case the value changed
    PropTakeAction = False
    txtEndSwath.Text = Format$(OrchardEndSwath)
    PropTakeAction = True
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

Private Sub txtStartSwath_LostFocus()
  If PropTakeAction Then
    OrchardStartSwath = Val(txtStartSwath.Text)
    'Clamp the new value
    If OrchardStartSwath < 1 Then OrchardStartSwath = 1
    If OrchardStartSwath > 20 Then OrchardStartSwath = 20
    If OrchardStartSwath > OrchardEndSwath Then OrchardStartSwath = OrchardEndSwath
    'Update the text in case the value changed
    PropTakeAction = False
    txtStartSwath.Text = Format$(OrchardStartSwath)
    PropTakeAction = True
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

Private Sub txtSwathsGnd_LostFocus()
  If PropTakeAction Then
    GroundSwaths = Val(txtSwathsGnd.Text)
    'Clamp the new value
    If GroundSwaths < 1 Then GroundSwaths = 1
    If GroundSwaths > 20 Then GroundSwaths = 20
    'Update the text in case the value changed
    PropTakeAction = False
    txtSwathsGnd.Text = Format$(GroundSwaths)
    PropTakeAction = True
  End If
End Sub

Private Sub txtSwathsGnd_KeyPress(KeyAscii As Integer)
  If PropTakeAction Then
    If KeyAscii = Asc(vbCr) Then
      txtSwathsGnd_LostFocus
      KeyAscii = 0
    End If
  End If
End Sub

