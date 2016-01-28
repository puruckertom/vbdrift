VERSION 5.00
Begin VB.Form frmPrefs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   4065
   ClientLeft      =   1440
   ClientTop       =   2925
   ClientWidth     =   5505
   ForeColor       =   &H80000008&
   HelpContextID   =   1227
   Icon            =   "PREFS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4065
   ScaleWidth      =   5505
   Begin VB.Frame fraAudience 
      Caption         =   "Starting Mode"
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   2415
      Begin VB.OptionButton optAudience 
         Caption         =   "Forestry"
         Height          =   255
         HelpContextID   =   1227
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optAudience 
         Caption         =   "Agricultural"
         Height          =   255
         HelpContextID   =   1227
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraTier 
      Caption         =   "Starting Application Method"
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   2415
      Begin VB.OptionButton optStartTier 
         Caption         =   "Tier III Aerial"
         Height          =   255
         HelpContextID   =   1227
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optStartTier 
         Caption         =   "Tier II Aerial"
         Height          =   255
         HelpContextID   =   1227
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton optStartTier 
         Caption         =   "Tier I Aerial"
         Height          =   255
         HelpContextID   =   1227
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optStartTier 
         Caption         =   "Tier I Ground"
         Height          =   255
         HelpContextID   =   1227
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optStartTier 
         Caption         =   "Tier I Orchard/Airblast"
         Height          =   255
         HelpContextID   =   1227
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame fraUnits 
      Caption         =   "Units"
      Height          =   975
      Left            =   2640
      TabIndex        =   17
      Top             =   0
      Width           =   2775
      Begin VB.OptionButton optUnits 
         Caption         =   "Metric"
         Height          =   255
         HelpContextID   =   1227
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optUnits 
         Caption         =   "English"
         Height          =   255
         HelpContextID   =   1227
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1227
      Left            =   4560
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1227
      Left            =   3600
      TabIndex        =   0
      Top             =   3600
      Width           =   855
   End
   Begin VB.Frame fraMisc 
      Caption         =   "Misc"
      Height          =   1815
      Left            =   2640
      TabIndex        =   18
      Top             =   960
      Width           =   2775
      Begin VB.CheckBox cbxWarnTierChange 
         Caption         =   "Warn on Tier &change"
         Height          =   255
         HelpContextID   =   1227
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox cbxPauseBeforeCalc 
         Caption         =   "&Pause before calculating"
         Height          =   255
         HelpContextID   =   1227
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox cbxSuppressTier3Warn 
         Caption         =   "Suppress Calculation &Warnings"
         Height          =   255
         HelpContextID   =   1227
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Library"
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   5295
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   375
         HelpContextID   =   1505
         Left            =   4080
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtUserLib 
         Height          =   285
         HelpContextID   =   1505
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: prefs.frm,v 1.9 2001/05/24 20:16:23 tom Exp $
'some settings require the Tier input form to be redisplayed
Dim RedisplayTierForm As Integer

Private Sub cmdBrowse_Click()
  Dim s As String
  'open the file dialog and get a library name
  If FileDialog(FD_OPEN, FD_TYPE_LIB, s) Then 'if the selection is good
    txtUserLib.Text = s
  End If
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOk_Click()
  FormToData    'retrieve settings
  WriteGeneralPrefs    'write settings
  Me.Tag = Format$(RedisplayTierForm, "True/False") 'return status
  Me.Hide
End Sub

Private Sub DataToForm()
'Place data in the UP area in the form controls
  With UP
    optStartTier(0).Value = .InitialTier = TIER_1 And .InitialAM = AM_AERIAL
    optStartTier(1).Value = .InitialTier = TIER_1 And .InitialAM = AM_GROUND
    optStartTier(2).Value = .InitialTier = TIER_1 And .InitialAM = AM_ORCHARD
    optStartTier(3).Value = .InitialTier = TIER_2 And .InitialAM = AM_AERIAL
    optStartTier(4).Value = .InitialTier = TIER_3 And .InitialAM = AM_AERIAL
  End With
  optAudience(UP.InitialAUD).Value = True
  optUnits(UP.Units).Value = True
  If UP.WarnOnTierChange Then
    cbxWarnTierChange.Value = 1
  Else
    cbxWarnTierChange.Value = 0
  End If
  If UP.PauseBeforeCalc Then
    cbxPauseBeforeCalc.Value = 1
  Else
    cbxPauseBeforeCalc.Value = 0
  End If
  If UP.SuppressTier3Warn Then
    cbxSuppressTier3Warn.Value = 1
  Else
    cbxSuppressTier3Warn.Value = 0
  End If
  txtUserLib.Text = UP.UserLib
End Sub

Private Sub Form_Load()
'initialize this form
  CenterForm Me
  If Not AGDRIFTREGULATORY Then
    optStartTier(0).Visible = False 'T1A
    optStartTier(3).Visible = False 'T2A
  End If
  DataToForm
  Me.Tag = "False"
  RedisplayTierForm = False
End Sub

Private Sub FormToData()
'Place form data in global area
  Dim c As Control
  For Each c In optStartTier
    If c.Value Then
      Select Case c.Index
      Case 0 'Tier 1 Aerial
        UP.InitialTier = TIER_1
        UP.InitialAM = AM_AERIAL
      Case 1 'Tier 1 Ground
        UP.InitialTier = TIER_1
        UP.InitialAM = AM_GROUND
      Case 2 'Tier 1 Orchard
        UP.InitialTier = TIER_1
        UP.InitialAM = AM_ORCHARD
      Case 3 'Tier 2 Aerial
        UP.InitialTier = TIER_2
        UP.InitialAM = AM_AERIAL
      Case 4 'Tier 3 Aerial
        UP.InitialTier = TIER_3
        UP.InitialAM = AM_AERIAL
      End Select
      Exit For
    End If
  Next
  If optAudience(0).Value Then
    UP.InitialAUD = AUD_SDTF
  ElseIf optAudience(1).Value Then
    UP.InitialAUD = AUD_FS
  End If
  If optUnits(0).Value Then
    UP.Units = UN_IMPERIAL
  ElseIf optUnits(1).Value Then
    UP.Units = UN_METRIC
  End If
  UnitsSelectSystem UP.Units 'reset system units
  If cbxWarnTierChange.Value = 1 Then
    UP.WarnOnTierChange = True
  Else
    UP.WarnOnTierChange = False
  End If
  If cbxPauseBeforeCalc.Value = 1 Then
    UP.PauseBeforeCalc = True
  Else
    UP.PauseBeforeCalc = False
  End If
  If cbxSuppressTier3Warn.Value = 1 Then
    UP.SuppressTier3Warn = True
  Else
    UP.SuppressTier3Warn = False
  End If
  UP.UserLib = Trim$(txtUserLib.Text)
End Sub

Private Sub optAudience_Click(Index As Integer)
  'Make sure the Tier is >1 for forestry
  If Index = 1 Then 'forestry
    If optStartTier(0).Value Or _
       optStartTier(1).Value Or _
       optStartTier(2).Value Then
      optStartTier(3).Value = True 'switch to tier 2
      optStartTier(4).Value = True 'switch to aerial
    End If
  End If
  'Enable Tier 1's only for Agricultural mode
  optStartTier(0).Enabled = optAudience(0).Value
  optStartTier(1).Enabled = optAudience(0).Value
  optStartTier(2).Enabled = optAudience(0).Value
End Sub

Private Sub optStartTier_Click(Index As Integer)
  'Make sure Mode is Agricultural for Tier 1
  If Index <= 2 Then
    If optAudience(1).Value Then
      optAudience(0).Value = True 'switch to Agricultural
    End If
  End If
  'Enable Forestry only for Tier 2/3
  optAudience(1).Enabled = (optStartTier(3).Value Or _
                    optStartTier(4).Value)
End Sub

Private Sub optUnits_Click(Index As Integer)
  RedisplayTierForm = True 'Changing units requires a form reset
End Sub

