VERSION 5.00
Begin VB.Form frmDropParam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametric Drop Size Distributiion"
   ClientHeight    =   3570
   ClientLeft      =   4695
   ClientTop       =   1905
   ClientWidth     =   5130
   ForeColor       =   &H80000008&
   HelpContextID   =   1460
   Icon            =   "DROPARAM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3570
   ScaleWidth      =   5130
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1460
      Left            =   3240
      TabIndex        =   0
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1460
      Left            =   4200
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame fraDisp 
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtRelSpan 
         Height          =   285
         HelpContextID   =   1460
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtVMD 
         Height          =   285
         HelpContextID   =   1460
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "V0.5"
         Height          =   195
         Left            =   1110
         TabIndex        =   11
         Top             =   315
         Width           =   405
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "D          :"
         Height          =   195
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblInput1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Relative Span:"
         Height          =   195
         Left            =   930
         TabIndex        =   10
         Top             =   765
         Width           =   1050
      End
      Begin VB.Label lblUnits0 
         AutoSize        =   -1  'True
         Caption         =   "µm"
         Height          =   195
         Left            =   2970
         TabIndex        =   9
         Top             =   285
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Conversion"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   5295
      Begin VB.CheckBox chkConvert 
         Caption         =   "Convert PMS to Malvern"
         Height          =   255
         HelpContextID   =   1460
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output"
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   4935
      Begin VB.OptionButton optSprayType 
         Caption         =   "Drop Size Classification"
         Height          =   255
         HelpContextID   =   1460
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox chkSwathDispAdjust 
         Caption         =   "Adjust Swath Displacement"
         Height          =   255
         HelpContextID   =   1460
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optSprayType 
         Caption         =   "Drop Size Distribution (Standard)"
         Height          =   255
         HelpContextID   =   1460
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton optSprayType 
         Caption         =   "Drop Size Distribution (Optimized)"
         Height          =   255
         HelpContextID   =   1460
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmDropParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: droparam.frm,v 1.9 2011/12/29 15:48:29 tom Exp $
Public Canceled As Boolean '"return" value for this form

Private Sub cmdCancel_Click()
  Canceled = True
  Hide
End Sub

Private Sub cmdOk_Click()
  If Val(txtVMD.Text) <= 0 Then
    MsgBox "Invalid Dv0.5", vbExclamation
    Exit Sub
  End If
  If Val(txtRelSpan.Text) <= 0 Then
    MsgBox "Invalid Relative Span", vbExclamation
    Exit Sub
  End If

  Canceled = False
  Hide
End Sub

Private Sub Form_Load()
  CenterForm Me
  
  chkConvert.Value = 0 'Don't do conversion
  chkConvert.Enabled = False 'Don't allow the user to change the setting
  optSprayType(0).Value = True 'Default to Spray Quality
  
  'Don't allow DSD output below Tier 3
  If UD.Tier < TIER_3 Then
    optSprayType(1).Enabled = False
    optSprayType(2).Enabled = False
  End If

  Canceled = True 'default return value
End Sub

