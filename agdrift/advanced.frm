VERSION 5.00
Begin VB.Form frmAdvanced 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Settings"
   ClientHeight    =   4500
   ClientLeft      =   2055
   ClientTop       =   2475
   ClientWidth     =   4905
   HelpContextID   =   1013
   Icon            =   "ADVANCED.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4500
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1013
      Left            =   3840
      TabIndex        =   1
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1013
      Left            =   2760
      TabIndex        =   0
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame fraAdvanced 
      Caption         =   "Advanced Settings"
      Height          =   3375
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   4695
      Begin VB.TextBox txtZref 
         Height          =   285
         HelpContextID   =   1503
         Left            =   3000
         TabIndex        =   23
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtMaxDownwindDist 
         Height          =   285
         HelpContextID   =   1435
         Left            =   3000
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtWindHeight 
         Height          =   285
         HelpContextID   =   1340
         Left            =   3000
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtMaxComputeTime 
         Height          =   285
         HelpContextID   =   1177
         Left            =   3000
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtVortexDecay 
         Height          =   285
         HelpContextID   =   1326
         Left            =   3000
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtDragCoeff 
         Height          =   285
         HelpContextID   =   1017
         Left            =   3000
         TabIndex        =   6
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtPropEff 
         Height          =   285
         HelpContextID   =   1229
         Left            =   3000
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtPressure 
         Height          =   285
         HelpContextID   =   1035
         Left            =   3000
         TabIndex        =   8
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Ground Reference:"
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label lblZrefUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   4200
         TabIndex        =   24
         Top             =   3000
         Width           =   330
      End
      Begin VB.Label lblMaxDownwindDistUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   4200
         TabIndex        =   22
         Top             =   1200
         Width           =   330
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Maximum Downwind Distance"
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Height for Wind Speed Measurement:"
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Maximum Computational Time:"
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Vortex Decay Rate:"
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Aircraft Drag Coefficient:"
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Propeller Efficiency:"
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Ambient Pressure:"
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Label lblWindHeightUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   4200
         TabIndex        =   14
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label13 
         Caption         =   "sec"
         Height          =   285
         Left            =   4200
         TabIndex        =   13
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblVortexDecayUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   4200
         TabIndex        =   12
         Top             =   1560
         Width           =   330
      End
      Begin VB.Label lblPressureUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   4200
         TabIndex        =   11
         Top             =   2640
         Width           =   330
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Knowledge of these parameters is essential before changing any of them."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: advanced.frm,v 1.7 2001/05/24 20:16:13 tom Exp $

Option Explicit

Private Sub DataToForm()
  txtWindHeight.Text = AGFormat$(UnitsDisplay(UD.MET.WindHeight, UN_LENGTH))
  txtMaxComputeTime.Text = AGFormat$(UD.CTL.MaxComputeTime)
  txtMaxDownwindDist.Text = AGFormat$(UnitsDisplay(UD.CTL.MaxDownwindDist, UN_LENGTH))
  txtVortexDecay.Text = AGFormat$(UnitsDisplay(UD.MET.VortexDecay, UN_SPEED))
  txtDragCoeff.Text = AGFormat$(UD.AC.DragCoeff)
  txtPropEff.Text = AGFormat$(UD.AC.PropEff)
  txtPressure.Text = AGFormat$(UnitsDisplay(UD.MET.Pressure, UN_AIRPRESSURE))
  txtZref.Text = AGFormat$(UnitsDisplay(UD.TRN.Zref, UN_LENGTH))
End Sub

Private Sub FormToData()
  UD.MET.WindHeight = UnitsInternal(Val(txtWindHeight.Text), UN_LENGTH)
  UD.CTL.MaxComputeTime = Val(txtMaxComputeTime.Text)
  UD.CTL.MaxDownwindDist = UnitsInternal(Val(txtMaxDownwindDist.Text), UN_LENGTH)
  UD.MET.VortexDecay = UnitsInternal(Val(txtVortexDecay.Text), UN_SPEED)
  UD.AC.DragCoeff = Val(txtDragCoeff.Text)
  UD.AC.PropEff = Val(txtPropEff.Text)
  UD.MET.Pressure = UnitsInternal(Val(txtPressure.Text), UN_AIRPRESSURE)
  UD.TRN.Zref = UnitsInternal(Val(txtZref.Text), UN_LENGTH)
  
  UpdateDataChangedFlag True 'Data was changed
  UC.Valid = False 'Calcs need to be redone
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  FormToData
  Unload Me
End Sub

Private Sub Form_Load()
  CenterForm Me
  
  'units
  lblWindHeightUnits = UnitsName(UN_LENGTH)
  lblMaxDownwindDistUnits = UnitsName(UN_LENGTH)
  lblVortexDecayUnits = UnitsName(UN_SPEED)
  lblPressureUnits = UnitsName(UN_AIRPRESSURE)
  lblZrefUnits = UnitsName(UN_LENGTH)
  
  DataToForm
End Sub

