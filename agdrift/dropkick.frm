VERSION 5.00
Begin VB.Form frmDropKick 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DropKick"
   ClientHeight    =   6795
   ClientLeft      =   2115
   ClientTop       =   1560
   ClientWidth     =   6810
   ForeColor       =   &H80000008&
   HelpContextID   =   1090
   Icon            =   "DROPKICK.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6795
   ScaleWidth      =   6810
   Begin VB.Frame fraNozzle 
      Caption         =   "Nozzle"
      Height          =   1695
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtSprayAngle 
         Height          =   285
         HelpContextID   =   1476
         Left            =   4320
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtDiam 
         Height          =   285
         HelpContextID   =   1131
         Left            =   4320
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtVMD 
         Height          =   285
         HelpContextID   =   1327
         Left            =   4320
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtRelSpan 
         Height          =   285
         HelpContextID   =   1475
         Left            =   4320
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optNozzleType 
         Caption         =   "Library"
         Height          =   255
         HelpContextID   =   1090
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optNozzleType 
         Caption         =   "User-defined"
         Height          =   255
         HelpContextID   =   1090
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Spray Angle:"
         Height          =   195
         Left            =   3360
         TabIndex        =   58
         Top             =   1005
         Width           =   900
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "deg"
         Height          =   195
         Left            =   5520
         TabIndex        =   57
         Top             =   990
         Width           =   270
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "V0.5"
         Height          =   195
         Left            =   3795
         TabIndex        =   44
         Top             =   330
         Width           =   405
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "D          :"
         Height          =   195
         Left            =   3645
         TabIndex        =   28
         Top             =   255
         Width           =   615
      End
      Begin VB.Label lblLibNozzle 
         AutoSize        =   -1  'True
         Caption         =   "Nozzle name"
         Height          =   195
         Left            =   480
         TabIndex        =   43
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "µm"
         Height          =   195
         Left            =   5520
         TabIndex        =   42
         Top             =   285
         Width           =   255
      End
      Begin VB.Label lblEffDiamUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   5505
         TabIndex        =   34
         Top             =   1350
         Width           =   330
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Effective Nozzle Diameter:"
         Height          =   195
         Left            =   2010
         TabIndex        =   29
         Top             =   1365
         Width           =   2280
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Relative Span:"
         Height          =   195
         Left            =   3225
         TabIndex        =   27
         Top             =   660
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1090
      Left            =   6000
      TabIndex        =   1
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1090
      Left            =   5160
      TabIndex        =   0
      Top             =   6360
      Width           =   735
   End
   Begin VB.Frame fraMaterial 
      Caption         =   "Spray Material"
      Height          =   1695
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   6615
      Begin VB.TextBox txtElongVisc 
         Height          =   285
         HelpContextID   =   1307
         Left            =   4320
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtShearVisc 
         Height          =   285
         HelpContextID   =   1253
         Left            =   4320
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtDynSurfTens 
         Height          =   285
         HelpContextID   =   1125
         Left            =   4320
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtDensity 
         Height          =   285
         HelpContextID   =   1255
         Left            =   4320
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton optMatType 
         Caption         =   "Library"
         Height          =   255
         HelpContextID   =   1090
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optMatType 
         Caption         =   "User-defined"
         Height          =   255
         HelpContextID   =   1090
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "cp"
         Height          =   195
         Left            =   5460
         TabIndex        =   45
         Top             =   1005
         Width           =   225
      End
      Begin VB.Label lblLibMaterial 
         AutoSize        =   -1  'True
         Caption         =   "Material name"
         Height          =   195
         Left            =   480
         TabIndex        =   35
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "cp"
         Height          =   195
         Left            =   5460
         TabIndex        =   37
         Top             =   645
         Width           =   225
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "dynes/cm"
         Height          =   195
         Left            =   5475
         TabIndex        =   36
         Top             =   285
         Width           =   840
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Elongational Viscosity:"
         Height          =   195
         Left            =   2310
         TabIndex        =   33
         Top             =   1005
         Width           =   1905
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Shear Viscosity:"
         Height          =   195
         Left            =   2895
         TabIndex        =   32
         Top             =   645
         Width           =   1320
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dynamic Surface Tension:"
         Height          =   195
         Left            =   2025
         TabIndex        =   31
         Top             =   300
         Width           =   2190
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Specific Gravity:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2790
         TabIndex        =   30
         Top             =   1380
         Width           =   1425
      End
   End
   Begin VB.Frame fraSprayData 
      Caption         =   "Spray Data"
      Height          =   2175
      Left            =   120
      TabIndex        =   25
      Top             =   3480
      Width           =   6615
      Begin VB.OptionButton optFlowType 
         Caption         =   "Input"
         Height          =   255
         HelpContextID   =   1150
         Index           =   1
         Left            =   840
         TabIndex        =   18
         Top             =   1800
         Width           =   735
      End
      Begin VB.OptionButton optFlowType 
         Caption         =   "Scaled"
         Height          =   255
         HelpContextID   =   1150
         Index           =   0
         Left            =   840
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtFlow 
         Height          =   285
         HelpContextID   =   1150
         Left            =   1680
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtPressure 
         Height          =   285
         HelpContextID   =   1442
         Left            =   1680
         TabIndex        =   16
         Top             =   990
         Width           =   1095
      End
      Begin VB.TextBox txtAngle 
         Height          =   285
         HelpContextID   =   1183
         Left            =   1650
         TabIndex        =   15
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtSpeed 
         Height          =   285
         HelpContextID   =   1310
         Left            =   1650
         TabIndex        =   14
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lblNote 
         Alignment       =   2  'Center
         Caption         =   "Specific Gravity, Air Speed, and Spray Volume Rate must be consistent with drift model inputs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   56
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label lblSprayRate 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4710
         TabIndex        =   55
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblSprayRateUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   5910
         TabIndex        =   54
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "Spray Volume Rate per Nozzle:"
         Height          =   555
         Left            =   3480
         TabIndex        =   53
         Top             =   1500
         Width           =   1170
      End
      Begin VB.Label lblFlightSpeed 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4680
         TabIndex        =   52
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblFlightSpeedUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   5880
         TabIndex        =   51
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flight Speed:"
         Height          =   195
         Left            =   3690
         TabIndex        =   50
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblFlowUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   49
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label lblFlowScaled 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   48
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Spray Volume Rate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   47
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label lblFlowUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Index           =   0
         Left            =   2880
         TabIndex        =   46
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label lblPressureUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2880
         TabIndex        =   22
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "deg"
         Height          =   195
         Left            =   2880
         TabIndex        =   26
         Top             =   675
         Width           =   330
      End
      Begin VB.Label lblAirSpeedUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2880
         TabIndex        =   41
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pressure:"
         Height          =   195
         Left            =   810
         TabIndex        =   40
         Top             =   1035
         Width           =   810
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nozzle Orientation:"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   690
         Width           =   1380
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Air Speed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   38
         Top             =   330
         Width           =   900
      End
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output"
      Height          =   1095
      Left            =   120
      TabIndex        =   59
      Top             =   5640
      Width           =   4935
      Begin VB.CheckBox chkSwathDispAdjust 
         Caption         =   "Adjust Swath Displacement"
         Height          =   255
         HelpContextID   =   1090
         Left            =   2520
         TabIndex        =   61
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optSprayType 
         Caption         =   "Drop Size Classification"
         Height          =   255
         HelpContextID   =   1090
         Index           =   0
         Left            =   240
         TabIndex        =   60
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optSprayType 
         Caption         =   "Drop Size Distribution (Optimized)"
         Height          =   255
         HelpContextID   =   1090
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton optSprayType 
         Caption         =   "Drop Size Distribution (Standard)"
         Height          =   255
         HelpContextID   =   1090
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmDropKick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: dropkick.frm,v 1.11 2001/05/24 20:16:18 tom Exp $
'this flag is used to tell the option buttons not to
'take action on their new values. This is required
'to differentiate between programatic state changes
'and user actions
Dim OptTakeAction As Integer  'if true, execute automatic change-related code
                              'for Spray Mat option button
Dim PropTakeAction As Integer 'if true, execute automatic change-related code
                              'for Property text boxes
Dim SaveNozType As Integer
Dim SaveMatType As Integer

Private Sub ChangeMatType(NewType As Integer)
'Select a new Material Type and do what is necessary to
'get new data
  Me.MousePointer = vbHourglass 'change pointer to hourglass
  Select Case NewType
    Case 0 'user-def
      'nothing to do here, except a form cleanup on exit
    Case 1 'library
      'turn off Property actions so that the lib form can
      'modify this form's controls without triggering other actions
      temp = PropTakeAction  'save flag state
      PropTakeAction = False 'disable actions on change
      frmDKMatLib.Show vbModal  'get the properties from lib
      PropTakeAction = temp  'restore flag value
      If frmDKMatLib.Tag = "False" Then 'Tag holds status info
        'reset original dist type
        temp = OptTakeAction  'save flag state
        OptTakeAction = False 'disable actions on change
        optMatType(SaveMatType).Value = True  'reset option button
        OptTakeAction = temp  'restore flag value
      End If
      Unload frmDKMatLib
  End Select
  'adjust the Type controls to reflect the current type
  UpdateTypeControls
  UpdatePropertyControls
  Me.MousePointer = vbDefault 'change pointer back to default
End Sub

Private Sub ChangeNozType(NewType As Integer)
'Select a new Nozzle Type and do what is necessary to
'get new data
  Me.MousePointer = vbHourglass 'change pointer to hourglass
  Select Case NewType
    Case 0 'user-def
      'nothing to do here, except a form cleanup on exit
    Case 1 'library
      'turn off Property actions so that the lib form can
      'modify this form's controls without triggering other actions
      temp = PropTakeAction  'save flag state
      PropTakeAction = False 'disable actions on change
      frmDKNozLib.Show vbModal  'get the properties from lib
      PropTakeAction = temp  'restore flag value
      If frmDKNozLib.Tag = "False" Then 'Tag holds status info
        'reset original dist type
        temp = OptTakeAction  'save flag state
        OptTakeAction = False 'disable actions on change
        optNozzleType(SaveNozType).Value = True  'reset option button
        OptTakeAction = temp  'restore flag value
      End If
      Unload frmDKNozLib
  End Select
  'adjust the Type controls to reflect the current type
  UpdateScaledFlow
  UpdateTypeControls
  UpdatePropertyControls
  Me.MousePointer = vbDefault 'change pointer back to default
End Sub

Private Sub cmdCancel_Click()
  Me.Tag = "False" 'return cancellation status
  Me.Hide
End Sub

Private Sub cmdOk_Click()
'execute DropKick calcs
  Dim DKstat As Integer
  Dim MaxErrLev As Integer
  Dim squal As Integer
  Dim np As Integer
  ReDim Diam(MAX_DROPS - 1) As Single
  ReDim mfrac(MAX_DROPS - 1) As Single
  Dim xDSD As DropSizeDistData
  'Change the form mouse pointer
  Me.MousePointer = vbHourglass 'hourglass
  'store the new form data
  FormToData
  'do the calcs
  DKstat = CalcDropKick(DK2, True, MaxErrLev, squal, np, Diam(), mfrac())
  'Change the form mouse pointer back
  Me.MousePointer = vbDefault 'default
  '
  If DKstat Then   'exit only if successful
    If squal < 0 Then
      'transfer new drop dist to parent form
      frmDropDist.ArrayToGrid np, Diam(), mfrac()
    Else
      'Spray Quality: get basic dsd and return to parent form
      frmDropDist.BasicDistToGrid squal
      'adjust swath displacement if required
      If chkSwathDispAdjust.Value = 1 Then
        frmDropDist.AdjustSwathDispFlag = True
        GetBasicDataDSDSwathDisp squal, _
          frmDropDist.AdjustSwathDispValue
      Else
        frmDropDist.AdjustSwathDispFlag = False
        frmDropDist.AdjustSwathDispValue = 0
      End If
    End If
    DK2.MaxErrorLevel = MaxErrLev
    Me.Tag = "True" 'return success status
    Me.Hide
  End If
End Sub

Private Sub DataToForm()
'transfer user data to form controls for editing
'DK2 is a global structure that is assumed to contain
'the dropkick data on which to operate
  
  'disable action flags
  tempopt = OptTakeAction                   'save flag value
  OptTakeAction = False                     'disable actions for the following
  tempprop = PropTakeAction                 'save flag value
  PropTakeAction = False                    'allow raw field modification
  
  'Nozzle frame
  optNozzleType(DK2.NozType) = True
  lblLibNozzle.Caption = Trim$(DK2.NameNoz)
  txtVMD.Text = AGFormat$(DK2.VMD)
  txtRelSpan.Text = AGFormat$(DK2.RelSpan)
  txtDiam.Text = AGFormat$(UnitsDisplay(DK2.EffDiam, UN_SMLENGTH2))
  txtSprayAngle.Text = AGFormat$(DK2.SprayAngle)

  'Material frame
  optMatType(DK2.MatType) = True
  lblLibMaterial.Caption = Trim$(DK2.NameMat)
  txtDynSurfTens.Text = AGFormat$(DK2.DynSurfTens)
  txtShearVisc.Text = AGFormat$(DK2.ShearVisc)
  txtDensity.Text = AGFormat$(DK2.Density)
  txtElongVisc.Text = AGFormat$(DK2.ElongVisc)

  'Spray Data frame
  txtSpeed.Text = AGFormat$(UnitsDisplay(DK2.Speed, UN_SPEED))
  txtAngle.Text = AGFormat$(DK2.NozAngle)
  txtPressure.Text = AGFormat$(UnitsDisplay(DK2.Pressure, UN_PRESSURE))
  optFlowType(DK2.FlowType).Value = True
  If DK2.FlowType = 0 Then
    txtFlow.Text = ""
  Else
    txtFlow.Text = AGFormat$(UnitsDisplay(DK2.flow, UN_FLOWRATE))
  End If
  
  'Output frame
  optSprayType(DK2.SprayType).Value = True
  
  'adjust form controls and restore flags
  UpdateScaledFlow
  UpdateTypeControls
  OptTakeAction = tempopt                       'restore flag value
  PropTakeAction = tempprop                     'restore flag value
End Sub

Private Sub Form_Load()
'Initialize the controls on this form

  'center the form
  CenterForm Me

  Me.Tag = "False"  'default return value
  
  lblEffDiamUnits.Caption = UnitsName(UN_SMLENGTH)
  lblAirSpeedUnits.Caption = UnitsName(UN_SPEED)
  lblFlightSpeed.Caption = AGFormat$(UnitsDisplay(UD.AC.TypSpeed, UN_SPEED))
  lblFlightSpeedUnits.Caption = UnitsName(UN_SPEED)
  lblPressureUnits.Caption = UnitsName(UN_PRESSURE)
  lblFlowUnits(0).Caption = UnitsName(UN_FLOWRATE)
  lblFlowUnits(1).Caption = UnitsName(UN_FLOWRATE)
  lblSprayRateUnits.Caption = UnitsName(UN_FLOWRATE)
  
  'Don't allow DSD output below Tier 3
  If UD.Tier < TIER_3 Then
    optSprayType(1).Enabled = False
    optSprayType(2).Enabled = False
  End If

  'allow option button changes to take action
  '(see declarations section)
  OptTakeAction = True
  PropTakeAction = True

  'fill the controls with user data
  DataToForm
End Sub

Private Sub FormToData()
'Place the form data in user data storage
  
  Dim nlong As Long
  Dim c As Control
  
  'find the current nozzle type selection
  For i = 0 To 1
    If optNozzleType(i).Value = True Then DT = i
  Next
  'save the current type selection
  DK2.NozType = DT
  
  DK2.NameNoz = lblLibNozzle.Caption
  DK2.VMD = Val(txtVMD.Text)
  DK2.RelSpan = Val(txtRelSpan.Text)
  DK2.EffDiam = UnitsInternal(Val(txtDiam.Text), UN_SMLENGTH2)
  DK2.SprayAngle = Val(txtSprayAngle.Text)

  'find the current type selection
  For i = 0 To 1
    If optMatType(i).Value = True Then DT = i
  Next
  'save the current type selection
  DK2.MatType = DT

  DK2.NameMat = lblLibMaterial.Caption
  DK2.DynSurfTens = Val(txtDynSurfTens.Text)
  DK2.ShearVisc = Val(txtShearVisc.Text)
  DK2.Density = Val(txtDensity.Text)
  DK2.ElongVisc = Val(txtElongVisc.Text)

  DK2.Speed = UnitsInternal(Val(txtSpeed.Text), UN_SPEED)
  DK2.NozAngle = Val(txtAngle.Text)
  DK2.Pressure = UnitsInternal(Val(txtPressure.Text), UN_PRESSURE)
  
  If optFlowType(0).Value Then
    DK2.FlowType = 0
    DK2.flow = Val(UnitsInternal(lblFlowScaled.Caption, UN_FLOWRATE))
  Else
    DK2.FlowType = 1
    DK2.flow = Val(UnitsInternal(txtFlow.Text, UN_FLOWRATE))
  End If

  For Each c In optSprayType()
    If c.Value Then
      DK2.SprayType = c.Index
      Exit For
    End If
  Next
End Sub

Private Sub optFlowType_Click(Index As Integer)
  If PropTakeAction Then
    If Index = 0 Then
      PropTakeAction = False
      txtFlow.Text = ""
      PropTakeAction = True
    End If
  End If
End Sub

Private Sub optMatType_Click(Index As Integer)
  If OptTakeAction Then ChangeMatType Index
End Sub

Private Sub optMatType_DblClick(Index As Integer)
  If OptTakeAction Then ChangeMatType Index
End Sub

Private Sub optNozzleType_Click(Index As Integer)
  If OptTakeAction Then ChangeNozType Index
End Sub

Private Sub optNozzleType_DblClick(Index As Integer)
  If OptTakeAction Then ChangeNozType Index
End Sub

Private Sub optSprayType_Click(Index As Integer)
  If Index = 0 Then
    chkSwathDispAdjust.Enabled = True
  Else
    chkSwathDispAdjust.Enabled = False
  End If
End Sub

Private Sub txtDensity_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optMatType(0).Value Then optMatType(0).Value = True
    UpdateScaledFlow
  End If
End Sub

Private Sub txtDiam_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optNozzleType(0).Value Then optNozzleType(0).Value = True
    UpdateScaledFlow
  End If
End Sub

Private Sub txtDynSurfTens_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optMatType(0).Value Then optMatType(0).Value = True
  End If
End Sub

Private Sub txtElongVisc_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optMatType(0).Value Then optMatType(0).Value = True
  End If
End Sub

Private Sub txtFlow_Change()
  If PropTakeAction Then
    If Not optFlowType(1).Value Then
      optFlowType(1).Value = True
    End If
  End If
End Sub

Private Sub txtPressure_Change()
  If PropTakeAction Then
    UpdateScaledFlow
  End If
End Sub

Private Sub txtRelSpan_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optNozzleType(0).Value Then optNozzleType(0).Value = True
  End If
End Sub

Private Sub txtShearVisc_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optMatType(0).Value Then optMatType(0).Value = True
  End If
End Sub

Private Sub txtSprayAngle_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optNozzleType(0).Value Then optNozzleType(0).Value = True
  End If
End Sub

Private Sub txtVMD_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optNozzleType(0).Value Then optNozzleType(0).Value = True
  End If
End Sub

Private Sub UpdatePropertyControls()
'update the state of the property controls

End Sub

Private Sub UpdateScaledFlow()
'Update the value of the scaled flow rate
  Dim Swath As Single
  Dim QdotN As Single
  Dim Qdot As Single
  Dim Press As Single
  Dim Deff As Single
  Dim SpecGrav As Single
  
  If UD.CTL.SwathWidthType = 0 Then
    Swath = UD.CTL.SwathWidth
  Else
    Swath = 2 * UD.AC.SemiSpan * UD.CTL.SwathWidth
  End If
  Press = UnitsInternal(Val(txtPressure.Text), UN_PRESSURE)
  Deff = UnitsInternal(Val(txtDiam.Text), UN_SMLENGTH2)
  SpecGrav = Val(txtDensity.Text)
  
  QdotN = 0.006 * UD.SM.FlowRate * Swath * UD.AC.TypSpeed / UD.NZ.NumNoz
  lblSprayRate.Caption = AGFormat$(UnitsDisplay(QdotN, UN_FLOWRATE))
  
  If SpecGrav > 0 Then
    Qdot = 66.6432 * Deff * Deff * Sqr(Press / SpecGrav)
    lblFlowScaled.Caption = AGFormat$(UnitsDisplay(Qdot, UN_FLOWRATE))
  Else
    lblFlowScaled.Caption = "0"
  End If
End Sub

Private Sub UpdateTypeControls()
'Adjust the Type Controls to conform to the current
'setting of the Type Option buttons

  'Nozzle--------------------------------
  'find the current nozzle type selection
  For i = 0 To 1
    If optNozzleType(i).Value = True Then DT = i
  Next
  'save the current type selection
  SaveNozType = DT

  'clear the library label
  'set related controls to a known state
  lblLibNozzle.Visible = False
  'make adjustments
  Select Case DT
    Case 0 'user-def
    Case 1 'library
      lblLibNozzle.Visible = True
  End Select

  
  'Spray Material--------------------------
  'find the current material type selection
  For i = 0 To 1
    If optMatType(i).Value = True Then DT = i
  Next
  'save the current type selection
  SaveMatType = DT

  'clear the library label
  'set related controls to a known state
  lblLibMaterial.Visible = False
  'make adjustments
  Select Case DT
    Case 0 'user-def
    Case 1 'library
      lblLibMaterial.Visible = True
  End Select

End Sub

