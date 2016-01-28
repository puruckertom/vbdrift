VERSION 5.00
Begin VB.Form frmDropKirk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "USDA ARS Nozzle Models"
   ClientHeight    =   5640
   ClientLeft      =   1035
   ClientTop       =   2640
   ClientWidth     =   6825
   ForeColor       =   &H80000008&
   HelpContextID   =   1461
   Icon            =   "DROPKIRK.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5640
   ScaleWidth      =   6825
   Begin VB.Frame fraNozzle 
      Caption         =   "Nozzle"
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtOrifice 
         Height          =   285
         HelpContextID   =   1394
         Left            =   3480
         TabIndex        =   4
         Text            =   "txtOrifice"
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cboOrifice 
         Height          =   315
         HelpContextID   =   1394
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cboNozzleName 
         Height          =   315
         HelpContextID   =   1438
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lblOrificeUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2880
         TabIndex        =   27
         Top             =   720
         Width           =   330
      End
      Begin VB.Label lblOrifice 
         AutoSize        =   -1  'True
         Caption         =   "lblOrifice"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1461
      Left            =   6000
      TabIndex        =   1
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1461
      Left            =   5160
      TabIndex        =   0
      Top             =   5160
      Width           =   735
   End
   Begin VB.Frame fraMaterial 
      Caption         =   "Spray Material"
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   6615
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.0"
         Height          =   255
         Left            =   1680
         TabIndex        =   31
         Top             =   720
         Width           =   630
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
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tap Water with 0.25% v/v Triton X-100"
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   360
         Width           =   3390
      End
   End
   Begin VB.Frame fraSprayData 
      Caption         =   "Spray Data"
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   6615
      Begin VB.TextBox txtNozzleAngle 
         Height          =   285
         HelpContextID   =   1183
         Left            =   3240
         TabIndex        =   7
         Text            =   "txtNozzleAngle"
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cboNozzleAngle 
         Height          =   315
         HelpContextID   =   1183
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtPressure 
         Height          =   285
         HelpContextID   =   1442
         Left            =   1560
         TabIndex        =   8
         Top             =   990
         Width           =   1095
      End
      Begin VB.TextBox txtSpeed 
         Height          =   285
         HelpContextID   =   1310
         Left            =   1530
         TabIndex        =   5
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lblNote 
         Alignment       =   2  'Center
         Caption         =   "Specific Gravity and Air Speed must be consistent with drift model inputs."
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
         Left            =   3360
         TabIndex        =   23
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lblFlightSpeed 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4680
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblFlightSpeedUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   5880
         TabIndex        =   21
         Top             =   360
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flight Speed:"
         Height          =   195
         Left            =   3690
         TabIndex        =   20
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblPressureUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2760
         TabIndex        =   11
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label lblNozzleAngleUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2760
         TabIndex        =   15
         Top             =   675
         Width           =   330
      End
      Begin VB.Label lblAirSpeedUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2760
         TabIndex        =   19
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pressure:"
         Height          =   195
         Left            =   690
         TabIndex        =   18
         Top             =   1035
         Width           =   810
      End
      Begin VB.Label lblNozzleAngle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "lblNozzleAngle"
         Height          =   195
         Left            =   465
         TabIndex        =   17
         Top             =   660
         Width           =   1035
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
         Left            =   600
         TabIndex        =   16
         Top             =   330
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Conversion"
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   3840
      Width           =   6615
      Begin VB.CheckBox chkConvert 
         Caption         =   "Convert PMS to Malvern"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output"
      Height          =   1095
      Left            =   120
      TabIndex        =   24
      Top             =   4440
      Width           =   4935
      Begin VB.CheckBox chkSwathDispAdjust 
         Caption         =   "Adjust Swath Displacement"
         Height          =   255
         HelpContextID   =   1461
         Left            =   2520
         TabIndex        =   33
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optSprayType 
         Caption         =   "Drop Size Classification"
         Height          =   255
         HelpContextID   =   1461
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optSprayType 
         Caption         =   "Drop Size Distribution (Optimized)"
         Height          =   255
         HelpContextID   =   1461
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   3255
      End
      Begin VB.OptionButton optSprayType 
         Caption         =   "Drop Size Distribution (Standard)"
         Height          =   255
         HelpContextID   =   1461
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmDropKirk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: dropkirk.frm,v 1.9 2006/11/08 15:18:10 tom Exp $
Option Explicit

'this flag is used to tell the option buttons not to
'take action on their new values. This is required
'to differentiate between programatic state changes
'and user actions
Private OptTakeAction As Integer  'if true, execute automatic change-related code
                              'for Spray Mat option button
Private PropTakeAction As Integer 'if true, execute automatic change-related code
                              'for Property text boxes

'These values remember the properties of the ARS Nozzle
'read from the library
Private OfcUnits As Integer  'Orifice units 0=number 1=inches
Private OfcLabel As String   'Name of orifice
Private OfcPrefix As String  'Orifice prefix
Private NumOfc As Integer    'Number of values (0=range)
Private OfcVal(19) As Single 'Orifice values (if range, (0) and (1) are min, max)
Private ModUnits As Integer  'Modifier units 0=number 1=degrees
Private ModLabel As String   'Name of Modifier
Private ModPrefix As String  'Modifier prefix
Private NumMod As Integer    'Number of Modifier values (0=range)
Private ModVal(19) As Single 'Modifier values (if range, (0) and (1) are min, max)

Private Sub DataToForm()
'transfer user data to form controls for editing
'BK2 is a global structure that is assumed to contain
'the dropkirk data on which to operate
  
  Dim i As Integer
  Dim s As String
  
  'Retrieve Nozzle data from the library
  'and store it in module-wide variables
  GetARSNozData BK2.NozType, _
    OfcUnits, OfcLabel, OfcPrefix, NumOfc, OfcVal(), _
    ModUnits, ModLabel, ModPrefix, NumMod, ModVal()

  'Nozzle frame
  cboNozzleName.ListIndex = BK2.NozType 'sets up other controls as well
  
  'Try to set Orifice and Nozzle Angle
  If OfcUnits = 0 Then  'orifice number
    s = OfcPrefix & Format$(BK2.Orifice)
  Else
    s = OfcPrefix & Format$(BK2.Orifice, "0.000")
  End If
  If NumOfc = 0 Then 'range
    txtOrifice.Text = s
    txtOrifice_LostFocus 'range checking
  Else 'discrete values
    With cboOrifice
      .ListIndex = 0
      For i = 0 To .ListCount - 1
        If .List(i) = s Then
          .ListIndex = i
          Exit For
        End If
      Next
    End With
  End If
  
  s = ModPrefix & Format$(BK2.NozAngle)
  If NumMod = 0 Then 'range
    txtNozzleAngle.Text = s
    txtNozzleAngle_LostFocus 'range checking
  Else
    With cboNozzleAngle
      .ListIndex = 0
      For i = 0 To .ListCount - 1
        If .List(i) = s Then
          .ListIndex = i
          Exit For
        End If
      Next
    End With
  End If
  
  'Spray Data frame
  txtSpeed.Text = AGFormat$(UnitsDisplay(BK2.Speed, UN_SPEED))
  txtPressure.Text = AGFormat$(UnitsDisplay(BK2.Pressure, UN_PRESSURE))
  
  'Data Source
  'For ARS Nozzle Models, all spectrum data is PMS. The SpectrumSource
  'flag indicates whether or not to convert it to Malvern.
  chkConvert.Value = BK2.SpectrumSource '0=don't convert 1=convert
  
  'Output frame
  optSprayType(BK2.SprayType).Value = True
  
End Sub

Private Sub FormToData()
'Place the form data in user data storage
  
  Dim nlong As Long
  Dim c As Control
  
  BK2.NozType = cboNozzleName.ListIndex
  BK2.NameNoz = cboNozzleName.Text
  If NumOfc = 0 Then 'range
    BK2.Orifice = Val(Mid$(txtOrifice.Text, Len(OfcPrefix) + 1)) 'strip off prefix
  Else
    BK2.Orifice = Val(Mid$(cboOrifice.Text, Len(OfcPrefix) + 1)) 'strip off prefix
  End If

  BK2.Speed = UnitsInternal(Val(txtSpeed.Text), UN_SPEED)
  If NumMod = 0 Then
    BK2.NozAngle = Val(Mid$(txtNozzleAngle.Text, Len(ModPrefix) + 1)) 'strip off prefix
  Else
    BK2.NozAngle = Val(Mid$(cboNozzleAngle.Text, Len(ModPrefix) + 1)) 'strip off prefix
  End If
  BK2.Pressure = UnitsInternal(Val(txtPressure.Text), UN_PRESSURE)
  
  BK2.SpectrumSource = chkConvert.Value
  
  For Each c In optSprayType()
    If c.Value Then
      BK2.SprayType = c.Index
      Exit For
    End If
  Next
End Sub

Private Sub cboNozzleName_Click()
'Set up the contents of the Nozzle controls to
'reflect that choices available for the named nozzle
  
  Dim i As Integer
  Dim SaveOrifice As String
  Dim SaveNozzleAngle As String

  'Save the current values (without prefix) depending on which control is in use
  If cboOrifice.Visible Then SaveOrifice = Mid$(cboOrifice.Text, Len(OfcPrefix) + 1) 'strip off prefix
  If txtOrifice.Visible Then SaveOrifice = Mid$(txtOrifice.Text, Len(OfcPrefix) + 1) 'strip off prefix
  If cboNozzleAngle.Visible Then SaveNozzleAngle = Mid$(cboNozzleAngle.Text, Len(ModPrefix) + 1) 'strip off prefix
  If txtNozzleAngle.Visible Then SaveNozzleAngle = Mid$(txtNozzleAngle.Text, Len(ModPrefix) + 1) 'strip off prefix
  
  'Retrieve Nozzle data from the library
  'and store it in module-wide variables
  GetARSNozData cboNozzleName.ListIndex, _
    OfcUnits, OfcLabel, OfcPrefix, NumOfc, OfcVal(), _
    ModUnits, ModLabel, ModPrefix, NumMod, ModVal()

  'Redo labels and units
  lblOrifice.Caption = OfcLabel
  If OfcUnits = 0 Then
    lblOrificeUnits.Caption = ""
  Else
    lblOrificeUnits.Caption = "in"
  End If
  If NumOfc = 0 Then 'continuous range
    cboOrifice.Visible = False
    txtOrifice.Visible = True
  Else 'discrete values
    txtOrifice.Visible = False
    cboOrifice.Visible = True
    cboOrifice.Clear
    For i = 0 To NumOfc - 1
      cboOrifice.AddItem OfcPrefix & Format$(OfcVal(i), "0.000")
    Next
  End If
  
  lblNozzleAngle.Caption = ModLabel
  If ModUnits = 0 Then
    lblNozzleAngleUnits = ""
  Else
    lblNozzleAngleUnits = "deg"
  End If
  If NumMod = 0 Then 'continuous range
    cboNozzleAngle.Visible = False
    txtNozzleAngle.Visible = True
  Else 'discrete values
    txtNozzleAngle.Visible = False
    cboNozzleAngle.Visible = True
    cboNozzleAngle.Clear
    For i = 0 To NumMod - 1
      cboNozzleAngle.AddItem ModPrefix & Format$(ModVal(i))
    Next
  End If
  
  'Try to restore Orifice and Nozzle Angle
  'This always works because we deal with strings
  'and let the controls adjust units and check limits
  If NumOfc = 0 Then 'range
    txtOrifice.Text = OfcPrefix & SaveOrifice
    txtOrifice_LostFocus 'range checking
  Else 'discrete values
    With cboOrifice
      .ListIndex = 0
      For i = 0 To .ListCount - 1
        If .List(i) = OfcPrefix & SaveOrifice Then
          .ListIndex = i
          Exit For
        End If
      Next
    End With
  End If
  If NumMod = 0 Then 'range
    txtNozzleAngle.Text = ModPrefix & SaveNozzleAngle
    txtNozzleAngle_LostFocus 'range checking
  Else
    With cboNozzleAngle
      .ListIndex = 0
      For i = 0 To .ListCount - 1
        If .List(i) = ModPrefix & SaveNozzleAngle Then
          .ListIndex = i
          Exit For
        End If
      Next
    End With
  End If
End Sub

Private Sub cmdCancel_Click()
  Me.Tag = "False" 'return cancellation status
  Me.Hide
End Sub

Private Sub cmdOk_Click()
'execute DropKirk calcs
  Dim DKstat As Integer
  Dim MaxErrLev As Integer
  Dim squal As Integer
  Dim np As Integer
  ReDim Diam(MAX_DROPS - 1) As Single
  ReDim mfrac(MAX_DROPS - 1) As Single
  'Change the form mouse pointer
  Me.MousePointer = vbHourglass 'hourglass
  'store the new form data
  FormToData
  'do the calcs
  DKstat = CalcDropKirk(BK2, True, MaxErrLev, squal, np, Diam(), mfrac())
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
    BK2.MaxErrorLevel = MaxErrLev
    Me.Tag = "True" 'return success status
    Me.Hide
  End If
End Sub

Private Sub Form_Load()
'Initialize the controls on this form
  Dim i As Integer
  Dim s As String
  
  'center the form
  CenterForm Me

  Me.Tag = "False"  'default return value
  
  'Load Nozzle Name combo
  i = 0: s = GetARSNozName(i)
  While s <> ""
    cboNozzleName.AddItem s
    i = i + 1: s = GetARSNozName(i)
  Wend
  'Align Orifice value controls
  txtOrifice.Left = cboOrifice.Left
  txtNozzleAngle.Left = cboNozzleAngle.Left
  
  lblAirSpeedUnits.Caption = UnitsName(UN_SPEED)
  lblFlightSpeed.Caption = AGFormat$(UnitsDisplay(UD.AC.TypSpeed, UN_SPEED))
  lblFlightSpeedUnits.Caption = UnitsName(UN_SPEED)
  lblPressureUnits.Caption = UnitsName(UN_PRESSURE)
  
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

Private Sub optSprayType_Click(Index As Integer)
  If Index = 0 Then
    chkSwathDispAdjust.Enabled = True
  Else
    chkSwathDispAdjust.Enabled = False
  End If
End Sub

Private Sub txtNozzleAngle_KeyPress(KeyAscii As Integer)
  If KeyAscii = Asc(vbCr) Then
    txtNozzleAngle_LostFocus 'range checking
    KeyAscii = 0
  End If
End Sub

Private Sub txtNozzleAngle_LostFocus()
'range checking
  Dim nozmin As Single
  Dim nozmax As Single
  Dim nozcur As Single
    
  nozcur = Val(Mid$(txtNozzleAngle.Text, Len(ModPrefix) + 1)) 'strip off prefix
  nozmin = ModVal(0)
  nozmax = ModVal(1)
  If nozcur < nozmin Then txtNozzleAngle.Text = ModPrefix & Format$(nozmin)
  If nozcur > nozmax Then txtNozzleAngle.Text = ModPrefix & Format$(nozmax)
  'place the cursor at the end
  txtNozzleAngle.SelStart = Len(txtNozzleAngle.Text)
  txtNozzleAngle.SelLength = 0
End Sub

Private Sub txtOrifice_KeyPress(KeyAscii As Integer)
  If KeyAscii = Asc(vbCr) Then
    txtOrifice_LostFocus 'range checking
    KeyAscii = 0
  End If
End Sub

Private Sub txtOrifice_LostFocus()
  'Range checking
  Dim ofcmin As Single
  Dim ofcmax As Single
  Dim ofccur As Single
  ofccur = Val(Mid$(txtOrifice.Text, Len(OfcPrefix) + 1)) 'strip off prefix
  ofcmin = OfcVal(0)
  ofcmax = OfcVal(1)
  If ofccur < ofcmin Then txtOrifice.Text = OfcPrefix & Format$(ofcmin)
  If ofccur > ofcmax Then txtOrifice.Text = OfcPrefix & Format$(ofcmax)
  'place the cursor at the end
  txtOrifice.SelStart = Len(txtOrifice.Text)
  txtOrifice.SelLength = 0
End Sub

