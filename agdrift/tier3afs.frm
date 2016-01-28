VERSION 5.00
Begin VB.Form frmTier3afs 
   BorderStyle     =   0  'None
   Caption         =   "Tier III Aerial Forestry Input"
   ClientHeight    =   6795
   ClientLeft      =   1095
   ClientTop       =   2520
   ClientWidth     =   9480
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   HelpContextID   =   1295
   Icon            =   "TIER3AFS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Tag             =   "tier3"
   Begin VB.PictureBox picLogo 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5160
      ScaleHeight     =   735
      ScaleWidth      =   4215
      TabIndex        =   53
      Top             =   6000
      Width           =   4215
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
         TabIndex        =   25
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
      Begin VB.Label lblTier 
         Caption         =   "Tier III Aerial Forestry"
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
         Left            =   2640
         TabIndex        =   26
         Top             =   120
         Width           =   1575
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
         TabIndex        =   24
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame fraRunTitle 
      Caption         =   "Title"
      Height          =   735
      Left            =   120
      TabIndex        =   30
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
   Begin VB.Frame fraAircraft 
      Caption         =   "Aircraft"
      Height          =   2535
      Left            =   120
      TabIndex        =   27
      Top             =   720
      Width           =   3015
      Begin VB.CommandButton cmdEditNozzles 
         Caption         =   "Nozzles and DSD"
         Height          =   495
         HelpContextID   =   1185
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdEditAc 
         Caption         =   "Aircraft"
         Height          =   375
         HelpContextID   =   1016
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtAcFlightLines 
         Height          =   285
         HelpContextID   =   1190
         Left            =   1440
         TabIndex        =   4
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtReleaseHeight 
         Height          =   285
         HelpContextID   =   1060
         Left            =   1440
         TabIndex        =   3
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblACDesc 
         Caption         =   "AC description"
         Height          =   375
         Left            =   1080
         TabIndex        =   51
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblACType 
         Caption         =   "aircraft type"
         Height          =   255
         Left            =   1560
         TabIndex        =   48
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblRelHeightUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   43
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label lblAcFlightLines 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flight Lines:"
         Height          =   195
         Left            =   495
         TabIndex        =   32
         Top             =   2085
         Width           =   840
      End
      Begin VB.Label lblReleaseHeight 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boom Height:"
         Height          =   195
         Left            =   405
         TabIndex        =   29
         Top             =   1710
         Width           =   960
      End
      Begin VB.Label lblAcName 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraSpray 
      Caption         =   "Spray Material"
      Height          =   1215
      Left            =   3240
      TabIndex        =   47
      Top             =   720
      Width           =   3015
      Begin VB.CommandButton cmdEditCarrier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Material"
         Height          =   375
         HelpContextID   =   1257
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblSMDesc 
         Caption         =   "SM description"
         Height          =   495
         Left            =   1080
         TabIndex        =   50
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblSprayMatType 
         Caption         =   "spray mat type"
         Height          =   255
         Left            =   1560
         TabIndex        =   49
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblCarrierType 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   1080
         TabIndex        =   46
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraMet 
      Caption         =   "Meteorology"
      Height          =   1695
      Left            =   3240
      TabIndex        =   31
      Top             =   1920
      Width           =   3015
      Begin VB.TextBox txtMetWindDir 
         Height          =   285
         HelpContextID   =   1328
         Left            =   1455
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtMetHumidity 
         Height          =   285
         HelpContextID   =   1230
         Left            =   1440
         TabIndex        =   14
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtMetTemperature 
         Height          =   285
         HelpContextID   =   1270
         Left            =   1440
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtMetWindSpeed 
         Height          =   285
         HelpContextID   =   1330
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblMetWindDir 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Wind Direction:"
         Height          =   195
         Left            =   255
         TabIndex        =   56
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label lblWindDirUnits 
         AutoSize        =   -1  'True
         Caption         =   "deg"
         Height          =   195
         Left            =   2400
         TabIndex        =   55
         Top             =   645
         Width           =   270
      End
      Begin VB.Label lblHumidityUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2400
         TabIndex        =   36
         Top             =   1380
         Width           =   150
      End
      Begin VB.Label lblTemperatureUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   38
         Top             =   1005
         Width           =   420
      End
      Begin VB.Label lblWindSpeedUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2385
         TabIndex        =   39
         Top             =   285
         Width           =   420
      End
      Begin VB.Label lblMetHumidity 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rel. Humidity:"
         Height          =   195
         Left            =   135
         TabIndex        =   35
         Top             =   1365
         Width           =   1200
      End
      Begin VB.Label lblMetTemperature 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Temperature:"
         Height          =   195
         Left            =   195
         TabIndex        =   34
         Top             =   990
         Width           =   1140
      End
      Begin VB.Label lblMetWindSpeed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Wind Speed:"
         Height          =   195
         Left            =   225
         TabIndex        =   33
         Top             =   285
         Width           =   1110
      End
   End
   Begin VB.Frame fraSwath 
      Caption         =   "Swath"
      Height          =   2775
      Left            =   120
      TabIndex        =   44
      Top             =   3240
      Width           =   3015
      Begin VB.CheckBox chkHalfBoom 
         Caption         =   "Half Boom Effect"
         Height          =   255
         HelpContextID   =   1481
         Left            =   360
         TabIndex        =   9
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox cboSwathWidthType 
         Height          =   315
         HelpContextID   =   1260
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtSwathWidth 
         Height          =   285
         HelpContextID   =   1260
         Left            =   1440
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtSwathDisp 
         Height          =   285
         HelpContextID   =   1080
         Left            =   1440
         TabIndex        =   8
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox cboSwathDispType 
         Height          =   315
         HelpContextID   =   1080
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lblSwathWidth 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   195
         Left            =   885
         TabIndex        =   40
         Top             =   885
         Width           =   450
      End
      Begin VB.Label lblSwathWidthType 
         AutoSize        =   -1  'True
         Caption         =   "Swath Width Definition:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label lblSwathDispType 
         AutoSize        =   -1  'True
         Caption         =   "Swath Displacement Definition:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   2205
      End
      Begin VB.Label lblSwathWidthUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2385
         TabIndex        =   41
         Top             =   885
         Width           =   420
      End
      Begin VB.Label lblSwathDispUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   42
         Top             =   1965
         Width           =   420
      End
      Begin VB.Label lblSwathDisp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   195
         Left            =   885
         TabIndex        =   45
         Top             =   1965
         Width           =   450
      End
   End
   Begin VB.Frame fraTransport 
      Caption         =   "Transport"
      Height          =   1215
      Left            =   6360
      TabIndex        =   68
      Top             =   720
      Width           =   3015
      Begin VB.TextBox txtFluxPlane 
         Height          =   285
         HelpContextID   =   1160
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblFluxPlaneUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   70
         Top             =   300
         Width           =   420
      End
      Begin VB.Label lblFluxPlane 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flux Plane:"
         Height          =   195
         Left            =   360
         TabIndex        =   69
         Top             =   285
         Width           =   975
      End
   End
   Begin VB.Frame fraCanopy 
      Caption         =   "Canopy"
      Height          =   2055
      Left            =   6360
      TabIndex        =   62
      Top             =   1920
      Width           =   3015
      Begin VB.TextBox txtNDRuff 
         Height          =   285
         HelpContextID   =   1485
         Left            =   1440
         TabIndex        =   19
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtNDDisp 
         Height          =   285
         HelpContextID   =   1486
         Left            =   1440
         TabIndex        =   20
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmdEditCanopy 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Canopy"
         Height          =   375
         HelpContextID   =   1487
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblNDDispUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   74
         Top             =   1620
         Width           =   420
      End
      Begin VB.Label lblNDRuffUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   73
         Top             =   1260
         Width           =   420
      End
      Begin VB.Label lblNDRuff 
         Alignment       =   2  'Center
         Caption         =   "Canopy Roughness:"
         Height          =   405
         Left            =   480
         TabIndex        =   72
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblNDDisp 
         Alignment       =   2  'Center
         Caption         =   "Canopy Displacement:"
         Height          =   405
         Left            =   360
         TabIndex        =   71
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label lblCanName 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   1080
         TabIndex        =   65
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblCANType 
         Caption         =   "canopy type"
         Height          =   255
         Left            =   1605
         TabIndex        =   64
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblCANDesc 
         Caption         =   "CAN description"
         Height          =   495
         Left            =   1080
         TabIndex        =   63
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Terrain"
      Height          =   1455
      Left            =   6360
      TabIndex        =   57
      Top             =   3960
      Width           =   3015
      Begin VB.TextBox txtCanopyHeight 
         Height          =   285
         HelpContextID   =   1067
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtUpslope 
         Height          =   285
         HelpContextID   =   1488
         Left            =   1440
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtSideslope 
         Height          =   285
         HelpContextID   =   1489
         Left            =   1455
         TabIndex        =   23
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblCanopyHeight 
         Alignment       =   2  'Center
         Caption         =   "Surface Roughness:"
         Height          =   495
         Left            =   480
         TabIndex        =   67
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblCanopyHeightUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   66
         Top             =   420
         Width           =   420
      End
      Begin VB.Label lblUpslope 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Upslope Angle"
         Height          =   195
         Left            =   300
         TabIndex        =   61
         Top             =   765
         Width           =   1035
      End
      Begin VB.Label lblUpslopeUnits 
         AutoSize        =   -1  'True
         Caption         =   "deg"
         Height          =   195
         Left            =   2385
         TabIndex        =   60
         Top             =   765
         Width           =   270
      End
      Begin VB.Label lblSideslopeUnits 
         AutoSize        =   -1  'True
         Caption         =   "deg"
         Height          =   195
         Left            =   2400
         TabIndex        =   59
         Top             =   1125
         Width           =   270
      End
      Begin VB.Label lblSideslope 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sideslope Angle"
         Height          =   195
         Left            =   210
         TabIndex        =   58
         Top             =   1125
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Atmospheric Stability"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3240
      TabIndex        =   75
      Top             =   3600
      Width           =   3015
      Begin VB.CommandButton cmdStability 
         Caption         =   "Stability"
         Height          =   375
         HelpContextID   =   1548
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblStabDesc 
         Caption         =   "Description"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1080
         TabIndex        =   76
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraAdvanced 
      Caption         =   "Advanced Settings"
      Height          =   855
      Left            =   3240
      TabIndex        =   54
      Top             =   4560
      Width           =   3015
      Begin VB.CommandButton cmdEditAdvanced 
         Caption         =   "Edit"
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmTier3afs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tier3afs.frm,v 1.13 2008/10/22 17:26:06 tom Exp $
'this flag is used to tell some controls not to
'take action on their new values. This is required
'to differentiate between programatic state changes
'and user actions
Dim PropTakeAction As Integer 'if true, execute automatic change-related code
                              'for Property text boxes

Private Sub cboSwathDispType_Click()
  If cboSwathDispType.ListIndex <> UD.CTL.SwathDispType Then
    'Warn if changing to Frac of AR
    If cboSwathDispType.ListIndex = 1 Then
      If Not UP.SuppressTier3Warn Then
        MsgBox "Selection of Fraction of Application Rate " + _
               "will produce some estimated results and " + _
               "will suppress some Flux Plane results.", _
               vbInformation + vbOKOnly
      End If
    End If
    UD.CTL.SwathDispType = cboSwathDispType.ListIndex
    txtSwathDisp_Change  'Updates internal value for units
    UpdateControlControls
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub cboSwathWidthType_Click()
  If PropTakeAction Then
    If cboSwathWidthType.ListIndex <> UD.CTL.SwathWidthType Then
      UD.CTL.SwathWidthType = cboSwathWidthType.ListIndex
      UpdateControlControls  'refresh the type control
      txtSwathWidth_Change   'refresh the value control
      UpdateDataChangedFlag True 'Data was changed
      UC.Valid = False 'Calcs need to be redone
    End If
  End If
End Sub

Private Sub chkHalfBoom_Click()
  If PropTakeAction Then
    UD.CTL.HalfBoom = chkHalfBoom.Value
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub cmdEditAc_Click()
  Me.MousePointer = vbHourglass
  frmAircraft.Show vbModal
  UpdateACTypeLabel 'Update it, it might have changed
  Me.MousePointer = vbDefault
End Sub

Private Sub cmdEditAdvanced_Click()
  Me.MousePointer = vbHourglass
  frmAdvanced.Show vbModal
  Me.MousePointer = vbDefault
End Sub

Private Sub cmdEditCanopy_Click()
  Me.MousePointer = vbHourglass
  frmCanopy.Show vbModal
  UpdateCanopyControls 'Update, stuff may have changed
  UpdateTerrainControls
  Me.MousePointer = vbDefault
End Sub

Private Sub cmdEditCarrier_Click()
  Me.MousePointer = vbHourglass
  EditSprayMaterial
  UpdateSMTypeLabel 'Update it, it might have changed
  Me.MousePointer = vbDefault
End Sub

Private Sub cmdEditNozzles_Click()
  Me.MousePointer = vbHourglass
  frmNozzles.Show vbModal
  DataToForm 'stuff may have changed
  Me.MousePointer = vbDefault
End Sub

Private Sub DataToForm()
'Places user data values in form controls
  Dim PTAtemp As Integer

  'Turn off property control actions
  PTAtemp = PropTakeAction  'save current value
  PropTakeAction = False

  'File name
  Me.Caption = FormCaption
  txtRunTitle.Text = UD.Title               'Title
  ' Drop size
  'Spray Material
  UpdateSMTypeLabel
  'Meteorology
  txtMetWindSpeed.Text = AGFormat$(UnitsDisplay(UD.MET.WS, UN_SPEED))
  txtMetWindDir.Text = AGFormat$(UD.MET.WD)
  lblWindSpeedUnits.Caption = UnitsName(UN_SPEED)
  txtMetTemperature.Text = AGFormat$(UnitsDisplay(UD.MET.temp, UN_TEMP))
  lblTemperatureUnits.Caption = UnitsName(UN_TEMP)
  txtMetHumidity.Text = AGFormat$(UD.MET.Humidity)
  UpdateStabTypeLabel
  'Aircraft
  txtReleaseHeight.Text = AGFormat$(UnitsDisplay(UD.CTL.Height, UN_LENGTH))  'Altitude
  lblRelHeightUnits.Caption = UnitsName(UN_LENGTH)
  txtAcFlightLines = Format$(UD.CTL.NumLines)   'No. flight lines
  UpdateACTypeLabel
  'Swath
  cboSwathWidthType.ListIndex = UD.CTL.SwathWidthType                   'Swath Width Type
  'Swath width units depend on Swath Width type
  If UD.CTL.SwathWidthType = 0 Then  'fixed width
    utype = UN_LENGTH
  Else                               '1.2 Wing, factor*Wing
    utype = UN_NONE
  End If
  txtSwathWidth = AGFormat$(UnitsDisplay(UD.CTL.SwathWidth, utype)) 'Swath Width
  lblSwathWidthUnits.Caption = UnitsName(utype)
  
  cboSwathDispType.ListIndex = UD.CTL.SwathDispType 'Swath Disp Type
  'Swath displacement units depend on Swath Displacement type
  If UD.CTL.SwathDispType = 2 Then      'fixed distance
    utype = UN_LENGTH
  Else
    utype = UN_NONE
  End If
  txtSwathDisp = AGFormat$(UnitsDisplay(UD.CTL.SwathDisp, utype)) 'Swath Displacement
  lblSwathDispUnits.Caption = UnitsName(utype)
  
  'Transport
  txtFluxPlane.Text = AGFormat$(UnitsDisplay(UD.CTL.FluxPlane, UN_LENGTH)) 'Flux Plane
  lblFluxPlaneUnits = UnitsName(UN_LENGTH)

  'Canopy
  UpdateCanopyControls
  
  'Terrain
  UpdateTerrainControls
  txtUpslope.Text = AGFormat$(UD.TRN.Upslope)
  txtSideslope.Text = AGFormat$(UD.TRN.Sideslope)
  
  UpdateControlControls 'adjust the Displacement controls

  'Restore the property action state
  PropTakeAction = PTAtemp
End Sub

Private Sub cmdStability_Click()
  frmStability.Show vbModal
  UpdateStabTypeLabel
  'Warning message for other than default
  If UD.MET.Insolation <> 4 And Not UP.SuppressTier3Warn Then
    MsgBox "The engineering approach implemented in this " & _
           "feature is based on previous published work " & _
           "but has not been validated in this context.", _
           vbInformation
  End If
End Sub

Private Sub UpdateStabTypeLabel()
'update the state of the Stability Type label
  lblStabDesc.Caption = GetTypeNameStability(UD.MET.Insolation)
End Sub

Private Sub Form_Load()
'Initialize the data entry form
  
  'Init Swath Width Combo box
  cboSwathWidthType.AddItem "Fixed Width"
  cboSwathWidthType.AddItem "1.2 x Wingspan"
  cboSwathWidthType.AddItem "Factor x Wingspan"

  'Init Swath Displacement Type box
  cboSwathDispType.AddItem "Fraction of Swath Width"
  cboSwathDispType.AddItem "Fraction of Application Rate"
  cboSwathDispType.AddItem "Fixed Distance"
  cboSwathDispType.AddItem "Aircraft Centerline"

  'Allow form controls to take effect
  PropTakeAction = True

  'Transfer User data to form controls
  DataToForm
End Sub

Private Sub Form_Resize()
'relocate controls when the form is resized
  Const MRGN = 120
  'position agdrift logo
  'the top must not go above the advanced frame
  'the left must not go past the control frame
  toplimit = fraAdvanced.Top + fraAdvanced.Height + MRGN
  leftlimit = fraSwath.Left + fraSwath.Width + MRGN
  logotop = Me.ScaleHeight - picLogo.Height - MRGN
  logoleft = Me.Width - picLogo.Width - MRGN
  If logotop < toplimit Then logotop = toplimit
  If logoleft < leftlimit Then logoleft = leftlimit
  picLogo.Top = logotop
  picLogo.Left = logoleft
  
  'position the title frame and text box
  'it must not get narrower than the Spray Material frame
  widlimit = fraSpray.Left + fraSpray.Width - fraRunTitle.Left
  titlewidth = Me.ScaleWidth - fraRunTitle.Left - MRGN
  If titlewidth < widlimit Then titlewidth = widlimit
  fraRunTitle.Width = titlewidth
  'text box
  txtRunTitle.Width = fraRunTitle.Width - txtRunTitle.Left - MRGN
End Sub

Private Sub txtAcFlightLines_Change()
  If PropTakeAction Then
    UD.CTL.NumLines = Val(txtAcFlightLines.Text)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtCanopyHeight_Change()
  If PropTakeAction Then
    UD.MET.SurfRough = UnitsInternal(Val(txtCanopyHeight.Text), UN_LENGTH)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtFluxPlane_Change()
  If PropTakeAction Then
    UD.CTL.FluxPlane = UnitsInternal(Val(txtFluxPlane.Text), UN_LENGTH)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtMetHumidity_Change()
  If PropTakeAction Then
    UD.MET.Humidity = Val(txtMetHumidity.Text)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtMetTemperature_Change()
  If PropTakeAction Then
    UD.MET.temp = UnitsInternal(Val(txtMetTemperature.Text), UN_TEMP)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtMetWindDir_Change()
  If PropTakeAction Then
    UD.MET.WD = Val(txtMetWindDir.Text)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtMetWindSpeed_Change()
  If PropTakeAction Then
    UD.MET.WS = UnitsInternal(Val(txtMetWindSpeed.Text), UN_SPEED)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtNDDisp_Change()
  If PropTakeAction Then
    If UD.CAN.Height > 0 Then
      UD.CAN.NDDisp = UnitsInternal(Val(txtNDDisp.Text), UN_LENGTH) / UD.CAN.Height
    Else
      UD.CAN.NDDisp = UnitsInternal(Val(txtNDDisp.Text), UN_LENGTH)
    End If
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtNDRuff_Change()
  If PropTakeAction Then
    If UD.CAN.Height > 0 Then
      UD.CAN.NDRuff = UnitsInternal(Val(txtNDRuff.Text), UN_LENGTH) / UD.CAN.Height
    Else
      UD.CAN.NDRuff = UnitsInternal(Val(txtNDRuff.Text), UN_LENGTH)
    End If
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtReleaseHeight_Change()
  If PropTakeAction Then
    UD.CTL.Height = UnitsInternal(Val(txtReleaseHeight.Text), UN_LENGTH)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtRunTitle_Change()
  If PropTakeAction Then
    UD.Title = txtRunTitle.Text
    UpdateDataChangedFlag True 'Data was changed
  End If
End Sub

Private Sub txtSideslope_Change()
  If PropTakeAction Then
    UD.TRN.Sideslope = Val(txtSideslope.Text)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtSwathDisp_Change()
  If PropTakeAction Then
    'units depend on SwathDispType
    If UD.CTL.SwathDispType = 2 Then      'fixed
      utype = UN_LENGTH
    Else
      utype = UN_NONE
    End If
    UD.CTL.SwathDisp = UnitsInternal(Val(txtSwathDisp.Text), utype)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtSwathWidth_Change()
  If PropTakeAction Then
    If UD.CTL.SwathWidthType = 0 Then  'fixed width
      utype = UN_LENGTH
    Else                               '1.2 Wing, factor*Wing
      utype = UN_NONE
    End If
    UD.CTL.SwathWidth = UnitsInternal(Val(txtSwathWidth.Text), utype)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub UpdateACTypeLabel()
'update the state of the Aircraft Type and
'description labels

  Select Case UD.AC.Type
  Case 0  'basic
    lblACType.Caption = "Basic"
  Case 1  'user-defined
    lblACType.Caption = "User-defined"
  Case 2  'library
    lblACType.Caption = "Library"
  End Select
  lblACDesc.Caption = "(" + Trim$(UD.AC.Name) + ")"
End Sub

Private Sub UpdateCanopyControls()
'update the display of Canopy-related controls.
'Roughness and Displacement are tied to Height

  Dim PTAsave As Boolean
  
  PTAsave = PropTakeAction
  PropTakeAction = False
  
  'canopy type label
  Select Case UD.CAN.Type
  Case 0  'None
    lblCANType.Caption = "None"
    lblCANDesc.Caption = ""
  Case 1  'story
    lblCANType.Caption = "Story, " + Format$(UD.CAN.NumEnv) + " levels"
    lblCANDesc.Caption = "(" + Trim$(UD.CAN.Name) + ")"
  Case 2  'optical
    Select Case UD.CAN.optType
    Case 1 'user-def
      lblCANType.Caption = "Optical, user-def"
    Case 2 'library
      lblCANType.Caption = "Optical, library"
    End Select
    lblCANDesc.Caption = "(" + Trim$(UD.CAN.Name) + ")"
  Case 3 'Basic
    lblCANType.Caption = "Basic"
    lblCANDesc.Caption = "(" + Trim$(UD.CAN.Name) + ")"
  Case Else
    lblCANType.Caption = ""
    lblCANDesc.Caption = ""
  End Select
  
  If UD.CAN.Type = 0 Then
    lblNDRuff.Enabled = False
    txtNDRuff.Enabled = False
    lblNDRuffUnits.Enabled = False
    txtNDRuff.Text = ""
    lblNDRuffUnits = UnitsName(UN_LENGTH)
    
    lblNDDisp.Enabled = False
    txtNDDisp.Enabled = False
    lblNDDispUnits.Enabled = False
    txtNDDisp.Text = ""
    lblNDDispUnits = UnitsName(UN_LENGTH)
  Else
    lblNDRuff.Enabled = True
    txtNDRuff.Enabled = True
    lblNDRuffUnits.Enabled = True
    txtNDRuff.Text = AGFormat$(UnitsDisplay(UD.CAN.NDRuff * UD.CAN.Height, UN_LENGTH))
    lblNDRuffUnits = UnitsName(UN_LENGTH)
    
    lblNDDisp.Enabled = True
    txtNDDisp.Enabled = True
    lblNDDispUnits.Enabled = True
    txtNDDisp.Text = AGFormat$(UnitsDisplay(UD.CAN.NDDisp * UD.CAN.Height, UN_LENGTH))
    lblNDDispUnits = UnitsName(UN_LENGTH)
  End If
  
  PropTakeAction = PTAsave
End Sub

Private Sub UpdateControlControls()
'Update the Control controls to match the current settings
' - Swath Displacement input is invisible for
'      1/2 Swath, 1 swath
' - Swath Displacement input units change for
'      % max, fixed value
' - Flux plane input is not available for
'      % max
' - Canopy Height input is not available for
'      % max
'
  'Swath Width
  Select Case UD.CTL.SwathWidthType
  Case 0  'Fixed value
    lblSwathWidth.Caption = "Swath Width:"
    txtSwathWidth.Visible = True
    lblSwathWidthUnits.Caption = UnitsName(UN_LENGTH)
  Case 1  '1.2 x Wingspan
    lblSwathWidth.Caption = ""
    txtSwathWidth.Visible = False
    txtSwathWidth.Text = "1.2"
    lblSwathWidthUnits.Caption = UnitsName(UN_NONE)
  Case 2  'Factor x WingSpan
    lblSwathWidth.Caption = "Factor:"
    txtSwathWidth.Visible = True
    lblSwathWidthUnits.Caption = UnitsName(UN_NONE)
  End Select
  
  'Swath Displacement
  Select Case UD.CTL.SwathDispType
  Case 0 'fraction of swath width
    txtSwathDisp.Visible = True
    lblSwathDispUnits.Visible = False
    lblSwathDisp.Visible = True
    lblSwathDisp.Caption = "Fraction:"
    lblFluxPlane.Enabled = True
    txtFluxPlane.Enabled = True
    lblFluxPlaneUnits.Enabled = True
  Case 1 'fraction of applied
    txtSwathDisp.Visible = True
    lblSwathDispUnits.Visible = False
    lblSwathDisp.Visible = True
    lblSwathDisp.Caption = "Fraction:"
    lblFluxPlane.Enabled = False
    txtFluxPlane.Enabled = False
    lblFluxPlaneUnits.Enabled = False
  Case 2  'Fixed Distance
    txtSwathDisp.Visible = True
    lblSwathDispUnits.Visible = True
    lblSwathDispUnits.Caption = UnitsName(UN_LENGTH)
    lblSwathDisp.Visible = True
    lblSwathDisp.Caption = "Distance:"
    lblFluxPlane.Enabled = True
    txtFluxPlane.Enabled = True
    lblFluxPlaneUnits.Enabled = True
  Case 3 'Aircraft Centerline
    txtSwathDisp.Visible = False
    lblSwathDispUnits.Visible = False
    lblSwathDisp.Visible = False
    lblFluxPlane.Enabled = True
    txtFluxPlane.Enabled = True
    lblFluxPlaneUnits.Enabled = True
  End Select
End Sub

Private Sub UpdateTerrainControls()
'Update the state of some terrain-related controls
  Dim PTAsave As Boolean
  PTAsave = PropTakeAction
  PropTakeAction = False
'  If UD.CAN.Type = 0 Then
'    lblCanopyHeight.Enabled = True
'    txtCanopyHeight.Enabled = True
'    txtCanopyHeight.Text = AGFormat$(UnitsDisplay(UD.MET.SurfRough, UN_LENGTH))
'    lblCanopyHeightUnits.Enabled = True
'  Else
'    lblCanopyHeight.Enabled = False
'    txtCanopyHeight.Enabled = False
'    txtCanopyHeight.Text = ""
'    lblCanopyHeightUnits.Enabled = False
'  End If
  txtCanopyHeight.Text = AGFormat$(UnitsDisplay(UD.MET.SurfRough, UN_LENGTH))
  lblCanopyHeightUnits = UnitsName(UN_LENGTH)
  PropTakeAction = PTAsave
End Sub

Private Sub UpdateSMTypeLabel()
'update the state of the Spray Material Type and
'description labels

  lblSMDesc.Caption = "(" + Trim$(UD.SM.Name) + ")"
  Select Case UD.SM.Type
  Case 0  'basic
    lblSprayMatType.Caption = "Basic"
  Case 1  'user-defined
    lblSprayMatType.Caption = "User-defined"
  Case 2  'library
    lblSprayMatType.Caption = "Library"
  End Select
End Sub

Private Sub txtUpslope_Change()
  If PropTakeAction Then
    UD.TRN.Upslope = Val(txtUpslope.Text)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub
