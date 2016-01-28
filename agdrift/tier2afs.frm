VERSION 5.00
Begin VB.Form frmTier2afs 
   BorderStyle     =   0  'None
   Caption         =   "Tier II Aerial Forestry Input"
   ClientHeight    =   6810
   ClientLeft      =   2295
   ClientTop       =   1500
   ClientWidth     =   9480
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   HelpContextID   =   1290
   Icon            =   "TIER2AFS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6810
   ScaleWidth      =   9480
   Tag             =   "tier2"
   Begin VB.PictureBox picLogo 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   5280
      ScaleHeight     =   735
      ScaleWidth      =   4095
      TabIndex        =   38
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
         TabIndex        =   40
         Top             =   0
         Width           =   195
      End
      Begin VB.Label lblTier 
         Caption         =   "Tier II Aerial Forestry"
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
         TabIndex        =   39
         Top             =   120
         Width           =   1470
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
         TabIndex        =   22
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame fraRunTitle 
      Caption         =   "Title"
      Height          =   735
      Left            =   120
      TabIndex        =   20
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
   Begin VB.Frame fraDropSize 
      Caption         =   "Drop Size Distribution"
      Height          =   1215
      Left            =   3240
      TabIndex        =   23
      Top             =   720
      Width           =   3015
      Begin VB.CommandButton cmdEditDrop 
         Caption         =   "DSD"
         Height          =   375
         HelpContextID   =   1100
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblDSDdesc 
         Caption         =   "DSD description"
         Height          =   570
         Left            =   1080
         TabIndex        =   37
         Top             =   525
         Width           =   1815
      End
      Begin VB.Label lblDropDistType 
         Caption         =   "drop dist type"
         Height          =   255
         Left            =   1575
         TabIndex        =   36
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   1065
         TabIndex        =   35
         Top             =   285
         Width           =   495
      End
   End
   Begin VB.Frame fraAircraft 
      Caption         =   "Aircraft"
      Height          =   1935
      Left            =   120
      TabIndex        =   41
      Top             =   720
      Width           =   3015
      Begin VB.ComboBox cboAircraft 
         Height          =   315
         HelpContextID   =   1023
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtReleaseHeight 
         Height          =   285
         HelpContextID   =   1060
         Left            =   1560
         TabIndex        =   3
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtAcFlightLines 
         Height          =   285
         HelpContextID   =   1190
         Left            =   1560
         TabIndex        =   4
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtBoomWidth 
         Height          =   285
         HelpContextID   =   1061
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblReleaseHeight 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boom Height:"
         Height          =   195
         Left            =   525
         TabIndex        =   48
         Top             =   1230
         Width           =   960
      End
      Begin VB.Label lblAcFlightLines 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flight Lines:"
         Height          =   195
         Left            =   615
         TabIndex        =   47
         Top             =   1605
         Width           =   840
      End
      Begin VB.Label lblRelHeightUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2505
         TabIndex        =   46
         Top             =   1245
         Width           =   420
      End
      Begin VB.Label lblACType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   150
         TabIndex        =   45
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblACName 
         AutoSize        =   -1  'True
         Caption         =   "AC name"
         Height          =   195
         Left            =   720
         TabIndex        =   44
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblBoomWidthUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2505
         TabIndex        =   43
         Top             =   885
         Width           =   120
      End
      Begin VB.Label lblBoomWidth 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boom Length:"
         Height          =   195
         Left            =   495
         TabIndex        =   42
         Top             =   870
         Width           =   990
      End
   End
   Begin VB.Frame fraSwath 
      Caption         =   "Swath"
      Height          =   2775
      Left            =   120
      TabIndex        =   49
      Top             =   2640
      Width           =   3015
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
         TabIndex        =   55
         Top             =   885
         Width           =   450
      End
      Begin VB.Label lblSwathWidthType 
         AutoSize        =   -1  'True
         Caption         =   "Swath Width Definition:"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label lblSwathDispType 
         AutoSize        =   -1  'True
         Caption         =   "Swath Displacement Definition:"
         Height          =   195
         Left            =   60
         TabIndex        =   53
         Top             =   1320
         Width           =   2205
      End
      Begin VB.Label lblSwathWidthUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2385
         TabIndex        =   52
         Top             =   885
         Width           =   420
      End
      Begin VB.Label lblSwathDispUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   51
         Top             =   1965
         Width           =   420
      End
      Begin VB.Label lblSwathDisp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   195
         Left            =   885
         TabIndex        =   50
         Top             =   1965
         Width           =   450
      End
   End
   Begin VB.Frame fraSpray 
      Caption         =   "Spray Material"
      Height          =   1935
      Left            =   3240
      TabIndex        =   32
      Top             =   1920
      Width           =   3015
      Begin VB.ComboBox cboCarrierType 
         Height          =   315
         HelpContextID   =   1070
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtFlowRate 
         Height          =   285
         HelpContextID   =   1150
         Left            =   1440
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtNonvol 
         Height          =   285
         HelpContextID   =   1392
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCarrierType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carrier Type:"
         Height          =   195
         Left            =   225
         TabIndex        =   30
         Top             =   1140
         Width           =   1110
      End
      Begin VB.Label lblFlowRate 
         Alignment       =   2  'Center
         Caption         =   "Spray Volume Rate:"
         Height          =   450
         Left            =   240
         TabIndex        =   31
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label lblFlowRateUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   34
         Top             =   780
         Width           =   420
      End
      Begin VB.Label lblNonvol 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nonvol. Fraction:"
         Height          =   195
         Left            =   105
         TabIndex        =   33
         Top             =   405
         Width           =   1215
      End
   End
   Begin VB.Frame fraMet 
      Caption         =   "Meteorology"
      Height          =   1575
      Left            =   3240
      TabIndex        =   21
      Top             =   3840
      Width           =   3015
      Begin VB.TextBox txtMetHumidity 
         Height          =   285
         HelpContextID   =   1230
         Left            =   1440
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtMetTemperature 
         Height          =   285
         HelpContextID   =   1270
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtMetWindSpeed 
         Height          =   285
         HelpContextID   =   1330
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblHumidityUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2400
         TabIndex        =   27
         Top             =   1140
         Width           =   150
      End
      Begin VB.Label lblTemperatureUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   28
         Top             =   765
         Width           =   420
      End
      Begin VB.Label lblWindSpeedUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2385
         TabIndex        =   29
         Top             =   405
         Width           =   420
      End
      Begin VB.Label lblMetHumidity 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rel. Humidity:"
         Height          =   195
         Left            =   135
         TabIndex        =   26
         Top             =   1125
         Width           =   1200
      End
      Begin VB.Label lblMetTemperature 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Temperature:"
         Height          =   195
         Left            =   195
         TabIndex        =   25
         Top             =   750
         Width           =   1140
      End
      Begin VB.Label lblMetWindSpeed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Wind Speed:"
         Height          =   195
         Left            =   225
         TabIndex        =   24
         Top             =   405
         Width           =   1110
      End
   End
   Begin VB.Frame fraTransport 
      Caption         =   "Transport"
      Height          =   1215
      Left            =   6360
      TabIndex        =   59
      Top             =   720
      Width           =   3015
      Begin VB.TextBox txtFluxPlane 
         Height          =   285
         HelpContextID   =   1160
         Left            =   1440
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblFluxPlane 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flux Plane:"
         Height          =   195
         Left            =   360
         TabIndex        =   61
         Top             =   405
         Width           =   975
      End
      Begin VB.Label lblFluxPlaneUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   60
         Top             =   420
         Width           =   420
      End
   End
   Begin VB.Frame fraCanopy 
      Caption         =   "Canopy"
      Height          =   1935
      Left            =   6360
      TabIndex        =   56
      Top             =   1920
      Width           =   3015
      Begin VB.TextBox txtNDDisp 
         Height          =   285
         HelpContextID   =   1486
         Left            =   1440
         TabIndex        =   19
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtNDRuff 
         Height          =   285
         HelpContextID   =   1485
         Left            =   1440
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtCanHgt 
         Height          =   285
         HelpContextID   =   1484
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblNDDisp 
         Alignment       =   2  'Center
         Caption         =   "Canopy Displacement:"
         Height          =   405
         Left            =   360
         TabIndex        =   65
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblNDRuff 
         Alignment       =   2  'Center
         Caption         =   "Canopy Roughness:"
         Height          =   405
         Left            =   480
         TabIndex        =   64
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblNDRuffUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   63
         Top             =   900
         Width           =   420
      End
      Begin VB.Label lblNDDispUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   62
         Top             =   1380
         Width           =   420
      End
      Begin VB.Label lblCanHgtUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   58
         Top             =   405
         Width           =   420
      End
      Begin VB.Label lblCanHgt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Canopy Height:"
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Top             =   390
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmTier2afs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tier2afs.frm,v 1.9 2008/10/22 17:26:06 tom Exp $
'this flag is used to tell some controls not to
'take action on their new values. This is required
'to differentiate between programatic state changes
'and user actions
Dim PropTakeAction As Integer 'if true, execute automatic change-related code
                              'for Property text boxes

Private Sub cboAircraft_Click()
  If PropTakeAction Then
    If cboAircraft.ListIndex <> UD.AC.BasicType Then
      GetBasicDataAC CInt(cboAircraft.ListIndex), UD.AC
      GetBasicDataNZ CInt(cboAircraft.ListIndex), UD.NZ
      lblACName.Caption = "(" & GetBasicNameAC2(UD.AC.BasicType) & ")"
      PropTakeAction = False
      txtBoomWidth.Text = AGFormat$(UD.NZ.BoomWidth)
      PropTakeAction = True
      UpdateDataChangedFlag True 'Data was changed
      UC.Valid = False 'Calcs need to be redone
    End If
  End If
End Sub

Private Sub cboCarrierType_Click()
  If PropTakeAction Then
    If cboCarrierType.ListIndex <> UD.SM.BasicType Then
      UD.SM.BasicType = cboCarrierType.ListIndex
      GetBasicDataSM UD.SM.BasicType, UD.SM
      UpdateSprayMaterialControls
      UpdateDataChangedFlag True 'Data was changed
      UC.Valid = False 'Calcs need to be redone
    End If
  End If
End Sub

Private Sub cboSwathDispType_Click()
  If cboSwathDispType.ListIndex <> UD.CTL.SwathDispType Then
    'Warn if changing to Frac of AR
    If cboSwathDispType.ListIndex = 1 Then
      MsgBox "Selection of Fraction of Application Rate " + _
             "will produce some estimated results and " + _
             "will suppress some Flux Plane results.", _
             vbInformation + vbOKOnly
    End If
    UD.CTL.SwathDispType = cboSwathDispType.ListIndex
    txtSwathDisp_Change 'Updates internal value for units
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

Private Sub cmdEditDrop_Click()
  Me.MousePointer = vbHourglass
  Load frmDropDist
  frmDropDist.lblDSDselection = 0 'Send the DSD index to the form
  frmDropDist.Show vbModal
  DataToForm 'dropdist name and swath stuff may have changed
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
  'Title
  txtRunTitle.Text = UD.Title               'Title
  'Drop size
  UpdateTypeLabel
  'Spray Material
  UpdateSprayMaterialControls
  'Meteorology
  txtMetWindSpeed.Text = AGFormat$(UnitsDisplay(UD.MET.WS, UN_SPEED))
  lblWindSpeedUnits.Caption = UnitsName(UN_SPEED)
  txtMetTemperature.Text = AGFormat$(UnitsDisplay(UD.MET.temp, UN_TEMP))
  lblTemperatureUnits.Caption = UnitsName(UN_TEMP)
  txtMetHumidity.Text = AGFormat$(UD.MET.Humidity)
  'Aircraft
  cboAircraft.ListIndex = UD.AC.BasicType   'Type
  lblACName.Caption = "(" & GetBasicNameAC2(UD.AC.BasicType) & ")"
  txtBoomWidth.Text = AGFormat$(UD.NZ.BoomWidth)
  txtReleaseHeight.Text = AGFormat$(UnitsDisplay(UD.CTL.Height, UN_LENGTH))  'Altitude
  lblRelHeightUnits.Caption = UnitsName(UN_LENGTH)
  txtAcFlightLines = Format$(UD.CTL.NumLines)   'No. flight lines
  'Swath
  cboSwathWidthType.ListIndex = UD.CTL.SwathWidthType       'Swath Width Type
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

'tbc
  lblTportHeightMinUnits = UnitsName(UN_LENGTH)
  lblTportHeightMaxUnits = UnitsName(UN_LENGTH)

  UpdateControlControls 'adjust the Displacement controls

  'Canopy
  UpdateCanopyControls
  
  'Restore the property action state
  PropTakeAction = PTAtemp
End Sub

Private Sub Form_Load()
'Initialize the data entry form
  
  'init Carrier Type list box
  For i = 0 To 1
    cboCarrierType.AddItem GetBasicNameSM(i)
  Next

  'Init Aircraft list box
  For i = 0 To 3
    cboAircraft.AddItem GetBasicNameAC(i)
  Next
  
  'Init Swath Width Combo box
  cboSwathWidthType.AddItem "Fixed Width"
  cboSwathWidthType.AddItem "1.2 x Wingspan"
  'cboSwathWidthType.AddItem "Factor x Wingspan"  'not allowed in Tier II

  'Init Swath Displacement Type box
  cboSwathDispType.AddItem "Fraction of Swath Width"
  cboSwathDispType.AddItem "Fraction of Application Rate"
  cboSwathDispType.AddItem "Fixed Distance"
  cboSwathDispType.AddItem "Aircraft Centerline"

  'allow control changes to take action
  '(see declarations section)
  PropTakeAction = True
  
  'Transfer User data to form controls
  DataToForm
End Sub

Private Sub Form_Resize()
'relocate controls when the form is resized
  'position agdrift logo
  'the top must not go above the met frame
  'the left must not go past the margin
  Const MRGN = 300
  toplimit = fraMet.Top + fraMet.Height + MRGN
  leftlimit = MRGN
  logotop = Me.Height - picLogo.Height - MRGN
  logoleft = Me.Width - picLogo.Width - MRGN
  If logotop < toplimit Then logotop = toplimit
  If logoleft < leftlimit Then logoleft = leftlimit
  picLogo.Top = logotop
  picLogo.Left = logoleft
  
  'position the title frame and text box
  'it must not get narrower than the Aircraft frame
  widlimit = fraAircraft.Left + fraAircraft.Width - fraRunTitle.Left
  titlewidth = Me.Width - fraRunTitle.Left - 100
  If titlewidth < widlimit Then titlewidth = widlimit
  fraRunTitle.Width = titlewidth
  'text box
  txtRunTitle.Width = fraRunTitle.Width - txtRunTitle.Left - 120
End Sub

Private Sub txtAcFlightLines_Change()
  If PropTakeAction Then
    UD.CTL.NumLines = Val(txtAcFlightLines.Text)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtBoomWidth_Change()
  Dim BoomWidth As Single
  If PropTakeAction Then
    If txtBoomWidth.Text <> "" Then
      BoomWidth = Val(txtBoomWidth.Text)
      If BoomWidth > 1 Then 'tbc - get real limits for agnozl
        GetBasicDataNZ UD.AC.BasicType, UD.NZ
        AdjustBasicNozzles BoomWidth, UD.AC.SemiSpan, UD.NZ
        UD.NZ.BoomWidth = BoomWidth
        UpdateDataChangedFlag True 'Data was changed
        UC.Valid = False 'Calcs need to be redone
      Else
        Beep 'warn the user of the bad value
      End If
    End If
  End If
End Sub

Private Sub txtFlowRate_Change()
  If PropTakeAction Then
    UD.SM.FlowRate = UnitsInternal(Val(txtFlowRate.Text), UN_RATEVOL)
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

Private Sub txtMetWindSpeed_Change()
  If PropTakeAction Then
    UD.MET.WS = UnitsInternal(Val(txtMetWindSpeed.Text), UN_SPEED)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtCanHgt_Change()
  If PropTakeAction Then
    UD.CAN.Height = UnitsInternal(Val(txtCanHgt.Text), UN_LENGTH)
    If UD.CAN.Height > 0 Then
      UD.CAN.Type = 3 'basic
    Else
      UD.CAN.Type = 0 'none
    End If
    UpdateCanopyControls True 'The argument prevents txtCanHgt from update
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

Private Sub txtNonvol_Change()
  If PropTakeAction Then
    UD.SM.NVFrac = Val(txtNonvol.Text)
    UD.SM.ACFrac = UD.SM.NVFrac 'Active=nonvol
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

Private Sub UpdateControlControls()
'Update the Control controls to match the current settings
' - Swath Displacement input is invisible for
'      1/2 Swath, 1 swath
' - Swath Displacement input units change for
'      % max, fixed value
' - Flux plane input is not available for
'      % max
'
  'Swath Width
  Select Case UD.CTL.SwathWidthType
  Case 0 'Fixed value
    lblSwathWidth.Caption = "Swath Width:"
    txtSwathWidth.Visible = True
    lblSwathWidthUnits.Caption = UnitsName(UN_LENGTH)
  Case 1 '1.2 x Wingspan
    lblSwathWidth.Caption = ""
    txtSwathWidth.Visible = False
    txtSwathWidth.Text = "1.2"
    lblSwathWidthUnits.Caption = UnitsName(UN_NONE)
  Case 2 'Factor x WingSpan
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
  Case 2 'Fixed Distance
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

Private Sub UpdateSprayMaterialControls()
'transfer stored data to form controls for Spray Material
  temp = PropTakeAction                        'save flag value
  PropTakeAction = False                       'allow raw field modification

  txtNonvol.Text = AGFormat$(UD.SM.NVFrac)
  txtFlowRate.Text = AGFormat$(UnitsDisplay(UD.SM.FlowRate, UN_RATEVOL))
  lblFlowRateUnits.Caption = UnitsName(UN_RATEVOL)
  cboCarrierType.ListIndex = UD.SM.BasicType 'Carrier Type
  
  PropTakeAction = temp                     'restore flag value
End Sub

Private Sub UpdateCanopyControls(Optional ExcludeHeight)
'update the display of Canopy-related controls.
'Roughness and Displacement are tied to Height
'if provided, ExcludeHeight prevents Height from being updated.
'This is useful for the Change events
'of the text boxes that display the values

  Dim PTAsave As Boolean
  
  PTAsave = PropTakeAction
  PropTakeAction = False
  
  If IsMissing(ExcludeHeight) Then
    txtCanHgt.Text = AGFormat$(UnitsDisplay(UD.CAN.Height, UN_LENGTH))
    lblCanHgtUnits.Caption = UnitsName(UN_LENGTH)
  End If
  
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

Private Sub UpdateTypeLabel()
'update the state of the DropDist Type and
'description labels

  lblDSDdesc.Caption = "(" + Trim$(UD.DSD(0).Name) + ")"
  Select Case UD.DSD(0).Type
  Case 0  'basic
    lblDropDistType.Caption = "Basic"
  Case 1  'dropkick
    lblDropDistType.Caption = "DropKick"
  Case 2  'user-defined
    lblDropDistType.Caption = "User-defined"
  Case 3  'library
    lblDropDistType.Caption = "Library"
  End Select
End Sub

