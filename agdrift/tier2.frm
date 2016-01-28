VERSION 4.00
Begin VB.Form frmTier2 
   BorderStyle     =   0  'None
   Caption         =   "Tier II Input"
   ClientHeight    =   6000
   ClientLeft      =   450
   ClientTop       =   1740
   ClientWidth     =   7320
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Height          =   6405
   HelpContextID   =   1290
   Icon            =   "TIER2.frx":0000
   Left            =   390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7320
   Tag             =   "tier2"
   Top             =   1395
   Width           =   7440
   Begin VB.PictureBox picLogo 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3960
      ScaleHeight     =   735
      ScaleWidth      =   3255
      TabIndex        =   53
      Top             =   5160
      Width           =   3255
      Begin VB.Label lblTM 
         AutoSize        =   -1  'True
         Caption         =   "TM"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   -1  'True
            strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   55
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblTier 
         AutoSize        =   -1  'True
         Caption         =   "Tier II"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   12
            underline       =   0   'False
            italic          =   -1  'True
            strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2370
         TabIndex        =   54
         Top             =   240
         Width           =   855
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
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   24
            underline       =   0   'False
            italic          =   -1  'True
            strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame fraSpray 
      Caption         =   "Spray Material"
      Height          =   1815
      Left            =   120
      TabIndex        =   42
      Top             =   2280
      Width           =   3255
      Begin VB.ComboBox cboCarrierType 
         Height          =   315
         HelpContextID   =   1070
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtFlowRate 
         Height          =   285
         HelpContextID   =   1150
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtActiveAmt 
         Height          =   285
         HelpContextID   =   1010
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtNvAmt 
         Height          =   285
         HelpContextID   =   1180
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCarrierType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carrier Type:"
         Height          =   195
         Left            =   225
         TabIndex        =   38
         Top             =   1500
         Width           =   1110
      End
      Begin VB.Label lblFlowRate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Spray Rate:"
         Height          =   195
         Left            =   300
         TabIndex        =   39
         Top             =   1125
         Width           =   1020
      End
      Begin VB.Label lblFlowRateUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   47
         Top             =   1140
         Width           =   420
      End
      Begin VB.Label lblActiveAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Active Rate:"
         Height          =   195
         Left            =   225
         TabIndex        =   46
         Top             =   765
         Width           =   1080
      End
      Begin VB.Label lblActiveAmtUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   45
         Top             =   780
         Width           =   420
      End
      Begin VB.Label lblNvAmt 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nonvol. Rate:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   405
         Width           =   1200
      End
      Begin VB.Label lblNvAmtUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   43
         Top             =   420
         Width           =   420
      End
   End
   Begin VB.Frame fraControl 
      Caption         =   "Control"
      Height          =   2415
      Left            =   3480
      TabIndex        =   36
      Top             =   2760
      Width           =   3735
      Begin VB.ComboBox cboSwathWidthType 
         Height          =   315
         HelpContextID   =   1260
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   2430
      End
      Begin VB.TextBox txtFluxPlane 
         Height          =   285
         HelpContextID   =   1160
         Left            =   2280
         TabIndex        =   16
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txtSwathWidth 
         Height          =   285
         HelpContextID   =   1260
         Left            =   2280
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtSwathDisp 
         Height          =   285
         HelpContextID   =   1080
         Left            =   2280
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox cboSwathDispType 
         Height          =   315
         HelpContextID   =   1080
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblSwathWidth 
         Alignment       =   1  'Right Justify
         Caption         =   "Value:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   645
         Width           =   2055
      End
      Begin VB.Label lblSwathWidthType 
         Alignment       =   2  'Center
         Caption         =   "Swath Width Definition:"
         Height          =   450
         Left            =   60
         TabIndex        =   52
         Top             =   195
         Width           =   1140
      End
      Begin VB.Label lblFluxPlane 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flux Plane:"
         Height          =   195
         Left            =   1200
         TabIndex        =   41
         Top             =   1965
         Width           =   975
      End
      Begin VB.Label lblFluxPlaneUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3240
         TabIndex        =   40
         Top             =   1980
         Width           =   420
      End
      Begin VB.Label lblSwathDispType 
         Alignment       =   2  'Center
         Caption         =   "Swath Displacement Definition:"
         Height          =   615
         Left            =   60
         TabIndex        =   29
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblSwathWidthUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3225
         TabIndex        =   33
         Top             =   645
         Width           =   420
      End
      Begin VB.Label lblSwathDispUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3240
         TabIndex        =   34
         Top             =   1485
         Width           =   420
      End
      Begin VB.Label lblSwathDisp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1485
         Width           =   2055
      End
   End
   Begin VB.Frame fraRunTitle 
      Caption         =   "Title"
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtRunTitle 
         Alignment       =   2  'Center
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   1300
         Left            =   120
         TabIndex        =   0
         Text            =   "Untitled"
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame fraDropSize 
      Caption         =   "Drop Size Distribution"
      Height          =   1455
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   3255
      Begin VB.CommandButton cmdEditDrop 
         Caption         =   "DSD"
         Height          =   375
         HelpContextID   =   1100
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblDSDdesc 
         Caption         =   "DSD description"
         Height          =   570
         Left            =   1080
         TabIndex        =   50
         Top             =   525
         Width           =   2055
      End
      Begin VB.Label lblDropDistType 
         Caption         =   "drop dist type"
         Height          =   255
         Left            =   1575
         TabIndex        =   49
         Top             =   285
         Width           =   1575
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   1065
         TabIndex        =   48
         Top             =   285
         Width           =   495
      End
   End
   Begin VB.Frame fraMet 
      Caption         =   "Meteorology"
      Height          =   1455
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   3255
      Begin VB.TextBox txtMetHumidity 
         Height          =   285
         HelpContextID   =   1230
         Left            =   1440
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtMetTemperature 
         Height          =   285
         HelpContextID   =   1270
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtMetWindSpeed 
         Height          =   285
         HelpContextID   =   1330
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblHumidityUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2400
         TabIndex        =   28
         Top             =   1140
         Width           =   150
      End
      Begin VB.Label lblTemperatureUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2400
         TabIndex        =   30
         Top             =   765
         Width           =   420
      End
      Begin VB.Label lblWindSpeedUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2385
         TabIndex        =   31
         Top             =   405
         Width           =   420
      End
      Begin VB.Label lblMetHumidity 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rel. Humidity:"
         Height          =   195
         Left            =   135
         TabIndex        =   27
         Top             =   1125
         Width           =   1200
      End
      Begin VB.Label lblMetTemperature 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Temperature:"
         Height          =   195
         Left            =   195
         TabIndex        =   26
         Top             =   750
         Width           =   1140
      End
      Begin VB.Label lblMetWindSpeed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Wind Speed:"
         Height          =   195
         Left            =   225
         TabIndex        =   25
         Top             =   405
         Width           =   1110
      End
   End
   Begin VB.Frame fraAircraft 
      Caption         =   "Aircraft"
      Height          =   1935
      Left            =   3480
      TabIndex        =   17
      Top             =   720
      Width           =   3735
      Begin VB.TextBox txtBoomWidth 
         Height          =   285
         HelpContextID   =   1061
         Left            =   2280
         TabIndex        =   56
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtAcFlightLines 
         Height          =   285
         HelpContextID   =   1190
         Left            =   2280
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtReleaseHeight 
         Height          =   285
         HelpContextID   =   1060
         Left            =   2280
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox cboAircraft 
         Height          =   315
         HelpContextID   =   1023
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblBoomWidth 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boom Length:"
         Height          =   195
         Left            =   1215
         TabIndex        =   58
         Top             =   870
         Width           =   990
      End
      Begin VB.Label lblBoomWidthUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3225
         TabIndex        =   57
         Top             =   885
         Width           =   120
      End
      Begin VB.Label lblACName 
         AutoSize        =   -1  'True
         Caption         =   "AC name"
         Height          =   195
         Left            =   720
         TabIndex        =   18
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblACType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   150
         TabIndex        =   51
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblRelHeightUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3225
         TabIndex        =   35
         Top             =   1245
         Width           =   420
      End
      Begin VB.Label lblAcFlightLines 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Number of Flight Lines:"
         Height          =   195
         Left            =   195
         TabIndex        =   24
         Top             =   1605
         Width           =   1980
      End
      Begin VB.Label lblReleaseHeight 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Boom Height:"
         Height          =   195
         Left            =   1050
         TabIndex        =   19
         Top             =   1230
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmTier2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' $Id: tier2.frm,v 1.3 2000/03/06 21:29:24 tom Exp $
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
      UpdateSprayMaterialValues
      UpdateDataChangedFlag True 'Data was changed
      UC.Valid = False 'Calcs need to be redone
    End If
  End If
End Sub

Private Sub cboSwathDispType_Click()
  If cboSwathDispType.ListIndex <> UD.CTL.SwathDispType Then
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
  Me.MousePointer = 11
  frmDropDist.Show 1
  UpdateTypeLabel 'Update it, it might have changed
  Me.MousePointer = 0
End Sub

Private Sub DataToForm()
'Places user data values in form controls
  Dim PTAtemp As Integer

  'Turn off property control actions
  PTAtemp = PropTakeAction  'save current value
  PropTakeAction = False
  
  'File name
  UpdateInputFormCaption
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
  'Control
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
  If UD.CTL.SwathDispType = 3 Then      'fixed distance
    utype = UN_LENGTH
  ElseIf UD.CTL.SwathDispType = 2 Then  '%
    utype = UN_PERCENT
  Else
    utype = UN_NONE
  End If
  txtSwathDisp = AGFormat$(UnitsDisplay(UD.CTL.SwathDisp, utype)) 'Swath Displacement
  lblSwathDispUnits.Caption = UnitsName(utype)
  
  txtFluxPlane.Text = AGFormat$(UnitsDisplay(UD.CTL.FluxPlane, UN_LENGTH)) 'Flux Plane
  lblFluxPlaneUnits = UnitsName(UN_LENGTH)

  UpdateControlControls 'adjust the Displacement controls

  'Restore the property action state
  PropTakeAction = PTAtemp
End Sub

Private Sub Form_Load()
  InitForm  'Initialize the form objects
End Sub

Private Sub Form_Resize()
  ResizeForm
End Sub

Private Sub InitForm()
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
  cboSwathDispType.AddItem "1/2 Swath Width"
  cboSwathDispType.AddItem "1 Swath Width"
  cboSwathDispType.AddItem "Fraction of Application Rate"
  cboSwathDispType.AddItem "Fixed Distance"
  cboSwathDispType.AddItem "Aircraft Centerline"

  'allow control changes to take action
  '(see declarations section)
  PropTakeAction = True
  
  'Transfer User data to form controls
  DataToForm
End Sub

Private Sub ResizeForm()
'relocate controls when the form is resized
  'position agdrift logo
  'the top must not go above the control frame
  'the left must not go past the met frame
  toplimit = fraControl.Top + fraControl.Height + 300
  leftlimit = fraMet.Left + fraMet.Width + 300
  logotop = Me.Height - picLogo.Height - 300
  logoleft = Me.Width - picLogo.Width - 300
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

Private Sub txtActiveAmt_Change()
  If PropTakeAction Then
    UD.SM.ACamt = UnitsInternal(Val(txtActiveAmt.Text), UN_RATEMASS)
    UpdateSprayMaterialValues
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtBoomWidth_Change()
  If PropTakeAction Then
    UD.NZ.BoomWidth = Val(txtBoomWidth.Text)
    UpdateDataChangedFlag True 'Data was changed
    UC.Valid = False 'Calcs need to be redone
  End If
End Sub

Private Sub txtFlowRate_Change()
  If PropTakeAction Then
    UD.SM.FlowRate = UnitsInternal(Val(txtFlowRate.Text), UN_RATEVOL)
    UpdateSprayMaterialValues
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

Private Sub txtNvAmt_Change()
  If PropTakeAction Then
    UD.SM.NVamt = UnitsInternal(Val(txtNvAmt.Text), UN_RATEMASS)
    UpdateSprayMaterialValues
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
    If UD.CTL.SwathDispType = 3 Then      'fixed
      utype = UN_LENGTH
    ElseIf UD.CTL.SwathDispType = 2 Then  '%
      utype = UN_PERCENT
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
  Case 0 '1/2 Swath Width
    txtSwathDisp.Visible = False
    lblSwathDispUnits.Visible = False
    lblSwathDisp.Visible = False
    lblFluxPlane.Enabled = True
    txtFluxPlane.Enabled = True
    lblFluxPlaneUnits.Enabled = True
  Case 1 '1 Swath Width
    txtSwathDisp.Visible = False
    lblSwathDispUnits.Visible = False
    lblSwathDisp.Visible = False
    lblFluxPlane.Enabled = True
    txtFluxPlane.Enabled = True
    lblFluxPlaneUnits.Enabled = True
  Case 2 'fraction of applied
    txtSwathDisp.Visible = True
    lblSwathDispUnits.Visible = False
    lblSwathDisp.Visible = True
    lblSwathDisp.Caption = "Fraction:"
    lblFluxPlane.Enabled = False
    txtFluxPlane.Enabled = False
    lblFluxPlaneUnits.Enabled = False
  Case 3 'Fixed Distance
    txtSwathDisp.Visible = True
    lblSwathDispUnits.Visible = True
    lblSwathDispUnits.Caption = UnitsName(UN_LENGTH)
    lblSwathDisp.Visible = True
    lblSwathDisp.Caption = "Distance:"
    lblFluxPlane.Enabled = True
    txtFluxPlane.Enabled = True
    lblFluxPlaneUnits.Enabled = True
  Case 4 'Aircraft Centerline
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

  txtNvAmt.Text = AGFormat$(UnitsDisplay(UD.SM.NVamt, UN_RATEMASS))
  lblNvAmtUnits.Caption = UnitsName(UN_RATEMASS)
  txtActiveAmt.Text = AGFormat$(UnitsDisplay(UD.SM.ACamt, UN_RATEMASS))
  lblActiveAmtUnits.Caption = UnitsName(UN_RATEMASS)
  txtFlowRate.Text = AGFormat$(UnitsDisplay(UD.SM.FlowRate, UN_RATEVOL))
  lblFlowRateUnits.Caption = UnitsName(UN_RATEVOL)
  cboCarrierType.ListIndex = UD.SM.BasicType 'Carrier Type
  
  PropTakeAction = temp                     'restore flag value
End Sub

Private Sub UpdateSprayMaterialValues()
'Adjust the Spray Material values in response to a change
  Dim flg As Long
  'compute flag for agfrac (first param):
  'icf  0=compute rates
  '     1=compute nonvolatile fraction
  'return val for agfrac:
  'flg  0=no changes to inputs
  '     1=nvamt changed
  '     2=nvamt and actamt changed
  Call agfrac(1, UD.SM.FlowRate, UD.SM.NVfrac, UD.SM.NVamt, UD.SM.ACamt, UD.SM.SpecGrav, flg)
  'transfer to the form controls
  PropSave = PropTakeAction 'Save flag
  PropTakeAction = False    'Turn off automatic actions
  If flg >= 1 Then txtNvAmt.Text = AGFormat$(UnitsDisplay(UD.SM.NVamt, UN_RATEMASS))
  If flg >= 2 Then txtActiveAmt.Text = AGFormat$(UnitsDisplay(UD.SM.ACamt, UN_RATEMASS))
  PropTakeAction = PropSave
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

