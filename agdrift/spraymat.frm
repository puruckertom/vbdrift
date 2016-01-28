VERSION 5.00
Begin VB.Form frmSprayMat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spray Material"
   ClientHeight    =   4665
   ClientLeft      =   1710
   ClientTop       =   2280
   ClientWidth     =   7530
   ForeColor       =   &H80000008&
   Icon            =   "spraymat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4665
   ScaleWidth      =   7530
   Begin VB.Frame fraDropDist 
      Caption         =   "Properties"
      Height          =   3975
      Left            =   2760
      TabIndex        =   11
      Top             =   120
      Width           =   4695
      Begin VB.Frame Frame1 
         Caption         =   "Tank Mix"
         Height          =   1935
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   4455
         Begin VB.CommandButton cmdSprayMatMix 
            Caption         =   "Tank Mix &Calculator"
            Height          =   375
            HelpContextID   =   1549
            Left            =   1560
            TabIndex        =   27
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox txtFlowRate 
            Height          =   285
            HelpContextID   =   1150
            Left            =   2250
            TabIndex        =   20
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtActive 
            Height          =   285
            HelpContextID   =   1010
            Left            =   2250
            TabIndex        =   19
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtNonvol 
            Height          =   285
            HelpContextID   =   1180
            Left            =   2250
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblFlowRateUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Left            =   3210
            TabIndex        =   26
            Top             =   1020
            Width           =   420
         End
         Begin VB.Label lblActive 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Active Rate:"
            Height          =   195
            Left            =   1035
            TabIndex        =   25
            Top             =   645
            Width           =   1080
         End
         Begin VB.Label lblActiveUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Left            =   3210
            TabIndex        =   24
            Top             =   660
            Width           =   420
         End
         Begin VB.Label lblNonvol 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nonvol. Rate:"
            Height          =   195
            Left            =   930
            TabIndex        =   23
            Top             =   285
            Width           =   1200
         End
         Begin VB.Label lblNonvolUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Left            =   3210
            TabIndex        =   22
            Top             =   300
            Width           =   420
         End
         Begin VB.Label lblFlowRate 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Spray Volume Rate:"
            Height          =   195
            Left            =   720
            TabIndex        =   21
            Top             =   960
            Width           =   1425
         End
      End
      Begin VB.TextBox txtSGnonv 
         Height          =   285
         HelpContextID   =   1255
         Left            =   2385
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtName 
         Height          =   285
         HelpContextID   =   1257
         Left            =   720
         TabIndex        =   6
         Text            =   "txtName"
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtSGtank 
         Height          =   285
         HelpContextID   =   1255
         Left            =   2370
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtEvapRate 
         Height          =   285
         HelpContextID   =   1133
         Left            =   2370
         TabIndex        =   9
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblSGnonv 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Specific Gravity (Nonvolatile):"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   1245
         Width           =   2085
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblSGtank 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Specific Gravity (Carrier):"
         Height          =   195
         Left            =   510
         TabIndex        =   12
         Top             =   885
         Width           =   1740
      End
      Begin VB.Label lblEvapRateUnits 
         AutoSize        =   -1  'True
         Caption         =   "µm²/deg C/sec"
         Height          =   195
         Left            =   3315
         TabIndex        =   13
         Top             =   1620
         Width           =   1290
      End
      Begin VB.Label lblEvapRate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Evaporation Rate:"
         Height          =   195
         Left            =   690
         TabIndex        =   14
         Top             =   1605
         Width           =   1560
      End
   End
   Begin VB.Frame fraDistType 
      Caption         =   "Spray Material Type"
      Height          =   3975
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optMatType 
         Caption         =   "&Library"
         Height          =   255
         HelpContextID   =   1257
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   2295
      End
      Begin VB.OptionButton optMatType 
         Caption         =   "&Basic"
         Height          =   255
         HelpContextID   =   1257
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton optMatType 
         Caption         =   "&User-defined"
         Height          =   255
         HelpContextID   =   1257
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cboBasicType 
         Height          =   315
         HelpContextID   =   1257
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1257
      Left            =   6600
      TabIndex        =   1
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1257
      Left            =   5640
      TabIndex        =   0
      Top             =   4200
      Width           =   855
   End
End
Attribute VB_Name = "frmSprayMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: spraymat.frm,v 1.9 2008/10/22 17:26:06 tom Exp $
'this flag is used to tell the option buttons not to
'take action on their new values. This is required
'to differentiate between programatic state changes
'and user actions
Dim OptTakeAction As Integer  'if true, execute automatic change-related code
                              'for Spray Mat option button
Dim PropTakeAction As Integer 'if true, execute automatic change-related code
                              'for Property text boxes

Dim SaveMatType As Integer  'place to save Material type

Private mSM As SprayMaterialData 'Local copy of spray material data

Private mCancelled As Boolean

Private Sub cboBasicType_Click()
'change the Properties controls to reflect the new Basic setting
  Dim xSM As SprayMaterialData
  temp = PropTakeAction   'save flag value
  PropTakeAction = False  'disable change-related functions
  
  txtName.Text = GetBasicNameSM(CInt(cboBasicType.ListIndex))
  GetBasicDataSM CInt(cboBasicType.ListIndex), xSM
  txtSGtank.Text = AGFormat$(xSM.SpecGrav)
  txtSGnonv.Text = AGFormat$(xSM.NonVGrav)
  txtEvapRate.Text = AGFormat$(xSM.EvapRate)
  UpdatePropertyControls
  
  PropTakeAction = temp   'restore flag value
End Sub

Private Sub ChangeMatType(NewType As Integer)
'Select a new Material Type and do what is necessary to
'get new data
  Me.MousePointer = vbHourglass 'change pointer to hourglass
  Select Case NewType
  Case 0 'basic
    cboBasicType_Click
  Case 1 'user-def
    'nothing to do here, except a form cleanup on exit
  Case 2 'library
    'turn off Property actions so that the lib form can
    'modify this form's controls without triggering other actions
    temp = PropTakeAction  'save flag state
    PropTakeAction = False 'disable actions on change
    
    Dim fLib As frmSprayLib
    Set fLib = New frmSprayLib
    With fLib
      'no need to set form properties going in; this is strictly a lookup operation
      .Show vbModal
      PropTakeAction = temp  'restore flag value
      If .Cancelled Then 'Tag holds status info
        'reset original dist type
        temp = OptTakeAction  'save flag state
        OptTakeAction = False 'disable actions on change
        optMatType(SaveMatType).Value = True  'reset option button
        OptTakeAction = temp  'restore flag value
      Else
        'retrieve the library properties and tansfer them to local controls
        temp = PropTakeAction  'save flag state
        PropTakeAction = False 'disable actions on change
        txtName.Text = .SMName
        txtSGnonv.Text = AGFormat$(.NonVGrav)
        txtEvapRate.Text = AGFormat$(.EvapRate)
        Dim flow As Single
        flow = UnitsInternal(txtFlowRate.Text, UN_RATEVOL)
        If UD.Smokey = 0 Then 'regulatory
          txtNonvol.Text = AGFormat$(UnitsDisplay(.NVFrac * flow * .NonVGrav, UN_RATEMASS)) 'nonvol amount
          txtActive.Text = AGFormat$(UnitsDisplay(.ACFrac * flow * .NonVGrav, UN_RATEMASS)) 'active amount
        Else
          txtNonvol.Text = AGFormat$(.NVFrac) 'nonvol fraction
          txtActive.Text = AGFormat$(.ACFrac) 'active fraction
        End If
        PropTakeAction = temp  'restore flag value
      End If
    End With
    Unload fLib
    Set fLib = Nothing
  End Select
  'adjust the Type controls to reflect the current type
  UpdateTypeControls
  UpdatePropertyControls
  Me.MousePointer = vbDefault 'change pointer back to default
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  FormToData
  mCancelled = False
  Me.Hide
End Sub

Private Sub DataToForm()
'transfer user data to form controls for editing
  
  'Options
  cboBasicType.ListIndex = mSM.BasicType  'combo box for Basic Type
  temp = OptTakeAction                      'save flag value
  OptTakeAction = False                     'disable actions for the following
  optMatType(mSM.Type) = True             'dist type radio buttons
  UpdateTypeControls
  OptTakeAction = temp                      'restore flag value
  
  'Properties
  temp = PropTakeAction                        'save flag value
  PropTakeAction = False                       'allow raw field modification
  txtName.Text = RTrim$(mSM.Name)            'description
  txtSGtank.Text = AGFormat$(mSM.SpecGrav)         'specific gravity (tank)
  txtSGnonv.Text = AGFormat$(mSM.NonVGrav)         'specific gravity (nonv)
  txtEvapRate.Text = AGFormat$(mSM.EvapRate)   'Evaporation rate
  If UD.Smokey = 0 Then 'regulatory
    txtNonvol.Text = AGFormat$(UnitsDisplay(mSM.NVFrac * mSM.FlowRate * mSM.NonVGrav, UN_RATEMASS)) 'nonvol amount
    txtActive.Text = AGFormat$(UnitsDisplay(mSM.ACFrac * mSM.FlowRate * mSM.NonVGrav, UN_RATEMASS)) 'active amount
  Else
    txtNonvol.Text = AGFormat$(mSM.NVFrac) 'nonvol fraction
    txtActive.Text = AGFormat$(mSM.ACFrac) 'active fraction
  End If
  txtFlowRate.Text = AGFormat$(UnitsDisplay(mSM.FlowRate, UN_RATEVOL)) 'Flow Rate
  UpdatePropertyControls
  PropTakeAction = temp                     'restore flag value

End Sub

'---------------------------------------------------------------------------
' cmdSprayMatMix_Click:
'
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-10-20  TBC  Created
'
'---------------------------------------------------------------------------
Private Sub cmdSprayMatMix_Click()
  Dim strErrLocation As String
  On Error GoTo Error_Handler

  Dim fMix As frmSprayMatMix
  Set fMix = New frmSprayMatMix
  With fMix
    FormToData
    
    'Send the form only those SM elements that it is interested in
    .SMType = mSM.Type
    .BasicType = mSM.BasicType
    .SMName = mSM.Name
    .SMLName = mSM.LName
    .CalcInputSelect = mSM.CalcInputSelect
    .NVFrac = mSM.NVFrac
    .ACFrac = mSM.ACFrac
    .ActSolFrac = mSM.ActSolFrac
    .AddSolFrac = mSM.AddSolFrac
    .ActNVFrac = mSM.ActNVFrac
    .AddNVFrac = mSM.AddNVFrac
    .FlowRate = mSM.FlowRate
    .FlowRateUnits = mSM.FlowRateUnits
    .SpecGrav = mSM.SpecGrav
    .NonVGrav = mSM.NonVGrav
    .EvapRate = mSM.EvapRate
    
    .Show vbModal
    
    If Not .Cancelled Then
      mSM.Type = .SMType
      mSM.BasicType = .BasicType
      mSM.Name = .SMName
      mSM.LName = .SMLName
      mSM.CalcInputSelect = .CalcInputSelect
      mSM.NVFrac = .NVFrac
      mSM.ACFrac = .ACFrac
      mSM.ActSolFrac = .ActSolFrac
      mSM.AddSolFrac = .AddSolFrac
      mSM.ActNVFrac = .ActNVFrac
      mSM.AddNVFrac = .AddNVFrac
      mSM.FlowRate = .FlowRate
      mSM.FlowRateUnits = .FlowRateUnits
      mSM.SpecGrav = .SpecGrav
      mSM.NonVGrav = .NonVGrav
      mSM.EvapRate = .EvapRate
      
      DataToForm
    End If
  End With
  Unload fMix
  Set fMix = Nothing


'====================================================
'Exit Point for cmdSprayMatMix_Click
'====================================================
Exit_cmdSprayMatMix_Click:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "cmdSprayMatMix_Click", "frmSprayMat", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_cmdSprayMatMix_Click
End Sub

Private Sub Form_Load()
  InitForm
End Sub

Private Sub FormToData()
'Place the form data in user data storage
  
  Dim nlong As Long

  'find the current type selection
  For i = 0 To 2
    If optMatType(i).Value = True Then DT = i
  Next
  'save the current type selection
  mSM.Type = DT

  'get Basic selection, even if the type isn't Basic
  mSM.BasicType = cboBasicType.ListIndex  'Basic selection
  mSM.Name = RTrim$(txtName.Text)
  mSM.LName = Len(RTrim$(txtName.Text))
  
  mSM.SpecGrav = Val(txtSGtank.Text)           'specific gravity (tank)
  mSM.NonVGrav = Val(txtSGnonv.Text)           'specific gravity (nonv)
  mSM.EvapRate = Val(txtEvapRate.Text)     'Evaporation rate
  mSM.FlowRate = UnitsInternal(Val(txtFlowRate.Text), UN_RATEVOL)    'Flow Rate
  If UD.Smokey = 0 Then 'regulatory
    If (mSM.FlowRate * mSM.NonVGrav) > 0 Then
      mSM.NVFrac = UnitsInternal(Val(txtNonvol.Text), UN_RATEMASS) / _
                     (mSM.FlowRate * mSM.NonVGrav)
      mSM.ACFrac = UnitsInternal(Val(txtActive.Text), UN_RATEMASS) / _
                     (mSM.FlowRate * mSM.NonVGrav)
    Else
      mSM.NVFrac = 0
      mSM.ACFrac = 0
    End If
  Else 'FS
      mSM.NVFrac = Val(txtNonvol.Text)
      mSM.ACFrac = Val(txtActive.Text)
  End If

  UpdateDataChangedFlag True 'Data was changed
  UC.Valid = False 'Calcs need to be redone
End Sub

Private Sub InitForm()
'Initialize the controls on this form

  mCancelled = True 'set to true only if OK pressed
  
  'center the form
  CenterForm Me

  'init the combo box
  cboBasicType.Clear
  For i = 0 To 1
    cboBasicType.AddItem GetBasicNameSM(i)
  Next

  'allow option button changes to take action
  '(see declarations section)
  OptTakeAction = True
  PropTakeAction = True

  'set the units labels
  If UD.Smokey = 0 Then 'regulatory
    lblNonvol.Caption = "Nonvol. Rate:"
    lblNonvolUnits.Caption = UnitsName(UN_RATEMASS)
    lblActive.Caption = "Active Rate:"
    lblActiveUnits.Caption = UnitsName(UN_RATEMASS)
  Else 'FS
    lblNonvol.Caption = "Nonvol. Fraction:"
    lblNonvolUnits.Caption = ""
    lblActive.Caption = "Active Fraction:"
    lblActiveUnits.Caption = ""
  End If
  lblFlowRateUnits.Caption = UnitsName(UN_RATEVOL)

  'fill the controls with user data
  DataToForm
End Sub

Private Sub optMatType_Click(Index As Integer)
  If OptTakeAction Then ChangeMatType Index
End Sub

Private Sub optMatType_DblClick(Index As Integer)
'same as click
  optMatType_Click Index
End Sub

Private Sub txtActive_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optMatType(1).Value Then optMatType(1).Value = True
  End If
End Sub

Private Sub txtEvapRate_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optMatType(1).Value Then optMatType(1).Value = True
  End If
End Sub

Private Sub txtFlowRate_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optMatType(1).Value Then optMatType(1).Value = True
  End If
End Sub

Private Sub txtName_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optMatType(1).Value Then optMatType(1).Value = True
  End If
End Sub

Private Sub txtNonvol_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optMatType(1).Value Then optMatType(1).Value = True
  End If
End Sub

Private Sub txtSGtank_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optMatType(1).Value Then optMatType(1).Value = True
  End If
End Sub

Private Sub txtSGnonv_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    If Not optMatType(1).Value Then optMatType(1).Value = True
  End If
End Sub

Private Sub UpdatePropertyControls()
'update the state of the property controls

  'Evaporation rate controls
  lblEvapRate.Enabled = True
  txtEvapRate.Enabled = True
  lblEvapRateUnits.Enabled = True
  'disable for Basic,Oil
  If optMatType(0).Value And cboBasicType.ListIndex = 0 Then  'oil
    lblEvapRate.Enabled = False
    txtEvapRate.Enabled = False
    lblEvapRateUnits.Enabled = False
  End If
End Sub

Private Sub UpdateTypeControls()
'Adjust the Type Controls to conform to the current
'setting of the Type Option buttons

  'find the current type selection
  For i = 0 To 2
    If optMatType(i).Value = True Then DT = i
  Next
  'save the current type selection
  SaveMatType = DT

  'set related controls to a known state
  cboBasicType.Enabled = False
  'make adjustments
  Select Case DT
    Case 0 'Basic
      cboBasicType.Enabled = True
    Case 1 'user-def
    Case 2 'library
  End Select

End Sub

Public Property Get SMType() As Integer

  SMType = mSM.Type
End Property

Public Property Let SMType(ByVal vNewValue As Integer)

  mSM.Type = vNewValue
End Property

Public Property Get BasicType() As Integer

  BasicType = mSM.BasicType
End Property

Public Property Let BasicType(ByVal vNewValue As Integer)

  mSM.BasicType = vNewValue
End Property

Public Property Get SMName() As String
  
  SMName = mSM.Name
End Property

Public Property Let SMName(ByVal vNewValue As String)

  mSM.Name = vNewValue
End Property

Public Property Get SMLName() As Integer

  SMLName = mSM.LName
End Property

Public Property Let SMLName(ByVal vNewValue As Integer)

  mSM.LName = vNewValue
End Property

Public Property Get CalcInputSelect() As Integer

  CalcInputSelect = mSM.CalcInputSelect
End Property

Public Property Let CalcInputSelect(ByVal vNewValue As Integer)

  mSM.CalcInputSelect = vNewValue
End Property

Public Property Get NVFrac() As Single

  NVFrac = mSM.NVFrac
End Property

Public Property Let NVFrac(ByVal vNewValue As Single)

  mSM.NVFrac = vNewValue
End Property

Public Property Get ACFrac() As Single

  ACFrac = mSM.ACFrac
End Property

Public Property Let ACFrac(ByVal vNewValue As Single)

  mSM.ACFrac = vNewValue
End Property

Public Property Get ActSolFrac() As Single

  ActSolFrac = mSM.ActSolFrac
End Property

Public Property Let ActSolFrac(ByVal vNewValue As Single)

  mSM.ActSolFrac = vNewValue
End Property

Public Property Get AddSolFrac() As Single

  AddSolFrac = mSM.AddSolFrac
End Property

Public Property Let AddSolFrac(ByVal vNewValue As Single)

  mSM.AddSolFrac = vNewValue
End Property

Public Property Get ActNVFrac() As Single

  ActNVFrac = mSM.ActNVFrac
End Property

Public Property Let ActNVFrac(ByVal vNewValue As Single)

  mSM.ActNVFrac = vNewValue
End Property

Public Property Get AddNVFrac() As Single

  AddNVFrac = mSM.AddNVFrac
End Property

Public Property Let AddNVFrac(ByVal vNewValue As Single)

  mSM.AddNVFrac = vNewValue
End Property

Public Property Get FlowRate() As Single

  FlowRate = mSM.FlowRate
End Property

Public Property Let FlowRate(ByVal vNewValue As Single)

  mSM.FlowRate = vNewValue
End Property

Public Property Get FlowRateUnits() As Integer

  FlowRateUnits = mSM.FlowRateUnits
End Property

Public Property Let FlowRateUnits(ByVal vNewValue As Integer)

  mSM.FlowRateUnits = vNewValue
End Property

Public Property Get SpecGrav() As Single

  SpecGrav = mSM.SpecGrav
End Property

Public Property Let SpecGrav(ByVal vNewValue As Single)

  mSM.SpecGrav = vNewValue
End Property

Public Property Get NonVGrav() As Single

  NonVGrav = mSM.NonVGrav
End Property

Public Property Let NonVGrav(ByVal vNewValue As Single)

  mSM.NonVGrav = vNewValue
End Property

Public Property Get EvapRate() As Single

  EvapRate = mSM.EvapRate
End Property

Public Property Let EvapRate(ByVal vNewValue As Single)

  mSM.EvapRate = vNewValue
End Property

Public Property Get Cancelled() As Boolean

  Cancelled = mCancelled
End Property

