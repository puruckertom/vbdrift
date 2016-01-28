VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmSprayMatMix 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spray Material Tank Mix"
   ClientHeight    =   8130
   ClientLeft      =   1710
   ClientTop       =   2280
   ClientWidth     =   9240
   ForeColor       =   &H80000008&
   HelpContextID   =   1549
   Icon            =   "frmSprayMatMix.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8130
   ScaleWidth      =   9240
   Begin VB.Frame fraRates 
      Caption         =   "Rates"
      Height          =   1815
      Left            =   4920
      TabIndex        =   31
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtActive 
         Height          =   285
         HelpContextID   =   1010
         Left            =   1920
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtNonvol 
         Height          =   285
         HelpContextID   =   1180
         Left            =   1920
         TabIndex        =   16
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblActive 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Active Rate:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   705
         TabIndex        =   35
         Top             =   765
         Width           =   1080
      End
      Begin VB.Label lblActiveUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2880
         TabIndex        =   34
         Top             =   780
         Width           =   420
      End
      Begin VB.Label lblNonvol 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nonvolatile Rate:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   570
         TabIndex        =   33
         Top             =   1125
         Width           =   1230
      End
      Begin VB.Label lblNonvolUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2880
         TabIndex        =   32
         Top             =   1140
         Width           =   420
      End
   End
   Begin VB.Frame fraTankMix 
      Caption         =   "Tank Mix"
      Height          =   5535
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   9015
      Begin VB.Frame fraActive 
         Caption         =   "Active Solution"
         Height          =   1455
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   2535
         Begin VB.TextBox txtTankActivePercent 
            Height          =   285
            HelpContextID   =   1549
            Left            =   1560
            TabIndex        =   6
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtActiveNVFrac 
            Height          =   285
            HelpContextID   =   1549
            Left            =   1560
            TabIndex        =   7
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblTankActivePercent 
            AutoSize        =   -1  'True
            Caption         =   "% of Tank Mix:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label lblActiveNVFrac 
            Caption         =   "Fraction of Active Solution that is nonvolatile:"
            Height          =   615
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame fraAdditive 
         Caption         =   "Additive Solution(s)"
         Height          =   1455
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   2535
         Begin VB.TextBox txtAdditiveNVFrac 
            Height          =   285
            HelpContextID   =   1549
            Left            =   1560
            TabIndex        =   9
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtTankAdditivePercent 
            Height          =   285
            HelpContextID   =   1549
            Left            =   1560
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblAdditiveNVFrac 
            Caption         =   "Fraction of Additive Solution(s) that is nonvolatile:"
            Height          =   615
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblTankAdditivePercent 
            AutoSize        =   -1  'True
            Caption         =   "% of Tank Mix:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   1050
         End
      End
      Begin VB.Frame fraWater 
         Caption         =   "Carrier"
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   3480
         Width           =   2535
         Begin VB.TextBox txtTankWaterPercent 
            BackColor       =   &H8000000F&
            Height          =   285
            HelpContextID   =   1543
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblTankWaterPercent 
            Caption         =   "% of Tank Mix:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1410
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total"
         Height          =   735
         Left            =   120
         TabIndex        =   36
         Top             =   4200
         Width           =   2535
         Begin VB.TextBox txtTotalPercent 
            BackColor       =   &H8000000F&
            Height          =   285
            HelpContextID   =   1543
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblTotalPercentLabel 
            Caption         =   "% of Tank Mix"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1335
         End
      End
      Begin MSChart20Lib.MSChart chtSummary 
         Height          =   4935
         Left            =   2760
         OleObjectBlob   =   "frmSprayMatMix.frx":030A
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblMessage"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1680
         TabIndex        =   30
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame fraProperties 
      Caption         =   "Properties"
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox cboFlowRateUnits 
         Height          =   315
         HelpContextID   =   1549
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkEvaporation 
         Caption         =   "Spray Material Evaporates"
         Height          =   255
         HelpContextID   =   1549
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         Height          =   285
         HelpContextID   =   1549
         Left            =   720
         TabIndex        =   2
         Text            =   "txtName"
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtFlowRate 
         Height          =   285
         HelpContextID   =   1549
         Left            =   1650
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblFlowRate 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Spray Volume Rate:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1425
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1549
      Left            =   8280
      TabIndex        =   1
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      HelpContextID   =   1549
      Left            =   7320
      TabIndex        =   0
      Top             =   7560
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calculation Control"
      Height          =   735
      Left            =   120
      TabIndex        =   38
      Top             =   7320
      Width           =   4455
      Begin VB.CommandButton cmdCalc 
         Caption         =   "Calc"
         Default         =   -1  'True
         Height          =   375
         HelpContextID   =   1549
         Left            =   3360
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optInputSelect 
         Caption         =   "Tank Mix"
         Height          =   255
         HelpContextID   =   1549
         Index           =   1
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optInputSelect 
         Caption         =   "Rates"
         Height          =   255
         HelpContextID   =   1549
         Index           =   0
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Enter"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSprayMatMix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: frmSprayMatMix.frm,v 1.2 2008/10/22 17:26:06 tom Exp $
Option Explicit
'File: frmSprayMat
'
'----------------------------------------------------------------------
'Re:
'    Spray Material form.
'
'----------------------------------------------------------------------
'
' Application defined Errors (vbObjectError + # )
' Number         Description
' ============   =================================================
'
'
'----------------------------------------------------------------------
' CONSTANTS:
'----------------------------------------------------------------------
'
'
'
'----------------------------------------------------------------------
' PRIVATE VARIABLES:
'----------------------------------------------------------------------

'this flag is used to tell the option buttons not to
'take action on their new values. This is required
'to differentiate between programatic state changes
'and user actions
Private OptTakeAction As Integer  'if true, execute automatic change-related code
                              'for Spray Mat option button
Private PropTakeAction As Integer 'if true, execute automatic change-related code
                              'for Property text boxes

'Form's copy of spray material data
Private mSM As SprayMaterialData

'Form exit status
Private mCancelled As Boolean 'If true, form was cancelled

'Set these up after loading the form
Private mACFrac As Single       'Tank Mix Active Fraction
Private mNVFrac As Single       'Tank Mix Nonvolatile Fraction


'---------------------------------------------------------------------------
' cboFlowRateUnits_Click:
'    Flow Rate Units combo box Click Event. Change the value of the flow rate to match the new units.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-10  TBC  Created
'
'---------------------------------------------------------------------------
Private Sub cboFlowRateUnits_Click()
  Dim strErrLocation As String
  On Error GoTo Error_Handler

  If PropTakeAction Then
    Dim sglFlow As Single
    Select Case cboFlowRateUnits.ListIndex
    Case 0 'change from L/min to L/ha
      sglFlow = UnitsInternal(Val(txtFlowRate.Text), UN_FLOWRATE)
      sglFlow = sglFlow / (UD.CTL.SwathWidth * UD.AC.TypSpeed * 0.006)
      txtFlowRate.Text = AGFormat(UnitsDisplay(sglFlow, UN_RATEVOL))
    Case 1 'change from L/ha to L/min
      sglFlow = UnitsInternal(Val(txtFlowRate.Text), UN_RATEVOL)
      sglFlow = sglFlow * (UD.CTL.SwathWidth * UD.AC.TypSpeed * 0.006)
      txtFlowRate.Text = AGFormat(UnitsDisplay(sglFlow, UN_FLOWRATE))
    End Select
  End If
  

'====================================================
'Exit Point for cboFlowRateUnits_Click
'====================================================
Exit_cboFlowRateUnits_Click:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "cboFlowRateUnits_Click", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_cboFlowRateUnits_Click
End Sub


'---------------------------------------------------------------------------
' chkEvaporation_Click:
'    Evaporation checkbox. Recalculate on change.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub chkEvaporation_Click()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If PropTakeAction Then
    Calculate
  End If


'====================================================
'Exit Point for chkEvaporation_Click
'====================================================
Exit_chkEvaporation_Click:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "chkEvaporation_Click", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_chkEvaporation_Click
End Sub

'---------------------------------------------------------------------------
' cmdCalc_Click:
'    Calculate button. Perform calculations.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub cmdCalc_Click()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  Calculate


'====================================================
'Exit Point for cmdCalc_Click
'====================================================
Exit_cmdCalc_Click:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "cmdCalc_Click", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_cmdCalc_Click
End Sub

'---------------------------------------------------------------------------
' cmdCancel_Click:
'    Cancel button. Unload this form without saving changes.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub cmdCancel_Click()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  Unload Me


'====================================================
'Exit Point for cmdCancel_Click
'====================================================
Exit_cmdCancel_Click:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "cmdCancel_Click", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_cmdCancel_Click
End Sub

'---------------------------------------------------------------------------
' cmdOK_Click:
'    OK button. Dismiss form and save changes.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub cmdOk_Click()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If FormToData Then
    mCancelled = False 'Form was not cancelled
    Me.Hide
  End If


'====================================================
'Exit Point for cmdOK_Click
'====================================================
Exit_cmdOK_Click:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "cmdOK_Click", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_cmdOK_Click
End Sub

'---------------------------------------------------------------------------
' Form_Load:
'    Form Load event. Set up the form controls.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub Form_Load()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  'Initialize Cancelled property. Set to false only if Ok is pressed
  mCancelled = True
  
  'center the form
  CenterForm Me

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
  'Set up the flow rate combo
  With cboFlowRateUnits
    .Clear
    .AddItem UnitsName(UN_RATEVOL)
    .AddItem UnitsName(UN_FLOWRATE)
  End With

  DataToForm
  
  'Initialize the tank mix calcs
  Calculate


'====================================================
'Exit Point for Form_Load
'====================================================
Exit_Form_Load:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "Form_Load", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_Form_Load
End Sub

'---------------------------------------------------------------------------
' FormToData:
'    Place the form data in user data storage
'
' Return: True if initial calculations are successful. False if not.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Function FormToData() As Boolean
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  Dim nlong As Long
  

  'Make sure there is valid data for the tank mix calc
  If Trim$(txtNonvol.Text) = "" Or Trim$(txtActive.Text) = "" Then
    MsgBox "Cannot set Nonvolatile Fraction and Active Fraction because " & _
           "the entered data is incorrect. Correct the data or click " & _
           "Cancel.", vbExclamation + vbOKOnly
    FormToData = False
    Exit Function
  End If
  
  mSM.Name = RTrim(txtName.Text)
  mSM.LName = Len(mSM.Name)
  mSM.BasicType = chkEvaporation.Value
  
  mSM.FlowRateUnits = cboFlowRateUnits.ListIndex
  If mSM.FlowRateUnits = 0 Then
    mSM.FlowRate = UnitsInternal(Val(txtFlowRate.Text), UN_RATEVOL)    'Flow Rate
  Else
    mSM.FlowRate = UnitsInternal(Val(txtFlowRate.Text), UN_FLOWRATE)    'Flow Rate
  End If
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
  mSM.ActSolFrac = Val(txtTankActivePercent.Text) / 100
  mSM.ActNVFrac = Val(txtActiveNVFrac.Text)
  mSM.AddSolFrac = Val(txtTankAdditivePercent.Text) / 100
  mSM.AddNVFrac = Val(txtAdditiveNVFrac.Text)
  Dim c As Control
  For Each c In optInputSelect()
    If c.Value Then
      mSM.CalcInputSelect = c.Index
      Exit For
    End If
  Next

  UpdateDataChangedFlag True 'Data was changed
  UC.Valid = False 'Calcs need to be redone
  
  FormToData = True


'====================================================
'Exit Point for FormToData
'====================================================
Exit_FormToData:
  Exit Function


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "FormToData", "frmSprayMatMix", strErrLocation
  gobjErrors.RaiseError
End Function

'---------------------------------------------------------------------------
' CalculateWithTankMix:
'    Calculate the Active Fraction and Nonvolatile fraction
'
'     Given:
'      AddSolFrac
'      ActSolFrac
'      AddNVFrac
'      ActNVFrac
'
'     Calculate:
'      NVfrac
'      ACfrac
'      MixWaterFrac
'      TankActNVfrac
'      TankAddNVfrac
'      TankVolFrac
'
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub CalculateWithTankMix()
  Dim strErrLocation As String
  On Error GoTo Error_Handler

  Dim MixWaterFrac As Single 'Fraction of Tank Mix that is Mixing Water
  Dim AddSolFrac As Single   'Fraction of Tank Mix that is Additive Solution(s)
  Dim ActSolFrac As Single   'Fraction of Tank Mix that is Active Solution
  Dim AddNVFrac As Single    'Fraction of Additive Solution that is nonvolatile
  Dim ActNVFrac As Single    'Fraction of Active Solution that is nonvolatile
  Dim TankAddNVfrac As Single 'Fraction of Tank Mix that is nonvolative Additive
  Dim TankActNVfrac As Single 'Fraction of Tank Mix that is nonvolative Active
  Dim TankVolFrac As Single   'Fraction of Tank mix that is volatile
  Dim SprayRate As Single     'Local copy of the spray Volume Rate
  
  ClearOutputControls
  
  'Gather input values from the controls
  AddSolFrac = Val(txtTankAdditivePercent.Text) / 100
  ActSolFrac = Val(txtTankActivePercent.Text) / 100
  AddNVFrac = Val(txtAdditiveNVFrac.Text)
  ActNVFrac = Val(txtActiveNVFrac.Text)
  SprayRate = Val(txtFlowRate.Text)
  
  'Make sure SprayRate is in L/ha
  If cboFlowRateUnits.ListIndex = 1 Then
    SprayRate = SprayRate / (UD.CTL.SwathWidth * UD.AC.TypSpeed * 0.006)
  End If
  
  'Check before calulating
  If ActSolFrac < 0 Or ActSolFrac > 1 Then
    lblMessage = "% Tank Mix for Active Solution must be from 0 to 100"
    Exit Sub
  End If
  If AddSolFrac < 0 Or AddSolFrac > 1 Then
    lblMessage = "% Tank Mix for Additive Solution(s) must be from 0 to 100"
    Exit Sub
  End If
  If ActSolFrac + AddSolFrac > 1 Then
    lblMessage = "Sum of % Tank Mixes cannot exceed 100"
    Exit Sub
  End If
  If AddNVFrac < 0 Or AddNVFrac > 1 Then
    lblMessage = "Fraction of Additive Solution that is nonvolatile must be from 0 to 1"
    Exit Sub
  End If
  If ActNVFrac < 0 Or ActNVFrac > 1 Then
    lblMessage = "Fraction of Active Solution that is nonvolatile must be from 0 to 1"
    Exit Sub
  End If
  lblMessage = ""
  
  'Calc output values
  If UD.Smokey Then
    mNVFrac = AddNVFrac * AddSolFrac + ActNVFrac * ActSolFrac
    mACFrac = ActNVFrac * ActSolFrac
    MixWaterFrac = 1 - ActSolFrac - AddSolFrac
    TankAddNVfrac = AddNVFrac * AddSolFrac
    TankActNVfrac = ActNVFrac * ActSolFrac
    TankVolFrac = 1 - ActNVFrac * ActSolFrac - AddNVFrac * AddSolFrac
  Else
    mNVFrac = AddNVFrac * AddSolFrac * NonVGrav * SprayRate + _
              ActNVFrac * ActSolFrac * SpecGrav * SprayRate
    mACFrac = ActNVFrac * ActSolFrac * SpecGrav * SprayRate
    MixWaterFrac = 1 - ActSolFrac - AddSolFrac
    TankAddNVfrac = AddNVFrac * AddSolFrac
    TankActNVfrac = ActNVFrac * ActSolFrac
    TankVolFrac = 1 - ActNVFrac * ActSolFrac - AddNVFrac * AddSolFrac
  End If
  
  'Populate controls
  PropTakeAction = False
  txtNonvol.Text = AGFormat(mNVFrac)
  txtActive.Text = AGFormat(mACFrac)
  txtTankWaterPercent.Text = AGFormat(MixWaterFrac * 100)
  txtTotalPercent.Text = AGFormat((MixWaterFrac + AddSolFrac + ActSolFrac) * 100)
  PropTakeAction = True
    
  'Update the pie chart
  UpdatePieChart TankActNVfrac, TankAddNVfrac, TankVolFrac, CBool(chkEvaporation.Value)

  cmdOK.Enabled = True

'====================================================
'Exit Point for CalculateWithTankMix
'====================================================
Exit_CalculateWithTankMix:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "CalculateWithTankMix", "frmSprayMatMix", strErrLocation
  gobjErrors.RaiseError
End Sub

'---------------------------------------------------------------------------
' optInputSelect_Click:
'    Select between entering Tank Mix or Active Rates
'
' Arguments:
'     Index: Control array index
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error logging
'
'---------------------------------------------------------------------------
Private Sub optInputSelect_Click(Index As Integer)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  UpdateInputControls


'====================================================
'Exit Point for optInputSelect_Click
'====================================================
Exit_optInputSelect_Click:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "optInputSelect_Click", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_optInputSelect_Click
End Sub

'---------------------------------------------------------------------------
' CalculateWithRates:
'    Calculate starting form values based on existing
'    numbers and some assumptions.
'    Assumptions: Fraction of Active Solution that is nonvolatile = 0.1
'                 Fraction of Additive Solution that is nonvolatile = 0.1
'
'     Given:
'      NVfrac
'      ACfrac
'      AddNVFrac
'      ActNVFrac
'      FlowRate
'
'     Calculate:
'      AddSolFrac
'      ActSolFrac
'      MixWaterFrac
'      TankActNVfrac
'      TankAddNVfrac
'      TankVolFrac
'
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub CalculateWithRates()
  Dim strErrLocation As String
  On Error GoTo Error_Handler

  Dim MixWaterFrac As Single 'Fraction of Tank Mix that is Mixing Water
  Dim AddSolFrac As Single   'Fraction of Tank Mix that is Additive Solution(s)
  Dim ActSolFrac As Single   'Fraction of Tank Mix that is Active Solution
  Dim AddNVFrac As Single    'Fraction of Additive Solution that is nonvolatile
  Dim ActNVFrac As Single    'Fraction of Active Solution that is nonvolatile
  Dim TankAddNVfrac As Single 'Fraction of Tank Mix that is nonvolative Additive
  Dim TankActNVfrac As Single 'Fraction of Tank Mix that is nonvolative Active
  Dim TankVolFrac As Single   'Fraction of Tank mix that is volatile
  Dim SprayRate As Single      'Spray Volume Rate in L/ha
  
  ClearOutputControls
  
  'Gather input values from the controls
  mACFrac = Val(txtActive.Text)
  mNVFrac = Val(txtNonvol.Text)
  ActNVFrac = Val(txtActiveNVFrac.Text)
  AddNVFrac = Val(txtAdditiveNVFrac.Text)
  SprayRate = Val(txtFlowRate.Text)
  
  'Make sure SprayRate is in L/ha
  If cboFlowRateUnits.ListIndex = 1 Then
    SprayRate = SprayRate / (UD.CTL.SwathWidth * UD.AC.TypSpeed * 0.006)
  End If
  
  'Check inputs before calculating
  If UD.Smokey Then
    If mACFrac < 0 Or mACFrac > 1 Then
      lblMessage = "Active Fraction must be from 0 to 1"
      Exit Sub
    End If
    If mNVFrac < 0 Or mNVFrac > 1 Then
      lblMessage = "Nonvolatile Fraction must be from 0 to 1"
      Exit Sub
    End If
  Else
    If mACFrac < 0 Then
      lblMessage = "Active Rate must be greater than or equal to 0"
      Exit Sub
    End If
    If mNVFrac < 0 Then
      lblMessage = "Nonvolatile Rate must be greater than or equal to 0"
      Exit Sub
    End If
  End If
  If ActNVFrac < 0 Or ActNVFrac > 1 Then
    lblMessage = "Fraction of Active Solution that is Nonvolatile must be from 0 to 1"
    Exit Sub
  End If
  If AddNVFrac < 0 Or AddNVFrac > 1 Then
    lblMessage = "Fraction of Additive Solution that is Nonvolatile must be from 0 to 1"
    Exit Sub
  End If
  If ActNVFrac = 0 Or AddNVFrac = 0 Then
    lblMessage = "Cannot determine % of Tank Mix when Fraction of Solution that is Nonvolatile is zero."
    Exit Sub
  End If
  If mACFrac > mNVFrac Then
    lblMessage = "Active Fraction cannot be greater than the Nonvolatile Fraction."
    Exit Sub
  End If
  If SprayRate <= 0 Then
    lblMessage = "Spray Volume Rate must be greater than 0."
    Exit Sub
  End If
  lblMessage = ""
  
  'Calculate output values
  If UD.Smokey Then
    ActSolFrac = mACFrac / ActNVFrac
    AddSolFrac = (mNVFrac - ActNVFrac * ActSolFrac) / AddNVFrac
    MixWaterFrac = 1 - ActSolFrac - AddSolFrac
    TankAddNVfrac = AddNVFrac * AddSolFrac
    TankActNVfrac = ActNVFrac * ActSolFrac
    TankVolFrac = 1 - ActNVFrac * ActSolFrac - AddNVFrac * AddSolFrac
  Else
    ActSolFrac = mACFrac / ActNVFrac / SpecGrav / SprayRate
    AddSolFrac = (mNVFrac - mACFrac) / AddNVFrac / NonVGrav / SprayRate
    MixWaterFrac = 1 - ActSolFrac - AddSolFrac
    TankAddNVfrac = AddNVFrac * AddSolFrac
    TankActNVfrac = ActNVFrac * ActSolFrac
    TankVolFrac = 1 - ActNVFrac * ActSolFrac - AddNVFrac * AddSolFrac
  End If
  If MixWaterFrac < 0 Then
    lblMessage = "Either Fraction of Active Solution that is nonvolatile or " & _
                 "Fraction of Additive Solution(s) that is nonvolatile is too low."
    Exit Sub
  End If
  
  'Populate controls
  PropTakeAction = False
  lblMessage.Caption = ""
  txtTankActivePercent.Text = AGFormat(ActSolFrac * 100)
  txtTankAdditivePercent.Text = AGFormat(AddSolFrac * 100)
  txtTankWaterPercent.Text = AGFormat(MixWaterFrac * 100)
  txtTotalPercent.Text = AGFormat((MixWaterFrac + AddSolFrac + ActSolFrac) * 100)
  PropTakeAction = True

  'Update the pie chart
  UpdatePieChart TankActNVfrac, TankAddNVfrac, TankVolFrac, CBool(chkEvaporation.Value)

  cmdOK.Enabled = True


'====================================================
'Exit Point for CalculateWithRates
'====================================================
Exit_CalculateWithRates:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "CalculateWithRates", "frmSprayMatMix", strErrLocation
  gobjErrors.RaiseError
End Sub

'---------------------------------------------------------------------------
' UpdatePieChart:
'    Update the pie cart with new values after a calculation
'
' Arguments:
'     TankActNVfrac: Active solution nonvolatile fraction
'     TankAddNVfrac: Additive solution nonvolatile fraction
'     TankVolFrac: Tank volume fraction
'     EvaporationFlag:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub UpdatePieChart(TankActNVfrac As Single, _
                           TankAddNVfrac As Single, _
                           TankVolFrac As Single, _
                           EvaporationFlag As Boolean)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  'pie chart
  With chtSummary
    .chartType = VtChChartType2dPie
    .RowCount = 1
    .ColumnCount = 3
    .ColumnLabelCount = 1
    .RowLabelCount = 0
    .ShowLegend = True
    .Row = 1
    .Column = 1
    .Data = TankActNVfrac * 100
    .ColumnLabel = "Nonvolatile Active (" & AGFormat(.Data) & " %)"
    .Column = 2
    .Data = TankAddNVfrac * 100
    .ColumnLabel = "Nonvolatile Additive(s) (" & AGFormat(.Data) & " %)"
    .Column = 3
    .Data = TankVolFrac * 100
    If EvaporationFlag Then
      .ColumnLabel = "Volatiles (" & AGFormat(.Data) & " %)"
    Else
      .ColumnLabel = "Other Nonvolatiles (" & AGFormat(.Data) & " %)"
    End If
    .Visible = True
  End With



'====================================================
'Exit Point for UpdatePieChart
'====================================================
Exit_UpdatePieChart:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "TankActNVfrac", CStr(TankActNVfrac)
  dicCommandLineArgs.Add "TankAddNVfrac", CStr(TankAddNVfrac)
  dicCommandLineArgs.Add "TankVolFrac", CStr(TankVolFrac)
  dicCommandLineArgs.Add "EvaporationFlag", IIf(EvaporationFlag, " True", " False")
  gobjErrors.Append Err, "UpdatePieChart", "frmSprayMatMix", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Sub

'---------------------------------------------------------------------------
' Calculate:
'    Perform tank mix calculations.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub Calculate()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If optInputSelect(0).Value Then 'Rate input
    CalculateWithRates
  ElseIf optInputSelect(1).Value Then 'Tank Mix input
    CalculateWithTankMix
  End If


'====================================================
'Exit Point for Calculate
'====================================================
Exit_Calculate:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "Calculate", "frmSprayMatMix", strErrLocation
  gobjErrors.RaiseError
End Sub

'---------------------------------------------------------------------------
' UpdateInputControls:
'    Update the state of the input controls to match the task at hand.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub UpdateInputControls()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If optInputSelect(0).Value Then 'Rate input
    txtActive.Locked = False
    txtActive.BackColor = txtName.BackColor
    txtNonvol.Locked = False
    txtNonvol.BackColor = txtName.BackColor
    
    txtTankActivePercent.Locked = True
    txtTankActivePercent.BackColor = fraProperties.BackColor
    txtTankAdditivePercent.Locked = True
    txtTankAdditivePercent.BackColor = fraProperties.BackColor
    
  ElseIf optInputSelect(1).Value Then 'Tank Mix input
    txtActive.Locked = True
    txtActive.BackColor = fraProperties.BackColor
    txtNonvol.Locked = True
    txtNonvol.BackColor = fraProperties.BackColor
    
    txtTankActivePercent.Locked = False
    txtTankActivePercent.BackColor = txtName.BackColor
    txtTankAdditivePercent.Locked = False
    txtTankAdditivePercent.BackColor = txtName.BackColor
  End If


'====================================================
'Exit Point for UpdateInputControls
'====================================================
Exit_UpdateInputControls:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "UpdateInputControls", "frmSprayMatMix", strErrLocation
  gobjErrors.RaiseError
End Sub

'---------------------------------------------------------------------------
' ClearOutputControls:
'    Clear the contents of the controls that display calculation output.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub ClearOutputControls()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  lblMessage.Caption = "Click Calc or press Enter to calculate"
  If optInputSelect(0).Value Then 'Rate input
    chtSummary.Visible = False
    PropTakeAction = False
    txtTankActivePercent.Text = ""
    txtTankAdditivePercent.Text = ""
    txtTankWaterPercent.Text = ""
    txtTotalPercent.Text = ""
    PropTakeAction = True
  ElseIf optInputSelect(1).Value Then 'Tank Mix input
    chtSummary.Visible = False
    PropTakeAction = False
    txtNonvol.Text = ""
    txtActive.Text = ""
    txtTankWaterPercent.Text = ""
    txtTotalPercent.Text = ""
    PropTakeAction = True
  End If
  cmdOK.Enabled = False 'Disable the Ok button. This will be enabled on successful calcs


'====================================================
'Exit Point for ClearOutputControls
'====================================================
Exit_ClearOutputControls:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "ClearOutputControls", "frmSprayMatMix", strErrLocation
  gobjErrors.RaiseError
End Sub

'---------------------------------------------------------------------------
' txtActive_Change:
'    Clear all controls when changed
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub txtActive_Change()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If PropTakeAction Then
    ClearOutputControls
  End If


'====================================================
'Exit Point for txtActive_Change
'====================================================
Exit_txtActive_Change:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "txtActive_Change", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_txtActive_Change
End Sub

'---------------------------------------------------------------------------
' txtActiveNVFrac_Change:
'    Clear all controls when changed
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub txtActiveNVFrac_Change()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If PropTakeAction Then
    ClearOutputControls
  End If


'====================================================
'Exit Point for txtActiveNVFrac_Change
'====================================================
Exit_txtActiveNVFrac_Change:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "txtActiveNVFrac_Change", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_txtActiveNVFrac_Change
End Sub

'---------------------------------------------------------------------------
' txtAdditiveNVFrac_Change:
'    Clear all controls when changed
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub txtAdditiveNVFrac_Change()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If PropTakeAction Then
    ClearOutputControls
  End If


'====================================================
'Exit Point for txtAdditiveNVFrac_Change
'====================================================
Exit_txtAdditiveNVFrac_Change:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "txtAdditiveNVFrac_Change", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_txtAdditiveNVFrac_Change
End Sub

'---------------------------------------------------------------------------
' txtFlowRate_Change:
'
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-10-30  TBC  Created
'
'---------------------------------------------------------------------------
Private Sub txtFlowRate_Change()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If PropTakeAction Then
    ClearOutputControls
  End If


'====================================================
'Exit Point for txtFlowRate_Change
'====================================================
Exit_txtFlowRate_Change:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "txtFlowRate_Change", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_txtFlowRate_Change
End Sub

'---------------------------------------------------------------------------
' txtNonvol_Change:
'    Clear all controls when changed
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub txtNonvol_Change()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If PropTakeAction Then
    ClearOutputControls
  End If


'====================================================
'Exit Point for txtNonvol_Change
'====================================================
Exit_txtNonvol_Change:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "txtNonvol_Change", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_txtNonvol_Change
End Sub

'---------------------------------------------------------------------------
' txtTankActivePercent_Change:
'    Clear all controls when changed
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub txtTankActivePercent_Change()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If PropTakeAction Then
    ClearOutputControls
  End If


'====================================================
'Exit Point for txtTankActivePercent_Change
'====================================================
Exit_txtTankActivePercent_Change:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "txtTankActivePercent_Change", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_txtTankActivePercent_Change
End Sub


'---------------------------------------------------------------------------
' txtTankAdditivePercent_Change:
'    Clear all controls when changed
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-06  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub txtTankAdditivePercent_Change()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If PropTakeAction Then
    ClearOutputControls
  End If


'====================================================
'Exit Point for txtTankAdditivePercent_Change
'====================================================
Exit_txtTankAdditivePercent_Change:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "txtTankAdditivePercent_Change", "frmSprayMatMix", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_txtTankAdditivePercent_Change
End Sub



'---------------------------------------------------------------------------
' DataToForm:
'    Fill form controls with local data.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-10-20  TBC  Created
'
'---------------------------------------------------------------------------
Private Sub DataToForm()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  'fill the controls with user data
  'Properties
  PropTakeAction = False                       'allow raw field modification
  txtName.Text = RTrim$(mSM.Name)            'description
  chkEvaporation.Value = mSM.BasicType       '0=oil (no evap) 1=water (evap)
  If UD.Smokey = 0 Then 'regulatory
    txtNonvol.Text = AGFormat$(UnitsDisplay(mSM.NVFrac * mSM.FlowRate * mSM.NonVGrav, UN_RATEMASS)) 'nonvol amount
    txtActive.Text = AGFormat$(UnitsDisplay(mSM.ACFrac * mSM.FlowRate * mSM.NonVGrav, UN_RATEMASS)) 'active amount
  Else
    txtNonvol.Text = AGFormat$(mSM.NVFrac) 'nonvol fraction
    txtActive.Text = AGFormat$(mSM.ACFrac) 'active fraction
  End If
  'Flowrate
  If mSM.FlowRateUnits = 0 Then
    txtFlowRate.Text = AGFormat$(UnitsDisplay(mSM.FlowRate, UN_RATEVOL)) 'Flow Rate
  Else
    txtFlowRate.Text = AGFormat$(UnitsDisplay(mSM.FlowRate, UN_FLOWRATE)) 'Flow Rate
  End If
  cboFlowRateUnits.ListIndex = mSM.FlowRateUnits '0 or 1
  txtTankActivePercent.Text = AGFormat$(mSM.ActSolFrac * 100)
  txtActiveNVFrac.Text = AGFormat$(mSM.ActNVFrac)
  
  txtTankAdditivePercent.Text = AGFormat$(mSM.AddSolFrac * 100)
  txtAdditiveNVFrac.Text = AGFormat$(mSM.AddNVFrac)
  PropTakeAction = True          'restore flag value
  
  'Select an initial input selection
  optInputSelect(mSM.CalcInputSelect).Value = True
  

'====================================================
'Exit Point for DataToForm
'====================================================
Exit_DataToForm:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "DataToForm", "frmSprayMatMix", strErrLocation
  gobjErrors.RaiseError
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

