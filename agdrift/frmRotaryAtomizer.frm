VERSION 5.00
Begin VB.Form frmRotaryAtomizer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FS Rotary Atomizer Models"
   ClientHeight    =   4785
   ClientLeft      =   1035
   ClientTop       =   2640
   ClientWidth     =   6825
   ForeColor       =   &H80000008&
   HelpContextID   =   1547
   Icon            =   "frmRotaryAtomizer.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4785
   ScaleWidth      =   6825
   Begin VB.Frame fraAtomizer 
      Caption         =   "Atomizer"
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   6615
      Begin VB.ComboBox cboAtomizerName 
         Height          =   315
         HelpContextID   =   1547
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1547
      Left            =   6000
      TabIndex        =   1
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1547
      Left            =   5160
      TabIndex        =   0
      Top             =   4320
      Width           =   735
   End
   Begin VB.Frame fraMaterial 
      Caption         =   "Spray Material"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   6615
      Begin VB.ComboBox cboSprayMaterial 
         Height          =   315
         HelpContextID   =   1547
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame fraSprayData 
      Caption         =   "Spray Data"
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   6615
      Begin VB.TextBox txtAirSpeed 
         BackColor       =   &H8000000F&
         Height          =   285
         HelpContextID   =   1547
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtBladeRPM 
         Height          =   285
         HelpContextID   =   1546
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtBladeAngle 
         Height          =   285
         HelpContextID   =   1544
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtFlowRate 
         BackColor       =   &H8000000F&
         Height          =   285
         HelpContextID   =   1547
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "(blank for model estimate)"
         Height          =   195
         Left            =   3480
         TabIndex        =   25
         Top             =   1080
         Width           =   1800
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Rotation Rate: (Optional)"
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   465
         TabIndex        =   24
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "rpm"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2760
         TabIndex        =   23
         Top             =   1035
         Width           =   255
      End
      Begin VB.Label lblNote 
         Alignment       =   2  'Center
         Caption         =   "Air Speed and Flow Rate must be consistent with drift model inputs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3360
         TabIndex        =   18
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblAirSpeedUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2760
         TabIndex        =   17
         Top             =   315
         Width           =   420
      End
      Begin VB.Label lblFlowRateUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         Top             =   1395
         Width           =   420
      End
      Begin VB.Label lblNozzleAngleUnits 
         AutoSize        =   -1  'True
         Caption         =   "deg"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2760
         TabIndex        =   13
         Top             =   675
         Width           =   270
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flow Rate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   570
         TabIndex        =   16
         Top             =   1395
         Width           =   930
      End
      Begin VB.Label lblNozzleAngle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Blade Angle:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   600
         TabIndex        =   15
         Top             =   660
         Width           =   900
      End
      Begin VB.Label lblAirSpeed 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   330
         Width           =   900
      End
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   4935
      Begin VB.OptionButton optDropDistributionType 
         Caption         =   "Drop Size Classification"
         ForeColor       =   &H80000008&
         Height          =   255
         HelpContextID   =   1547
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optDropDistributionType 
         Caption         =   "Drop Size Distribution"
         ForeColor       =   &H80000008&
         Height          =   255
         HelpContextID   =   1547
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmRotaryAtomizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: frmRotaryAtomizer.frm,v 1.2 2008/10/22 17:26:06 tom Exp $
Option Explicit
'File: frmRotaryAtomizer
'
'----------------------------------------------------------------------
'Re:
'    Rotary Atomizer form. Generates a DSD from Rotary Atomizer data using the FS model
'
'     Properties:
'       Cancelled
'
'       AtomizerIndex
'       SprayMaterialIndex
'       AirSpeed
'       BladeAngle
'       BladeRPM
'       FlowRate
'       DropDistributionType  0=Drop Size Classification 1=Drop Size Distribution
'
'       DropDistributionClassification 0-14 only valid if DropSizeOutputType=0
'       DropDistributionNumber
'       DropDistributionDiamter()
'       DropDistributionMassFraction()
'
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
Private mbolCancelled As Boolean
Private mintAtomizerIndex As Integer
Private mintSprayMaterialIndex As Integer
Private mintDropDistributionClassification As Integer
Private mintDropDistributionNumber As Integer
Private msglDropDistributionDiameter() As Single
Private msglDropDistributionMassFraction() As Single

'---------------------------------------------------------------------------
' cboAtomizerName_Click:
'    Atomizer Name combo click event. Update the Atomizer selection.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub cboAtomizerName_Click()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  With cboAtomizerName
    mintAtomizerIndex = .ItemData(.ListIndex)
    'If the current Spray Material is not legal for the new atomizer,
    'clear the Spray Material Selection
    If mintAtomizerIndex = 0 Then 'AU400
      With cboSprayMaterial
        If .ListIndex >= 0 Then
          If .ItemData(.ListIndex) = 1 Or .ItemData(.ListIndex) = 2 Then
            .Clear
          End If
        End If
      End With
    End If
  End With


'====================================================
'Exit Point for cboAtomizerName_Click
'====================================================
Exit_cboAtomizerName_Click:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "cboAtomizerName_Click", "frmRotaryAtomizer", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_cboAtomizerName_Click
End Sub

'---------------------------------------------------------------------------
' cboAtomizerName_DropDown:
'    Atomizer name combo dropdown event. Populate the list with all possible values.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub cboAtomizerName_DropDown()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  With cboAtomizerName
    .Clear
    Dim i As Integer
    For i = 0 To 1
      .AddItem GetNameHKNoz(i)
      .ItemData(.NewIndex) = i
    Next
    .ListIndex = mintAtomizerIndex
  End With


'====================================================
'Exit Point for cboAtomizerName_DropDown
'====================================================
Exit_cboAtomizerName_DropDown:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "cboAtomizerName_DropDown", "frmRotaryAtomizer", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_cboAtomizerName_DropDown
End Sub


'---------------------------------------------------------------------------
' cboSprayMaterial_Click:
'    Spary Material combo click event. Update selection of Spray Material.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub cboSprayMaterial_Click()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  With cboSprayMaterial
    mintSprayMaterialIndex = .ItemData(.ListIndex)
  End With


'====================================================
'Exit Point for cboSprayMaterial_Click
'====================================================
Exit_cboSprayMaterial_Click:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "cboSprayMaterial_Click", "frmRotaryAtomizer", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_cboSprayMaterial_Click
End Sub


'---------------------------------------------------------------------------
' cboSprayMaterial_DropDown:
'    Spary Material Combo dropdown event. Populate the list will all possible values.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub cboSprayMaterial_DropDown()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  With cboSprayMaterial
    .Clear
    Select Case mintAtomizerIndex
    Case 0 'AU4000
      .AddItem GetNameHKMat(0)
      .ItemData(.NewIndex) = 0
      .AddItem GetNameHKMat(3)
      .ItemData(.NewIndex) = 3
    Case 1 'AU5000
      .AddItem GetNameHKMat(0)
      .ItemData(.NewIndex) = 0
      .AddItem GetNameHKMat(1)
      .ItemData(.NewIndex) = 1
      .AddItem GetNameHKMat(2)
      .ItemData(.NewIndex) = 2
      .AddItem GetNameHKMat(3)
      .ItemData(.NewIndex) = 3
    End Select
    Dim i As Integer
    For i = 0 To .ListCount - 1
      If .ItemData(i) = mintSprayMaterialIndex Then
        .ListIndex = i
        Exit For
      End If
    Next
  End With


'====================================================
'Exit Point for cboSprayMaterial_DropDown
'====================================================
Exit_cboSprayMaterial_DropDown:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "cboSprayMaterial_DropDown", "frmRotaryAtomizer", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_cboSprayMaterial_DropDown
End Sub


'---------------------------------------------------------------------------
' cmdCancel_Click:
'    Cancel button. Set the Cancelled property, hide the form.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub cmdCancel_Click()
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  Cancelled = True
  Me.Hide


'====================================================
'Exit Point for cmdCancel_Click
'====================================================
Exit_cmdCancel_Click:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "cmdCancel_Click", "frmRotaryAtomizer", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_cmdCancel_Click
End Sub

'---------------------------------------------------------------------------
' cmdOK_Click:
'    OK button Click event. Compute the new drop distribution and hide the form.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub cmdOk_Click()
  Dim strErrLocation As String
  On Error GoTo Error_Handler

  'Compute drop distribution
  
  'Input sanity check
  If cboAtomizerName.ListIndex = -1 Then
    cboAtomizerName.SetFocus
    MsgBox "Please select an Atomizer", vbCritical
    Exit Sub
  End If
  If cboSprayMaterial.ListIndex = -1 Then
    cboSprayMaterial.SetFocus
    MsgBox "Please select a Spray Material", vbCritical
    Exit Sub
  End If
  If mintAtomizerIndex = 0 And _
     (mintSprayMaterialIndex = 1 Or mintSprayMaterialIndex = 2) Then
    cboSprayMaterial.SetFocus
    MsgBox "Please select a different Spray Material for this Atomizer.", vbCritical
    Exit Sub
  End If
  If Not IsNumeric(txtBladeAngle.Text) Then
    txtBladeAngle.SetFocus
    MsgBox "Please enter a valid Blade Angle", vbCritical
    Exit Sub
  End If
  'If bladeRPM is blank, it will be set to zero;
  'allow positive numbers or an empty field
  If (Not IsNumeric(txtBladeRPM.Text) And Trim(txtBladeRPM.Text) <> "") Or _
     (IsNumeric(txtBladeRPM.Text) And Val(txtBladeRPM.Text) <= 0) Then
    txtBladeRPM.SetFocus
    MsgBox "Please enter a valid Rotation Rate or leave blank to calculate it.", vbCritical
    Exit Sub
  End If
  
  'Set up to call calculation routine
  Dim usrHK As HKData
  Dim iunit As Long
  Dim lfl As Long
  Dim icls As Long
  Dim npts As Long
  Dim dv(500) As Single
  Dim xv(500) As Single
  Dim ier As Long
  Dim realwd(2) As Single
  Dim cdat As String * 40
  Dim clen As Long
  Dim Msg As String
  Dim iDrop As Integer
  'set up inputs
  With usrHK
    .MaxErrorLevel = 0
    .MatType = SprayMaterialIndex
    .RotType = AtomizerIndex
    .Speed = AirSpeed
    .BladeAngle = BladeAngle
    .BladeRPM = BladeRPM
    .FlowRate = FlowRate
    .SprayType = DropDistributionType
  End With
  iunit = CLng(UnitsSystem)
  lfl = 0 'Initialize loop status
  'Call the calculation routine
  Do
    agrot usrHK, iunit, lfl, icls, npts, dv(0), xv(0), ier, realwd(0), cdat, clen

    'Check for errors
    Select Case ier
    Case 0  'success, calcs are done
      'Copy output values
      DropDistributionClassification = CInt(icls)
      DropDistributionNumber = CInt(npts)
      For iDrop = 0 To DropDistributionNumber - 1
        DropDistributionDiameter(iDrop) = dv(iDrop)
        DropDistributionMassFraction(iDrop) = xv(iDrop)
      Next iDrop
      Cancelled = False
      Me.Hide
      Exit Do
    Case 1 'warning with msg and data
        Msg = "Warning!" + Chr$(13)
        Msg = Msg + "Although calculations may continue," + Chr$(13)
        Msg = Msg + Chr$(34) + Trim$(cdat) + Chr$(34) + Chr$(13)
        Msg = Msg + "is beyond recommended limits. The limits are:" + Chr$(13)
        Msg = Msg + Chr$(13)
        Msg = Msg + "Min: " + AGFormat$(realwd(1)) + Chr$(13)
        Msg = Msg + "Val: " + AGFormat$(realwd(0)) + Chr$(13)
        Msg = Msg + "Max: " + AGFormat$(realwd(2)) + Chr$(13)
        Msg = Msg + Chr$(13)
        Msg = Msg + "Continue with calculations?"
        If MsgBox(Msg, vbExclamation + vbYesNo) = vbNo Then
          Exit Do
        End If
    Case 2 'error with msg and data
        Msg = "Error!" + Chr$(13)
        Msg = Msg + "Calculations cannot continue because" + Chr$(13)
        Msg = Msg + Chr$(34) + Trim$(cdat) + Chr$(34) + Chr$(13)
        Msg = Msg + "is out of range. The limits are:" + Chr$(13)
        Msg = Msg + Chr$(13)
        Msg = Msg + "Min: " + AGFormat$(realwd(1)) + Chr$(13)
        Msg = Msg + "Val: " + AGFormat$(realwd(0)) + Chr$(13)
        Msg = Msg + "Max: " + AGFormat$(realwd(2))
        MsgBox Msg, vbCritical + vbOKOnly
      Exit Do
    Case 3 'warning with msg
        Msg = "Warning! "
        Msg = Msg & Left$(cdat, clen)
        Msg = Msg + Chr$(13)
        Msg = Msg + "Continue with calculations?"
        If MsgBox(Msg, vbExclamation + vbYesNo) = vbNo Then
          Exit Do
        End If
    Case 4 'error with msg
        Msg = "Error! "
        Msg = Msg & Left$(cdat, clen)
        MsgBox Msg, vbCritical + vbOKOnly
      Exit Do
    Case 5 'informational msg
        Msg = Left$(cdat, clen)
        MsgBox Msg, vbInformation + vbOKOnly
    Case Else 'just in case...
      MsgBox "Unexpected error level from agrot.", vbCritical
      Exit Do
    End Select
  Loop


'====================================================
'Exit Point for cmdOK_Click
'====================================================
Exit_cmdOK_Click:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "cmdOK_Click", "frmRotaryAtomizer", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_cmdOK_Click
End Sub

'---------------------------------------------------------------------------
' Form_Load:
'    Form Load event. Set up form controls.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Sub Form_Load()
  Dim strErrLocation As String
  On Error GoTo Error_Handler

  CenterForm Me
  
  mbolCancelled = True 'False only if OK clicked
  
  lblAirSpeedUnits.Caption = UnitsName(UN_SPEED)
  lblFlowRateUnits.Caption = UnitsName(UN_FLOWRATE)


'====================================================
'Exit Point for Form_Load
'====================================================
Exit_Form_Load:
  Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "Form_Load", "frmRotaryAtomizer", strErrLocation

  gobjErrors.UserMessage
  gobjErrors.WriteToErrorLog
  gobjErrors.Clear
  Resume Exit_Form_Load
End Sub

'---------------------------------------------------------------------------
' Cancelled:
'    Cancelled property Get event
'
' Return: Property Value
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Get Cancelled() As Boolean
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  Cancelled = mbolCancelled


'====================================================
'Exit Point for Cancelled
'====================================================
Exit_Cancelled:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "Cancelled", "frmRotaryAtomizer", strErrLocation
  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' Cancelled:
'    Cancelled property Set event.
'
' Arguments:
'     vNewValue: New property value.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Let Cancelled(ByVal vNewValue As Boolean)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  mbolCancelled = vNewValue


'====================================================
'Exit Point for Cancelled
'====================================================
Exit_Cancelled:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "vNewValue", IIf(vNewValue, " True", " False")
  gobjErrors.Append Err, "Cancelled", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' AtomizerIndex:
'
'
' Return:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Get AtomizerIndex() As Integer
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  AtomizerIndex = mintAtomizerIndex


'====================================================
'Exit Point for AtomizerIndex
'====================================================
Exit_AtomizerIndex:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "AtomizerIndex", "frmRotaryAtomizer", strErrLocation
  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' AtomizerIndex:
'
'
' Arguments:
'     vNewValue:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Let AtomizerIndex(ByVal vNewValue As Integer)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  mintAtomizerIndex = vNewValue
  With cboAtomizerName
    .Clear
    .AddItem GetNameHKNoz(mintAtomizerIndex)
    .ItemData(.NewIndex) = mintAtomizerIndex
    .ListIndex = 0
  End With


'====================================================
'Exit Point for AtomizerIndex
'====================================================
Exit_AtomizerIndex:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "vNewValue", CStr(vNewValue)
  gobjErrors.Append Err, "AtomizerIndex", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' SprayMaterialIndex:
'
'
' Return:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Get SprayMaterialIndex() As Integer
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  SprayMaterialIndex = mintSprayMaterialIndex


'====================================================
'Exit Point for SprayMaterialIndex
'====================================================
Exit_SprayMaterialIndex:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "SprayMaterialIndex", "frmRotaryAtomizer", strErrLocation
  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' SprayMaterialIndex:
'
'
' Arguments:
'     vNewValue:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Let SprayMaterialIndex(ByVal vNewValue As Integer)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  mintSprayMaterialIndex = vNewValue
  With cboSprayMaterial
    .Clear
    .AddItem GetNameHKMat(vNewValue)
    .ItemData(.NewIndex) = mintSprayMaterialIndex
    .ListIndex = 0
  End With


'====================================================
'Exit Point for SprayMaterialIndex
'====================================================
Exit_SprayMaterialIndex:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "vNewValue", CStr(vNewValue)
  gobjErrors.Append Err, "SprayMaterialIndex", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' AirSpeed:
'
'
' Return:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Get AirSpeed() As Single
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If IsNumeric(txtAirSpeed.Text) Then
    AirSpeed = UnitsInternal(CSng(txtAirSpeed.Text), UN_SPEED)
  Else
    AirSpeed = 0
  End If


'====================================================
'Exit Point for AirSpeed
'====================================================
Exit_AirSpeed:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "AirSpeed", "frmRotaryAtomizer", strErrLocation
  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' AirSpeed:
'
'
' Arguments:
'     vNewValue:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Let AirSpeed(ByVal vNewValue As Single)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  txtAirSpeed.Text = AGFormat(UnitsDisplay(vNewValue, UN_SPEED))


'====================================================
'Exit Point for AirSpeed
'====================================================
Exit_AirSpeed:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "vNewValue", CStr(vNewValue)
  gobjErrors.Append Err, "AirSpeed", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' BladeAngle:
'
'
' Return:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Get BladeAngle() As Single
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If IsNumeric(txtBladeAngle.Text) Then
    BladeAngle = CSng(txtBladeAngle.Text)
  Else
    BladeAngle = 0
  End If


'====================================================
'Exit Point for BladeAngle
'====================================================
Exit_BladeAngle:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "BladeAngle", "frmRotaryAtomizer", strErrLocation
  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' BladeAngle:
'
'
' Arguments:
'     vNewValue:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Let BladeAngle(ByVal vNewValue As Single)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  txtBladeAngle.Text = AGFormat(vNewValue)


'====================================================
'Exit Point for BladeAngle
'====================================================
Exit_BladeAngle:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "vNewValue", CStr(vNewValue)
  gobjErrors.Append Err, "BladeAngle", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' BladeRPM:
'
'
' Return:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Get BladeRPM() As Single
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If IsNumeric(txtBladeRPM.Text) Then
    BladeRPM = CSng(txtBladeRPM.Text)
  Else
    BladeRPM = 0
  End If


'====================================================
'Exit Point for BladeRPM
'====================================================
Exit_BladeRPM:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "BladeRPM", "frmRotaryAtomizer", strErrLocation
  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' BladeRPM:
'
'
' Arguments:
'     vNewValue:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Let BladeRPM(ByVal vNewValue As Single)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  'If BladeRPM is 0, the DLL calculates it. In this form,
  'The user leaves BladeRPM blank for this to happen.
  If vNewValue <= 0 Then
    txtBladeRPM.Text = ""
  Else
    txtBladeRPM.Text = AGFormat(vNewValue)
  End If


'====================================================
'Exit Point for BladeRPM
'====================================================
Exit_BladeRPM:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "vNewValue", CStr(vNewValue)
  gobjErrors.Append Err, "BladeRPM", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' Flowrate:
'
'
' Return:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Get FlowRate() As Single
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  If IsNumeric(txtFlowRate.Text) Then
    FlowRate = UnitsInternal(CSng(txtFlowRate.Text), UN_FLOWRATE)
  Else
    FlowRate = 0
  End If


'====================================================
'Exit Point for Flowrate
'====================================================
Exit_Flowrate:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "Flowrate", "frmRotaryAtomizer", strErrLocation
  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' Flowrate:
'
'
' Arguments:
'     vNewValue:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Let FlowRate(ByVal vNewValue As Single)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  txtFlowRate.Text = AGFormat(UnitsDisplay(vNewValue, UN_FLOWRATE))


'====================================================
'Exit Point for Flowrate
'====================================================
Exit_Flowrate:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "vNewValue", CStr(vNewValue)
  gobjErrors.Append Err, "Flowrate", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' DropDistributionType:
'
'
' Return:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Get DropDistributionType() As Integer
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  Dim optItem As OptionButton
  For Each optItem In optDropDistributionType()
    If optItem.Value Then
      DropDistributionType = optItem.Index
      Exit Property
    End If
  Next optItem
  DropDistributionType = -1


'====================================================
'Exit Point for DropDistributionType
'====================================================
Exit_DropDistributionType:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "DropDistributionType", "frmRotaryAtomizer", strErrLocation
  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' DropDistributionType:
'
'
' Arguments:
'     vNewValue:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Let DropDistributionType(ByVal vNewValue As Integer)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  optDropDistributionType(vNewValue).Value = True


'====================================================
'Exit Point for DropDistributionType
'====================================================
Exit_DropDistributionType:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "vNewValue", CStr(vNewValue)
  gobjErrors.Append Err, "DropDistributionType", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' DropDistributionClassification:
'
'
' Return:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Get DropDistributionClassification() As Integer
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  DropDistributionClassification = mintDropDistributionClassification


'====================================================
'Exit Point for DropDistributionClassification
'====================================================
Exit_DropDistributionClassification:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "DropDistributionClassification", "frmRotaryAtomizer", strErrLocation
  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' DropDistributionClassification:
'
'
' Arguments:
'     vNewValue:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Property Let DropDistributionClassification(ByVal vNewValue As Integer)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  mintDropDistributionClassification = vNewValue


'====================================================
'Exit Point for DropDistributionClassification
'====================================================
Exit_DropDistributionClassification:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "vNewValue", CStr(vNewValue)
  gobjErrors.Append Err, "DropDistributionClassification", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' DropDistributionNumber:
'
'
' Return:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Get DropDistributionNumber() As Integer
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  DropDistributionNumber = mintDropDistributionNumber


'====================================================
'Exit Point for DropDistributionNumber
'====================================================
Exit_DropDistributionNumber:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  gobjErrors.Append Err, "DropDistributionNumber", "frmRotaryAtomizer", strErrLocation
  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' DropDistributionNumber:
'
'
' Arguments:
'     vNewValue:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Property Let DropDistributionNumber(ByVal vNewValue As Integer)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  mintDropDistributionNumber = vNewValue
  If mintDropDistributionNumber > 0 Then
    ReDim msglDropDistributionDiameter(mintDropDistributionNumber - 1)
    ReDim msglDropDistributionMassFraction(mintDropDistributionNumber - 1)
  Else
    ReDim msglDropDistributionDiameter(0)
    ReDim msglDropDistributionMassFraction(0)
  End If


'====================================================
'Exit Point for DropDistributionNumber
'====================================================
Exit_DropDistributionNumber:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "vNewValue", CStr(vNewValue)
  gobjErrors.Append Err, "DropDistributionNumber", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' DropDistributionDiameter:
'
'
' Arguments:
'     Index:
'
' Return:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Get DropDistributionDiameter(Index As Integer) As Single
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  DropDistributionDiameter = msglDropDistributionDiameter(Index)


'====================================================
'Exit Point for DropDistributionDiameter
'====================================================
Exit_DropDistributionDiameter:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "Index", CStr(Index)
  gobjErrors.Append Err, "DropDistributionDiameter", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' DropDistributionDiameter:
'
'
' Arguments:
'     Index:
'     vNewValue:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Private Property Let DropDistributionDiameter(Index As Integer, ByVal vNewValue As Single)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  msglDropDistributionDiameter(Index) = vNewValue


'====================================================
'Exit Point for DropDistributionDiameter
'====================================================
Exit_DropDistributionDiameter:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "Index", CStr(Index)
  dicCommandLineArgs.Add "vNewValue", CStr(vNewValue)
  gobjErrors.Append Err, "DropDistributionDiameter", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' DropDistributionMassFraction:
'
'
' Arguments:
'     Index:
'
' Return:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Get DropDistributionMassFraction(Index As Integer) As Single
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  DropDistributionMassFraction = msglDropDistributionMassFraction(Index)


'====================================================
'Exit Point for DropDistributionMassFraction
'====================================================
Exit_DropDistributionMassFraction:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "Index", CStr(Index)
  gobjErrors.Append Err, "DropDistributionMassFraction", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

'---------------------------------------------------------------------------
' DropDistributionMassFraction:
'
'
' Arguments:
'     Index:
'     vNewValue:
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2005-05-09  TBC  Added error handling
'
'---------------------------------------------------------------------------
Public Property Let DropDistributionMassFraction(Index As Integer, ByVal vNewValue As Single)
  Dim strErrLocation As String
  On Error GoTo Error_Handler
  
  msglDropDistributionMassFraction(Index) = vNewValue


'====================================================
'Exit Point for DropDistributionMassFraction
'====================================================
Exit_DropDistributionMassFraction:
  Exit Property


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
  Dim dicCommandLineArgs As Dictionary
  Set dicCommandLineArgs = New Dictionary
  dicCommandLineArgs.CompareMode = TextCompare
  dicCommandLineArgs.Add "Index", CStr(Index)
  dicCommandLineArgs.Add "vNewValue", CStr(vNewValue)
  gobjErrors.Append Err, "DropDistributionMassFraction", "frmRotaryAtomizer", strErrLocation, dicCommandLineArgs
  Set dicCommandLineArgs = Nothing

  gobjErrors.RaiseError
End Property

