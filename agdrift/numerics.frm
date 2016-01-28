VERSION 5.00
Begin VB.Form frmNumerics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Numerical Values"
   ClientHeight    =   6540
   ClientLeft      =   5100
   ClientTop       =   3060
   ClientWidth     =   4935
   ForeColor       =   &H80000008&
   HelpContextID   =   1200
   Icon            =   "NUMERICS.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6540
   ScaleWidth      =   4935
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      HelpContextID   =   1200
      Left            =   2280
      TabIndex        =   2
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      HelpContextID   =   1200
      Left            =   3120
      TabIndex        =   1
      Top             =   6120
      Width           =   735
   End
   Begin VB.Frame fraDSD 
      Caption         =   "Drop Size Distribution"
      Height          =   2535
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox cboDSD 
         Height          =   315
         HelpContextID   =   1200
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblSpan 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblSpanName 
         Alignment       =   1  'Right Justify
         Caption         =   "Relative Span:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   1680
      End
      Begin VB.Label lblF141Name 
         Alignment       =   1  'Right Justify
         Caption         =   "< 141 µm:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1680
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "V0.9"
         Height          =   195
         Left            =   1410
         TabIndex        =   11
         Top             =   1485
         Width           =   330
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "D         :"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1410
         Width           =   1680
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "V0.5"
         Height          =   195
         Left            =   1395
         TabIndex        =   20
         Top             =   1140
         Width           =   330
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "D         :"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1065
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "V0.1"
         Height          =   195
         Left            =   1410
         TabIndex        =   6
         Top             =   780
         Width           =   330
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "D         :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   705
         Width           =   1680
      End
      Begin VB.Label lblF141Units 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   22
         Top             =   2160
         Width           =   120
      End
      Begin VB.Label lblF141 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblD10 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblVMD 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblD90 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblD10Units 
         AutoSize        =   -1  'True
         Caption         =   "µm"
         Height          =   195
         Left            =   3240
         TabIndex        =   16
         Top             =   720
         Width           =   210
      End
      Begin VB.Label lblVMDUnits 
         AutoSize        =   -1  'True
         Caption         =   "µm"
         Height          =   195
         Left            =   3240
         TabIndex        =   17
         Top             =   1080
         Width           =   210
      End
      Begin VB.Label lblD90Units 
         AutoSize        =   -1  'True
         Caption         =   "µm"
         Height          =   195
         Left            =   3240
         TabIndex        =   19
         Top             =   1440
         Width           =   210
      End
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1200
      Left            =   3960
      TabIndex        =   0
      Top             =   6120
      Width           =   855
   End
   Begin VB.Frame fraDepos 
      Caption         =   "Deposition"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   4695
      Begin VB.Label lblMeanDepName 
         Alignment       =   1  'Right Justify
         Caption         =   "Mean Deposition:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label lblMeanDep 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblCOVName 
         Alignment       =   1  'Right Justify
         Caption         =   "COV:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label lblCOV 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblSwathDispUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lblSwathDisp 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblSwathDispName 
         Alignment       =   1  'Right Justify
         Caption         =   "Swath Displacement:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1680
      End
   End
   Begin VB.Frame fraAccount 
      Caption         =   "Accountancy of Active"
      Height          =   2055
      Left            =   120
      TabIndex        =   24
      Top             =   3960
      Width           =   4695
      Begin VB.Label lblCanDepName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Canopy Deposition:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label lblCanDep 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblCanDepUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   43
         Top             =   240
         Width           =   150
      End
      Begin VB.Label lblEvapFracUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   42
         Top             =   1680
         Width           =   150
      End
      Begin VB.Label lblDownwindDep 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1905
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblDownwindDepName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Downwind Deposition:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label lblDownwindDepUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   39
         Top             =   960
         Width           =   150
      End
      Begin VB.Label lblAirborneDriftUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   38
         Top             =   1320
         Width           =   150
      End
      Begin VB.Label lblAppEffUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   3240
         TabIndex        =   37
         Top             =   600
         Width           =   150
      End
      Begin VB.Label lblAppEff 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   36
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblAppEffName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Application Efficiency:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label lblEvapFracName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carrier Evaporated:"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   1680
      End
      Begin VB.Label lblEvapFrac 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   33
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblAirborneDriftName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Airborne Drift:"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   1680
      End
      Begin VB.Label lblAirborneDrift 
         Alignment       =   1  'Right Justify
         Caption         =   "00000.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   31
         Top             =   1320
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmNumerics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: numerics.frm,v 1.10 2006/11/08 15:18:11 tom Exp $

Private Sub cboDSD_Click()
  UpdateDSDStats
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdPrint_Click()
'print the current UserData
  Dim BeginPage As Integer
  Dim EndPage As Integer
  Dim NumCopies As Integer
  Dim ReportText As String
  Dim i As Integer
  Dim pages As Integer

  If PrinterExists() Then
    If PrintDialog(BeginPage, EndPage, NumCopies) Then
      ReportText = GenFormData()
      For i = 1 To NumCopies
        PrintData ReportText, False, pages, Mag
      Next
    End If
  End If
End Sub

Private Sub cmdSave_Click()
  Dim fn As String
  
  If Not FileDialog(FD_SAVEAS, FD_TYPE_TEXT, fn) Then
    Exit Sub
  End If
  On Error GoTo ErrHandcmdSave
  Open fn For Output As #1
  Print #1, GenFormData()
  Close #1
ExitcmdSave:
  Exit Sub
  
ErrHandcmdSave:
  MsgBox "Error writing file: " + fn + vbCr + Error$(Err)
  Resume ExitcmdSave
End Sub

Private Sub DataToForm()
'Place numerics data in form controls
  
  'DSD
  'always available
  '
  UpdateDSDStats
  
  'Deposition
  'not available for:
  '  - Tier I
  '  - Tier II/III if calcs are not done
  fraDepos.Enabled = True
  lblSwathDispName.Enabled = True
  lblSwathDisp.Enabled = True
  lblSwathDisp.Caption = ""
  lblSwathDispUnits.Caption = ""
  If (UD.Tier = TIER_1) Then
    fraDepos.Enabled = False
    lblSwathDispName.Enabled = False
    lblSwathDisp.Enabled = False
    lblSwathDisp.Caption = "n/a"
  ElseIf Not UC.Valid Then
    lblSwathDisp.Enabled = False
    lblSwathDisp.Caption = "need calcs"
  ElseIf UC.SwathDisp = -1 Then
    lblSwathDisp.Enabled = False
    lblSwathDisp.Caption = "out of range!"
  Else
    lblSwathDisp.Caption = AGFormat$(UnitsDisplay(UC.SwathDisp, UN_LENGTH))
    lblSwathDispUnits.Caption = UnitsName(UN_LENGTH)
  End If
  
  'Accountancy
  'not available for:
  '  - Tier I/II
  '  - Tier III - if calcs are not done
  '             - AirborneDrift/AppEff n/a if
  '               SwathDispType is Fraction of Applied
  lblCOVName.Enabled = True
  lblCOV.Enabled = True
  lblMeanDepName.Enabled = True
  lblMeanDep.Enabled = True
  fraAccount.Enabled = True
  lblCanDepName.Enabled = True
  lblCanDepUnits.Enabled = True
  lblAppEffName.Enabled = True
  lblAppEffUnits.Enabled = True
  lblDownwindDepName.Enabled = True
  lblDownwindDepUnits.Enabled = True
  lblAirborneDriftName.Enabled = True
  lblAirborneDriftUnits.Enabled = True
  lblEvapFracName.Enabled = True
  lblEvapFracUnits.Enabled = True
  
  lblCanDep.Enabled = True
  lblCanDep.Caption = ""
  lblAppEff.Enabled = True
  lblAppEff.Caption = ""
  lblDownwindDep.Enabled = True
  lblDownwindDep.Caption = ""
  lblAirborneDrift.Enabled = True
  lblAirborneDrift.Caption = ""
  lblEvapFrac.Enabled = True
  lblEvapFrac.Caption = ""
  Select Case UD.Tier
  Case TIER_1, TIER_2 'tier 1 or 2
    fraAccount.Enabled = False
    lblCanDepName.Enabled = False
    lblCanDepUnits.Enabled = False
    lblAppEffName.Enabled = False
    lblAppEffUnits.Enabled = False
    lblDownwindDepName.Enabled = False
    lblDownwindDepUnits.Enabled = False
    lblAirborneDriftName.Enabled = False
    lblAirborneDriftUnits.Enabled = False
    lblEvapFracName.Enabled = False
    lblEvapFracUnits.Enabled = False
    
    lblCanDep.Enabled = False
    lblCanDep.Caption = "n/a"
    lblAppEff.Enabled = False
    lblAppEff.Caption = "n/a"
    lblDownwindDep.Enabled = False
    lblDownwindDep.Caption = "n/a"
    lblAirborneDrift.Enabled = False
    lblAirborneDrift.Caption = "n/a"
    lblEvapFrac.Enabled = False
    lblEvapFrac.Caption = "n/a"
    lblCOVName.Enabled = False
    lblCOV.Enabled = False
    lblCOV.Caption = "n/a"
    lblMeanDepName.Enabled = False
    lblMeanDep.Enabled = False
    lblMeanDep.Caption = "n/a"
  Case TIER_3 'tier 3
    'COV/Mean Depos
    If Not UC.Valid Then
      lblCOV.Enabled = False
      lblCOV.Caption = "need calcs"
      lblMeanDep.Enabled = False
      lblMeanDep.Caption = "need calcs"
    ElseIf UC.SwathDisp = -1 Then
      lblCOV.Enabled = False
      lblCOV.Caption = "out of range!"
      lblMeanDep.Enabled = False
      lblMeanDep.Caption = "out of range!"
    Else
      lblCOV.Caption = AGFormat$(UC.SBCOV)
      lblMeanDep.Caption = AGFormat$(UC.SBMeanDep)
    End If
    'Canopy Deposition
    If UD.CTL.SwathDispType = 1 Then 'Fraction of applied
      lblCanDep.Enabled = False
      lblCanDep.Caption = "n/a"
    ElseIf Not UC.Valid Then
      lblCanDep.Enabled = False
      lblCanDep.Caption = "need calcs"
    ElseIf UC.SwathDisp = -1 Then
      lblCanDep.Enabled = False
      lblCanDep.Caption = "out of range!"
    Else
      lblCanDep.Caption = AGFormat$(UC.CanopyDep)
    End If
    'Airborne Drift
    If Not UC.Valid Then
      lblAirborneDrift.Enabled = False
      lblAirborneDrift.Caption = "need calcs"
    ElseIf UC.SwathDisp = -1 Then
      lblAirborneDrift.Enabled = False
      lblAirborneDrift.Caption = "out of range!"
    Else
      lblAirborneDrift.Caption = AGFormat$(UC.AirborneDrift)
    End If
    'Downwind Deposition
    If UD.CTL.SwathDispType = 1 Then 'Fraction of applied
      lblDownwindDep.Enabled = False
      lblDownwindDep.Caption = "n/a"
    ElseIf Not UC.Valid Then
      lblDownwindDep.Enabled = False
      lblDownwindDep.Caption = "need calcs"
    ElseIf UC.SwathDisp = -1 Then
      lblDownwindDep.Enabled = False
      lblDownwindDep.Caption = "out of range!"
    Else
      lblDownwindDep.Caption = AGFormat$(UC.DownwindDep)
    End If
    'Carrier Evaporated
    If Not UC.Valid Then
      lblEvapFrac.Enabled = False
      lblEvapFrac.Caption = "need calcs"
    ElseIf UC.SwathDisp = -1 Then
      lblEvapFrac.Enabled = False
      lblEvapFrac.Caption = "out of range!"
    Else
      lblEvapFrac.Caption = AGFormat$(UC.EvapFrac)
    End If
    'Application Efficiency
    If UD.CTL.SwathDispType = 1 Then 'Fraction of applied
      lblAppEff.Enabled = False
      lblAppEff.Caption = "n/a"
    ElseIf Not UC.Valid Then
      lblAppEff.Enabled = False
      lblAppEff.Caption = "need calcs"
    ElseIf UC.SwathDisp = -1 Then
      lblAppEff.Enabled = False
      lblAppEff.Caption = "out of range!"
    Else
      lblAppEff.Caption = AGFormat$(UC.AppEff)
    End If
  End Select
End Sub

Private Sub UpdateDSDStats()
'Update the values of the Dropsize parameters based on the DSD combo setting
  Dim PlotVar As Long 'var name is selected for consistency
  Dim D10 As Single
  Dim VMD As Single
  Dim D90 As Single
  Dim Span As Single
  Dim F141 As Single
  Dim DP As Single
  Dim adum As Single
  Dim s As String
  Dim state As Boolean
    
  'Default Stats
  D10 = 0
  VMD = 0
  D90 = 0
  Span = 0
  F141 = 0
  
  'Extract PlotVar from DSD combo
  PlotVar = cboDSD.ItemData(cboDSD.ListIndex)
  
  'Generate the stats from the selected DSD
  Select Case PlotVar
  Case PV_VFINC0
    Call agdsrn(0, CLng(UD.DSD(0).NumDrop), UD.DSD(0).Diam(0), UD.DSD(0).MassFrac(0), _
                VMD, Span, D10, D90, F141, DP)
  Case PV_VFINC1
    Call agdsrn(0, CLng(UD.DSD(1).NumDrop), UD.DSD(1).Diam(0), UD.DSD(1).MassFrac(0), _
                VMD, Span, D10, D90, F141, DP)
  Case PV_VFINC2
    Call agdsrn(0, CLng(UD.DSD(2).NumDrop), UD.DSD(2).Diam(0), UD.DSD(2).MassFrac(0), _
                VMD, Span, D10, D90, F141, DP)
  Case PV_DWDSDINC
    If UC.Valid Then Call agdsrn(0, CLng(UC.NumDWDSD), UC.DWDSDDiam(0), UC.DWDSDFrac(0), _
                VMD, Span, D10, D90, F141, DP)
  Case PV_FXDSDINC
    If UC.Valid Then Call agdsrn(0, CLng(UC.NumFXDSD), UC.FXDSDDiam(0), UC.FXDSDFrac(0), _
                VMD, Span, D10, D90, F141, DP)
  Case PV_SBDSDINC
    If UC.Valid Then Call agdsrn(0, CLng(UC.NumSBDSD), UC.SBDSDDiam(0), UC.SBDSDFrac(0), _
                VMD, Span, D10, D90, F141, DP)
  Case PV_CNDSDINC
    If UC.Valid Then Call agdsrn(0, CLng(UC.NumCNDSD), UC.CNDSDDiam(0), UC.CNDSDFrac(0), _
                VMD, Span, D10, D90, F141, DP)
  End Select
  
  'set labels
  If Not UC.Valid And ((PlotVar And PVA_SOURCE_MASK) = PVA_UC) Then 'Calcs required
    s = "need calcs"
    state = True
    lblD10.Caption = s: lblD10.Enabled = False
    lblVMD.Caption = s: lblVMD.Enabled = False
    lblD90.Caption = s: lblD90.Enabled = False
    lblSpan.Caption = s: lblSpan.Enabled = False
    lblF141.Caption = s: lblF141.Enabled = False
  Else
    lblD10.Caption = AGFormat$(D10): lblD10.Enabled = True
    lblVMD.Caption = AGFormat$(VMD): lblVMD.Enabled = True
    lblD90.Caption = AGFormat$(D90): lblD90.Enabled = True
    lblSpan.Caption = AGFormat$(Span): lblSpan.Enabled = True
    lblF141.Caption = AGFormat$(F141): lblF141.Enabled = True
  End If
End Sub

Private Sub Form_Load()
'initialize this form
  Dim c As Control
  
  'Center the form on the screen
  CenterForm Me
  
  'Load up the DSD combo
  'Put the relevant PlotVar into the ItemData as an identifier
  With cboDSD
    If UD.Tier > TIER_2 Then
      If PlotIsAvailableExtended(PV_VFINC0) Then .AddItem "Initial DSD 1": .ItemData(.NewIndex) = PV_VFINC0
      If PlotIsAvailableExtended(PV_VFINC1) Then .AddItem "Initial DSD 2": .ItemData(.NewIndex) = PV_VFINC1
      If PlotIsAvailableExtended(PV_VFINC2) Then .AddItem "Initial DSD 3": .ItemData(.NewIndex) = PV_VFINC2
    Else
      'For Tier 1,2 only one of these should be available at one time
      If PlotIsAvailableExtended(PV_VFINC0) Then .AddItem "Initial DSD": .ItemData(.NewIndex) = PV_VFINC0
      If PlotIsAvailableExtended(PV_VFINC1) Then .AddItem "Initial DSD": .ItemData(.NewIndex) = PV_VFINC1
      If PlotIsAvailableExtended(PV_VFINC2) Then .AddItem "Initial DSD": .ItemData(.NewIndex) = PV_VFINC2
    End If
    If PlotIsAvailable(PV_DWDSDINC) Then .AddItem "Downwind": .ItemData(.NewIndex) = PV_DWDSDINC
    If PlotIsAvailable(PV_FXDSDINC) Then .AddItem "Vertical Profile": .ItemData(.NewIndex) = PV_FXDSDINC
    If PlotIsAvailable(PV_SBDSDINC) Then .AddItem "Spray Block": .ItemData(.NewIndex) = PV_SBDSDINC
    If PlotIsAvailable(PV_CNDSDINC) Then .AddItem "Canopy": .ItemData(.NewIndex) = PV_CNDSDINC
    .ListIndex = 0
  End With
  
  'For other than Tier 3 FS, remove the Canopy Deposition line
  If Not (UD.Tier = TIER_3 And UD.Smokey = AUD_FS) Then
    lblCanDepName.Visible = False
    lblCanDep.Visible = False
    lblCanDepUnits.Visible = False
    'create a collection of controls to Modify
    Dim collect As New Collection
    collect.Add lblAppEffName
    collect.Add lblAppEff
    collect.Add lblAppEffUnits
    collect.Add lblDownwindDepName
    collect.Add lblDownwindDep
    collect.Add lblDownwindDepUnits
    collect.Add lblAirborneDriftName
    collect.Add lblAirborneDrift
    collect.Add lblAirborneDriftUnits
    collect.Add lblEvapFracName
    collect.Add lblEvapFrac
    collect.Add lblEvapFracUnits
    'Now shift everything in the collection up
    For Each c In collect
      c.Top = c.Top - 360
    Next
    'Make other adjustments too
    fraAccount.Height = fraAccount.Height - 360
    cmdSave.Top = cmdSave.Top - 360
    cmdPrint.Top = cmdPrint.Top - 360
    cmdOK.Top = cmdOK.Top - 360
    Me.Height = Me.Height - 360
  End If
  
  'display the numbers on the form
  DataToForm
End Sub

Private Function GenFormData() As String
'Generate report text for this form to be used for printing
  
  Dim gfd As String  'temporary storage for report text
  Dim s As String        'workspace string

  gfd = "" 'start with a blank string
  
  AppendStr gfd, "AgDRIFT® Numerical Values", True
  AppendStr gfd, "", True
  
  AppendStr gfd, fraDSD.Caption & ":", True
  AppendStr gfd, cboDSD.List(cboDSD.ListIndex), True
  AppendStr gfd, "Dv0.1 " & lblD10.Caption & " " & lblD10Units.Caption, True
  AppendStr gfd, "Dv0.5 " & lblVMD.Caption & " " & lblVMDUnits.Caption, True
  AppendStr gfd, "Dv0.9 " & lblD90.Caption & " " & lblD90Units.Caption, True
  AppendStr gfd, lblSpanName.Caption & " " & lblSpan.Caption, True
  AppendStr gfd, lblF141Name.Caption & " " & lblF141.Caption & " " & lblF141Units.Caption, True
  AppendStr gfd, "", True

  AppendStr gfd, fraDepos.Caption & ":", True
  AppendStr gfd, lblSwathDispName.Caption & " " & lblSwathDisp.Caption & " " & lblSwathDispUnits.Caption, True
  AppendStr gfd, "", True

  AppendStr gfd, fraAccount.Caption & ":", True
  AppendStr gfd, lblAppEffName.Caption & " " & lblAppEff.Caption & " " & lblAppEffUnits.Caption, True
  AppendStr gfd, lblDownwindDepName.Caption & " " & lblDownwindDep.Caption & " " & lblDownwindDepUnits.Caption, True
  AppendStr gfd, lblAirborneDriftName.Caption & " " & lblAirborneDrift.Caption & " " & lblAirborneDriftUnits.Caption, True
  AppendStr gfd, lblEvapFracName.Caption & " " & lblEvapFrac.Caption & " " & lblEvapFracUnits.Caption, True
  AppendStr gfd, "", True
  
  AppendStr gfd, "Tier: " & String$(UD.Tier, "I"), True
  AppendStr gfd, "RunID:", True
  AppendStr gfd, "  " & GetRunID(), True
  AppendStr gfd, "", True
  
  GenFormData = gfd
End Function

