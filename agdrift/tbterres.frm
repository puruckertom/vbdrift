VERSION 5.00
Begin VB.Form frmTBTerrestrial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Terrestrial Assessment"
   ClientHeight    =   5355
   ClientLeft      =   3345
   ClientTop       =   1560
   ClientWidth     =   5955
   ForeColor       =   &H80000008&
   HelpContextID   =   1456
   Icon            =   "TBTERRES.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5355
   ScaleWidth      =   5955
   Begin VB.Frame fraPond 
      Caption         =   "Terrestrial Field Definition"
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   5775
      Begin VB.TextBox txtDist 
         Height          =   285
         HelpContextID   =   1506
         Left            =   3240
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optBodyType 
         Caption         =   "User-defined Area Average"
         Height          =   255
         HelpContextID   =   1456
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton optBodyType 
         Caption         =   "Point Deposition"
         Height          =   255
         HelpContextID   =   1456
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblDist 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Downwind Width of Area Average:"
         Height          =   195
         Left            =   630
         TabIndex        =   22
         Top             =   1005
         Width           =   2460
      End
      Begin VB.Label lblDistUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   4170
         TabIndex        =   23
         Top             =   1005
         Width           =   420
      End
   End
   Begin VB.Frame fraTier1 
      Caption         =   "Tier I Settings"
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   5775
      Begin VB.TextBox txtActiveRate 
         Height          =   285
         HelpContextID   =   1010
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label lblActiveRateUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3600
         TabIndex        =   26
         Top             =   285
         Width           =   420
      End
      Begin VB.Label lblActiveRate 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Active Rate:"
         Height          =   195
         Left            =   1320
         TabIndex        =   25
         Top             =   285
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      HelpContextID   =   1456
      Left            =   2520
      TabIndex        =   3
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "Plo&t"
      Height          =   375
      HelpContextID   =   1456
      Left            =   1680
      TabIndex        =   4
      Top             =   4920
      Width           =   735
   End
   Begin VB.Frame fraCalc 
      Caption         =   "Calculations"
      Height          =   2535
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   5775
      Begin VB.TextBox txtCalc 
         Height          =   285
         HelpContextID   =   1509
         Index           =   4
         Left            =   2400
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         HelpContextID   =   1509
         Index           =   3
         Left            =   4080
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         HelpContextID   =   1509
         Index           =   2
         Left            =   2400
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         HelpContextID   =   1509
         Index           =   1
         Left            =   2400
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         HelpContextID   =   1508
         Index           =   0
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblUnits4 
         AutoSize        =   -1  'True
         Caption         =   "mg/cm²"
         Height          =   195
         Left            =   3600
         TabIndex        =   27
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label lblUnits1 
         Caption         =   "Fraction of Applied"
         Height          =   210
         Left            =   3600
         TabIndex        =   16
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label lblCalcDistUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3600
         TabIndex        =   17
         Top             =   390
         Width           =   420
      End
      Begin VB.Label lblResults1 
         Alignment       =   2  'Center
         Caption         =   "Distance To Point or Area Average From Edge of Application Area:"
         Height          =   675
         Left            =   120
         TabIndex        =   18
         Top             =   285
         Width           =   2175
      End
      Begin VB.Label lblUnits3 
         AutoSize        =   -1  'True
         Caption         =   "lb/ac"
         Height          =   195
         Left            =   5280
         TabIndex        =   19
         Top             =   1350
         Width           =   375
      End
      Begin VB.Label lblResults2 
         Alignment       =   2  'Center
         Caption         =   "Initial Average Deposition:"
         Height          =   495
         Left            =   480
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblUnits2 
         AutoSize        =   -1  'True
         Caption         =   "g/ha"
         Height          =   195
         Left            =   3600
         TabIndex        =   21
         Top             =   1350
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      HelpContextID   =   1456
      Left            =   3360
      TabIndex        =   2
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calc"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1456
      Left            =   4200
      TabIndex        =   1
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Clos&e"
      Height          =   375
      HelpContextID   =   1456
      Left            =   5040
      TabIndex        =   0
      Top             =   4920
      Width           =   855
   End
End
Attribute VB_Name = "frmTBTerrestrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tbterres.frm,v 1.12 2003/08/22 19:49:05 tom Exp $
'
'This form requires its own copy of the calculated Deposition and
'Pond-Integrated Deposition arrays to work. To use this form, Load
'it, call frmTBAquatic.LoadDeposition() with the depositioins that
'you want to use (i.e. UC.etc, or that of the MAA form)

Option Explicit

Dim PropTakeAction As Integer
Dim BodyType As Integer         'Tracks body type selection
Dim NeedCalcs As Integer        'tracks calculation status
Dim CalcOutputMarker As Integer 'Tracks user-selected input box
Dim PreviousUnits As Integer    'Tracks units setting

'Local copy of Depositiion and Pond-integrated deposition
'This data set may come from the UC area or from elsewhere,
'i.e.the MAA toolbox
Dim NumDep As Long
Dim DepDist() As Single
Dim DepVal() As Single
Dim NumPID As Long
Dim PIDDist() As Single
Dim PIDVal() As Single

Public Sub LoadDeposition(numd As Integer, depd() As Single, depv() As Single, _
                          nump As Integer, pidd() As Single, pidv() As Single)
'Initialize form-local copy of depositions
  NumDep = numd
  ReDim DepDist(NumDep - 1)
  ReDim DepVal(NumDep - 1)
  CopyMemory DepDist(0), depd(0), NumDep * Len(DepDist(0))
  CopyMemory DepVal(0), depv(0), NumDep * Len(DepDist(0))
  NumPID = nump
  ReDim PIDDist(NumPID - 1)
  ReDim PIDVal(NumPID - 1)
  CopyMemory PIDDist(0), pidd(0), NumPID * Len(PIDDist(0))
  CopyMemory PIDVal(0), pidv(0), NumPID * Len(PIDDist(0))
End Sub

Private Sub Calculate()
'Calculate the Toxicity (exposure)
  Dim ISTYPE As Long
  Dim INTYPE As Long
  Dim XLENG As Single
  Dim XACT As Single
  Dim XLAND As Single
  Dim XAPPL As Single
  Dim XDEPS As Single
  Dim XDEPD As Single
  Dim XCONC As Single

  Dim Msg As String
  Dim NPlong As Long
  Dim i As Integer

  'Check inputs
  If CalcOutputMarker < 0 Then  'no calc spec
    Msg = "Enter a value in one of the boxes in the "
    Msg = Msg & "calculation frame and press Calc."
    MsgBox Msg, vbOKOnly + vbInformation
    Exit Sub
  End If

  ' Change the form mouse pointer
  Me.MousePointer = vbHourglass 'hourglass
  
  'Extract the input data from the form controls
  ISTYPE = CLng(BodyType) 'This var is maintained by the form
  INTYPE = CLng(CalcOutputMarker) 'This var is maintained by the form
  XLENG = UnitsInternal(Val(txtDist.Text), UN_LENGTH)
  XACT = UnitsInternal(Val(txtActiveRate.Text) / (UD.SM.FlowRate * UD.SM.NonVGrav), UN_RATEMASS)
  XLAND = UnitsInternal(Val(txtCalc(0).Text), UN_LENGTH)
  XAPPL = Val(txtCalc(1).Text)
  XDEPS = Val(txtCalc(2).Text)
  XDEPD = Val(txtCalc(3).Text)
  XCONC = Val(txtCalc(4).Text)

  'Set up the global area to store the calculated results
  TPD.X1D = True 'X data is one-dimensional
  ReDim TPD.np(0)
  ReDim TPD.X(MAX_CALCDATA)
  ReDim TPD.Y(MAX_CALCDATA, 0)
  
  Call agterr(UD, _
             NumDep, DepDist(0), DepVal(0), _
             NumPID, PIDDist(0), PIDVal(0), _
             ISTYPE, INTYPE, XLENG, XACT, XLAND, _
             XAPPL, XDEPS, XDEPD, XCONC, _
             NPlong, TPD.X(0), TPD.Y(0, 0))

  'Finish up plot data
  'We can convert the data here, because we know the units will
  'not change within the life of the data
  TPD.NC = 1 'number of data sets defined
  TPD.np(0) = CInt(NPlong)
  For i = 0 To TPD.np(0) - 1
    TPD.X(i) = UnitsDisplay(TPD.X(i), UN_LENGTH)
  Next

  'stuff the results into the form controls
  PropTakeAction = False
  If XLAND < 0 Then
    txtCalc(0).Text = "out of range!"
  Else
    txtCalc(0).Text = AGFormat$(UnitsDisplay(XLAND, UN_LENGTH))
  End If
  If XAPPL < 0 Then
    txtCalc(1).Text = "out of range!"
  Else
    txtCalc(1).Text = AGFormat$(XAPPL)
  End If
  If XDEPS < 0 Then
    txtCalc(2).Text = "out of range!"
  Else
    txtCalc(2).Text = AGFormat$(XDEPS)
  End If
  If XDEPD < 0 Then
    txtCalc(3).Text = "out of range!"
  Else
    txtCalc(3).Text = AGFormat$(XDEPD)
  End If
  If XCONC < 0 Then
    txtCalc(4).Text = "out of range!"
  Else
    txtCalc(4).Text = AGFormat$(XCONC)
  End If
  PropTakeAction = True

  NeedCalcs = False
  
  Me.MousePointer = vbDefault 'default
End Sub

Private Sub ClearCalcOutput()
'clear calc output fields
'Don't clear the one pointed to by CalcOutputMarker
  Dim PTAsave As Integer
  Dim c As Control
  
  PTAsave = PropTakeAction
  PropTakeAction = False
  For Each c In txtCalc
    If c.Index = CalcOutputMarker Then
      c.ForeColor = vbRed
    Else
      c.Text = ""
      c.ForeColor = vbBlack
    End If
  Next
  PropTakeAction = PTAsave
End Sub

Private Sub cmdCalc_Click()
  If NeedCalcs Then Calculate
End Sub

Private Sub cmdExport_Click()
  If NeedCalcs Then Calculate
  If Not NeedCalcs Then
    If GenPlotTitles(PV_TERR, False) And GenPlotUnits(PV_TERR) Then
      frmExportToolbox.Show vbModal
    End If
  End If
End Sub

Private Sub cmdOk_Click()
  Hide
End Sub

Private Sub cmdPlot_Click()
  Dim saveDataSource(4) As String
  Dim saveDataTitle(4) As String
  Dim strTitle As String
  Dim lngID As Long
  Dim i As Integer
  Dim j As Integer
  Dim s As String
  On Error GoTo cmdPlotErrHand
  If NeedCalcs Then Calculate
  If NeedCalcs Then Exit Sub
  
  'Save sources and titles
  For i = 0 To 4: saveDataSource(i) = PlotGetDataSource(i): Next 'save existing
  For i = 0 To 4: saveDataTitle(i) = PlotGetDataTitle(i): Next 'save existing
  'Clear cources and titles
  For i = 0 To 4: PlotSetDataSource i, "": Next 'clear
  For i = 0 To 4: PlotSetDataTitle i, "": Next 'clear
  'Set sources and titles
  PlotSetDataSource 0, "ToolboxData: 0"
  GenPlotTitleStrings PV_TERR, False, strTitle, s, s, lngID
  PlotSetDataTitle 0, strTitle + ": " + txtDist.Text + " " + lblDistUnits.Caption
   
  'plot
  If SetupPlot(PV_TERR) Then frmPlot.Show vbModal
    
  'Restore sorces and titles
  For i = 0 To 4: PlotSetDataSource i, saveDataSource(i): Next
  For i = 0 To 4: PlotSetDataTitle i, saveDataTitle(i): Next
  Exit Sub

cmdPlotErrHand:
  Select Case UnexpectedError("cmdPlot,Click")
  Case vbAbort  'Abort - Stop the whole program
    End
  Case vbRetry  'Retry - Resume at the same line
    Resume
  Case vbIgnore 'Ignore - Resume at the next line
    Resume Next
  End Select
End Sub

Private Sub cmdPrint_Click()
'print the current UserData
  Dim BeginPage As Integer
  Dim EndPage As Integer
  Dim NumCopies As Integer
  Dim ReportText As String
  Dim i As Integer
  Dim pages As Integer
  Dim Mag As Variant

  If PrinterExists() Then
    If PrintDialog(BeginPage, EndPage, NumCopies) Then
      ReportText = GenFormData()
      For i = 1 To NumCopies
        PrintData ReportText, False, pages, Mag
      Next
    End If
  End If
End Sub

Private Sub Form_Activate()
'This routine is executed each time the form is shown
  Dim tmpBodyType As Integer
  Dim c As Control
  Dim PTAsave As Integer

  'Turn off control sensitivity
  PTAsave = PropTakeAction
  PropTakeAction = False
  
  'Adjust the controls to suit the Tier
  If UD.Tier = 1 Then
    fraTier1.Enabled = True
    lblActiveRate.Enabled = True
    txtActiveRate.Enabled = True
    If txtActiveRate.Text = "" Then _
      txtActiveRate.Text = AGFormat$(UnitsDisplay( _
        UD.SM.ACfrac * UD.SM.FlowRate * UD.SM.NonVGrav, UN_RATEMASS))
    lblActiveRateUnits.Enabled = True
  Else
    fraTier1.Enabled = False
    lblActiveRate.Enabled = False
    txtActiveRate.Enabled = False
    txtActiveRate.Text = ""
    lblActiveRateUnits.Enabled = False
  End If
  
  'If the units have changed since the last time
  'update a few things
  If UP.Units <> PreviousUnits Then
    UpdateUnitsLabels
    'Find the aquatic body type
    tmpBodyType = -1
    For Each c In optBodyType
      If c.Value Then
        tmpBodyType = c.Index
        Exit For
      End If
    Next
    'Convert any existing user-defined values
    If PreviousUnits <> -1 Then
      If txtDist.Text <> "" Then
        txtDist.Text = AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtDist.Text), UN_LENGTH, PreviousUnits), UN_LENGTH))
      End If
    End If
    'Change the units of the input dimensions
    PropTakeAction = True
    optBodyType_Click tmpBodyType  'reselect the option
    PropTakeAction = False
    'Change the units of the Active Rate if this is not the first time the form is shown
    If PreviousUnits <> -1 Then
      txtActiveRate.Text = AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtActiveRate.Text), _
        UN_RATEMASS, PreviousUnits), UN_RATEMASS))
    End If
    PreviousUnits = UP.Units 'save the new setting
  End If
  'perform a new calc no matter what
  
  'Restore control sensitivity
  PropTakeAction = PTAsave
  
  Calculate
End Sub

Private Sub Form_Load()
'This routine is executed only when the form is first loaded.
'Setting PreviousUnits to -1 assures that Form_Activate,
'which is executed after this routine, will update the units
'labels and perform an initial calc.

  CenterForm Me 'Center the form on the screen
  
  PropTakeAction = True       'Activate control reactions
  CalcOutputMarker = -1       'Init Calc text box marker
  optBodyType(0).Value = True 'Default body type
  txtActiveRate.Text = AGFormat$(UnitsDisplay( _
    UD.SM.ACfrac * UD.SM.FlowRate * UD.SM.NonVGrav, UN_RATEMASS))
  txtCalc(0).Text = AGFormat$(UnitsDisplay(60.96, UN_LENGTH)) 'Initial calc val
  PreviousUnits = -1          'Init Units save flag
End Sub

Private Function GenFormData() As String
'Generate report text for this form to be used for printing
  
  Dim gfd As String  'temporary storage for report text
  Dim s As String        'workspace string
  Dim i As Integer

  gfd = "" 'start with a blank string
  
  AppendStr gfd, "AgDRIFT® Terrestrial Assessment", True
  AppendStr gfd, "", True
  
  AppendStr gfd, "Terrestrial Field: " & optBodyType(BodyType).Caption, True
  AppendStr gfd, "  " & lblDist.Caption & " " & txtDist.Text & " " & lblDistUnits.Caption, True
  AppendStr gfd, "", True

  If UD.Tier = TIER_1 Then
    AppendStr gfd, "Active Rate:", True
    AppendStr gfd, "  " & lblActiveRate.Caption & " " & txtActiveRate.Text & " " & lblActiveRateUnits.Caption, True
    AppendStr gfd, "", True
  End If
  
  AppendStr gfd, "Calculations:", True
  AppendStr gfd, "  " & lblResults1.Caption & " " & txtCalc(0).Text & " " & lblCalcDistUnits.Caption, False
  If CalcOutputMarker = 0 Then
    AppendStr gfd, " (input)", True
  Else
    AppendStr gfd, "", True
  End If
  AppendStr gfd, "  " & lblResults2.Caption & " " & txtCalc(1).Text & " " & lblUnits1.Caption, False
  If CalcOutputMarker = 1 Then
    AppendStr gfd, " (input)", True
  Else
    AppendStr gfd, "", True
  End If
  AppendStr gfd, "  " & lblResults2.Caption & " " & txtCalc(2).Text & " " & lblUnits2.Caption, False
  If CalcOutputMarker = 2 Then
    AppendStr gfd, " (input)", True
  Else
    AppendStr gfd, "", True
  End If
  AppendStr gfd, "  " & lblResults2.Caption & " " & txtCalc(3).Text & " " & lblUnits3.Caption, False
  If CalcOutputMarker = 3 Then
    AppendStr gfd, " (input)", True
  Else
    AppendStr gfd, "", True
  End If
  AppendStr gfd, "  " & lblResults2.Caption & " " & txtCalc(4).Text & " " & lblUnits4.Caption, False
  If CalcOutputMarker = 3 Then
    AppendStr gfd, " (input)", True
  Else
    AppendStr gfd, "", True
  End If
  AppendStr gfd, "", True
  
  AppendStr gfd, "Tier: " & String$(UD.Tier, "I"), True
  AppendStr gfd, "RunID:", True
  AppendStr gfd, "  " & GetRunID(), True
  AppendStr gfd, "", True
  
  GenFormData = gfd
End Function

Private Sub optBodyType_Click(Index As Integer)
  If PropTakeAction Then
    PropTakeAction = False
    Select Case Index
    Case 0  'Single-Plant
      txtDist.Text = AGFormat$(UnitsDisplay(63.613, UN_LENGTH))
      cmdPlot.Enabled = False
      cmdExport.Enabled = False
    Case 2  'User-defined
      cmdPlot.Enabled = True
      cmdExport.Enabled = True
    End Select
    BodyType = Index
    ClearCalcOutput
    NeedCalcs = True
    PropTakeAction = True
  End If
End Sub

Private Sub txtActiveRate_Change()
'When this control changes, clear the
'calc output.
  If PropTakeAction Then
    ClearCalcOutput
    NeedCalcs = True
  End If
End Sub

Private Sub txtCalc_Change(Index As Integer)
'When this control changes, clear the other
'members of the array
  If PropTakeAction Then
    PropTakeAction = False
    If Trim$(txtCalc(Index).Text) = "" Then
      CalcOutputMarker = -1
    Else
      CalcOutputMarker = Index
    End If
    ClearCalcOutput
    NeedCalcs = True
    PropTakeAction = True
  End If
End Sub

Private Sub txtDist_Change()
  If PropTakeAction Then
    optBodyType(2).Value = True 'flip type to user-defined
    ClearCalcOutput
    NeedCalcs = True
  End If
End Sub

Private Sub UpdateUnitsLabels()
  lblDistUnits.Caption = UnitsName(UN_LENGTH)
  lblActiveRateUnits.Caption = UnitsName(UN_RATEMASS)
  lblCalcDistUnits.Caption = UnitsName(UN_LENGTH)
End Sub

