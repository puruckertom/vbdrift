VERSION 5.00
Begin VB.Form frmTBSprayBlock 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spray Block Assessment"
   ClientHeight    =   6240
   ClientLeft      =   1320
   ClientTop       =   3030
   ClientWidth     =   5745
   ForeColor       =   &H80000008&
   HelpContextID   =   1405
   Icon            =   "TBSPRAYB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6240
   ScaleWidth      =   5745
   Begin VB.CommandButton cmdPlot 
      Caption         =   "Plo&t"
      Height          =   375
      HelpContextID   =   1405
      Left            =   2280
      TabIndex        =   3
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      HelpContextID   =   1405
      Left            =   3120
      TabIndex        =   2
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calc"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1405
      Left            =   3960
      TabIndex        =   1
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Clos&e"
      Height          =   375
      HelpContextID   =   1405
      Left            =   4800
      TabIndex        =   0
      Top             =   5760
      Width           =   855
   End
   Begin VB.Frame fraDefinition 
      Caption         =   "Definition"
      Height          =   3015
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   5535
      Begin VB.Frame fraPond 
         Caption         =   "Water Body Description"
         Height          =   1935
         Left            =   480
         TabIndex        =   21
         Top             =   960
         Width           =   4935
         Begin VB.OptionButton optPondType 
            Caption         =   "User-defined Water Body"
            Height          =   255
            HelpContextID   =   1405
            Index           =   2
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txtDepth 
            Height          =   285
            HelpContextID   =   1507
            Left            =   2760
            TabIndex        =   10
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtDistance 
            Height          =   285
            HelpContextID   =   1506
            Left            =   2760
            TabIndex        =   9
            Top             =   1080
            Width           =   855
         End
         Begin VB.OptionButton optPondType 
            Caption         =   "EPA-Defined Wetland"
            Height          =   255
            HelpContextID   =   1405
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   2895
         End
         Begin VB.OptionButton optPondType 
            Caption         =   "EPA-Defined Pond"
            Height          =   255
            HelpContextID   =   1405
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lblDepthUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Left            =   3720
            TabIndex        =   25
            Top             =   1485
            Width           =   420
         End
         Begin VB.Label lblDepth 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Average Depth:"
            Height          =   195
            Left            =   1080
            TabIndex        =   24
            Top             =   1485
            Width           =   1590
         End
         Begin VB.Label lblDistanceUnits 
            Caption         =   "units"
            Height          =   255
            Left            =   3720
            TabIndex        =   23
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblDistance 
            Alignment       =   1  'Right Justify
            Caption         =   "Downwind Water Body Width:"
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   1080
            Width           =   2295
         End
      End
      Begin VB.OptionButton optDepos 
         Caption         =   "Pond-Integrated Deposition"
         Height          =   255
         HelpContextID   =   1405
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   3135
      End
      Begin VB.OptionButton optDepos 
         Caption         =   "Deposition"
         Height          =   255
         HelpContextID   =   1405
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraTier1 
      Caption         =   "Tier I Settings"
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   3120
      Width           =   5535
      Begin VB.TextBox txtActiveRate 
         Height          =   285
         HelpContextID   =   1010
         Left            =   3240
         TabIndex        =   11
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lblActiveRate 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Active Rate:"
         Height          =   195
         Left            =   2160
         TabIndex        =   30
         Top             =   285
         Width           =   900
      End
      Begin VB.Label lblActiveRateUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   4185
         TabIndex        =   29
         Top             =   285
         Width           =   420
      End
   End
   Begin VB.Frame fraCalc 
      Caption         =   "Calculations"
      Height          =   1815
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   5535
      Begin VB.TextBox txtCalc 
         Height          =   285
         HelpContextID   =   1543
         Index           =   3
         Left            =   1920
         TabIndex        =   15
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         HelpContextID   =   1542
         Index           =   2
         Left            =   3720
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         HelpContextID   =   1542
         Index           =   1
         Left            =   1920
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         HelpContextID   =   1542
         Index           =   0
         Left            =   1920
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblConcUnits 
         AutoSize        =   -1  'True
         Caption         =   "ng/L (ppt)"
         Height          =   195
         Left            =   3240
         TabIndex        =   31
         Top             =   1365
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "lb/ac"
         Height          =   195
         Left            =   5040
         TabIndex        =   27
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lblConc 
         Alignment       =   1  'Right Justify
         Caption         =   "Concentration Level:"
         Height          =   225
         Left            =   120
         TabIndex        =   26
         Top             =   1365
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "g/ha"
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Fraction of Applied"
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         Caption         =   "Deposition Level:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmTBSprayBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tbsprayb.frm,v 1.9 2001/08/30 14:00:35 tom Exp $
Option Explicit

Dim PropTakeAction As Integer
Dim NeedCalcs As Integer     'tracks calculation status
Dim CalcOutputMarker As Integer 'tracks output selection
Dim PreviousUnits As Integer 'Tracks units setting

Private Sub ClearOutputFields()
'clear Deposition Level fields
'Don't clear the one pointed to by CalcOutputMarker
  Dim PTAsave As Integer
  Dim c As Control
  
  PTAsave = PropTakeAction
  PropTakeAction = False 'desensitize controls
  For Each c In txtCalc
    If c.Index = CalcOutputMarker Then
      c.ForeColor = vbRed
    Else
      c.Text = ""
      c.ForeColor = vbBlack
    End If
  Next
  PropTakeAction = PTAsave 'restore control sensitivity
End Sub

Private Sub Calculate()
'Calculate the Toxicity (exposure)
  Dim IDEP As Long
  Dim INTYPE As Long
  Dim XLENG As Single
  Dim XDEEP As Single
  Dim XACT As Single
  Dim XAPPL As Single
  Dim XDEPS As Single
  Dim XDEPD As Single
  Dim XCONC As Single

  Dim Msg As String
  Dim NPlong As Long
  Dim i As Integer
  Dim c As Control

  ' Change the form mouse pointer
  Me.MousePointer = vbHourglass 'hourglass
  
  'Extract the input data from the form controls
  For Each c In optDepos
    If c.Value Then IDEP = c.Index
  Next
  INTYPE = CLng(CalcOutputMarker)
  XLENG = UnitsInternal(txtDistance.Text, UN_LENGTH)
  XDEEP = UnitsInternal(txtDepth.Text, UN_LENGTH)
  XACT = UnitsInternal(Val(txtActiveRate.Text) / (UD.SM.FlowRate * UD.SM.NonVGrav), UN_RATEMASS)
  XAPPL = Val(txtCalc(0).Text)
  XDEPS = Val(txtCalc(1).Text)
  XDEPD = Val(txtCalc(2).Text)
  XCONC = Val(txtCalc(3).Text)

  'Set up the global area to store the calculated results
  TPD.X1D = True 'X data is one-dimensional
  ReDim TPD.np(0)
  ReDim TPD.X(49)
  ReDim TPD.Y(49, 0)
  
  'Do the calculations
  Call agsblk(UD, _
              CLng(UC.NumSgl), UC.SglDist(0), UC.SglVal(0), _
              UC.HalfVal(0), _
              IDEP, INTYPE, XLENG, XDEEP, XACT, _
              XAPPL, XDEPS, XDEPD, XCONC, _
              NPlong, TPD.X(0), TPD.Y(0, 0))
  
  'Finish up plot data
  'We can convert the data here, because we know the units will
  'not change within the life of the data
  TPD.NC = 1 'number of data sets defined
  TPD.np(0) = CInt(NPlong)
  For i = 0 To TPD.np(0) - 1
    TPD.X(i) = UnitsDisplay(TPD.X(i), UN_LENGTH)
    TPD.Y(i, 0) = UnitsDisplay(TPD.Y(i, 0), UN_LENGTH)
  Next

  'Place calculated values back in the controls
  PropTakeAction = False
  Select Case CalcOutputMarker
  Case 0
    txtCalc(1) = AGFormat$(XDEPS)
    txtCalc(2) = AGFormat$(XDEPD)
    If IDEP = 1 Then txtCalc(3) = AGFormat$(XCONC)
  Case 1
    txtCalc(0) = AGFormat$(XAPPL)
    txtCalc(2) = AGFormat$(XDEPD)
    If IDEP = 1 Then txtCalc(3) = AGFormat$(XCONC)
  Case 2
    txtCalc(0) = AGFormat$(XAPPL)
    txtCalc(1) = AGFormat$(XDEPS)
    If IDEP = 1 Then txtCalc(3) = AGFormat$(XCONC)
  Case 3
    txtCalc(0) = AGFormat$(XAPPL)
    txtCalc(1) = AGFormat$(XDEPS)
    txtCalc(2) = AGFormat$(XDEPD)
  End Select
  PropTakeAction = True
  
  NeedCalcs = False
  
  Me.MousePointer = vbDefault 'default
End Sub

Private Sub cmdCalc_Click()
  If NeedCalcs Then Calculate
End Sub

Private Sub cmdExport_Click()
  If NeedCalcs Then Calculate
  If Not NeedCalcs Then
    If GenPlotTitles(PV_SBLK, False) And GenPlotUnits(PV_SBLK) Then
      frmExportToolbox.Show vbModal
    End If
  End If
End Sub

Private Sub cmdOk_Click()
  'Reset calc flag, since we don't know if new
  'main calcs will be performed or loaded
  NeedCalcs = True
  ClearOutputFields
  Hide
End Sub

Private Sub cmdPlot_Click()
  Dim saveDataSource(4) As String
  Dim saveDataTitle(4) As String
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
  PlotSetDataSource 0, "ToolboxData: 0": PlotSetDataTitle 0, "Buffer"
    
  'plot
  If SetupPlot(PV_SBLK) Then frmPlot.Show vbModal
    
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

Private Sub Form_Activate()
'This routine is executed each time the form is shown

  'adjust controls to suit the tier
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
  
  'If the units have changed since the last time this
  'form was shown, update a few things
  If UP.Units <> PreviousUnits Then
    'update units labels
    lblActiveRateUnits = UnitsName(UN_RATEMASS)
    lblDistanceUnits.Caption = UnitsName(UN_LENGTH)
    lblDepthUnits.Caption = UnitsName(UN_LENGTH)
    
    'Convert any existing user-defined values
    '(if this is not the first time this form has been shown)
    If PreviousUnits <> -1 Then
      If txtDistance.Text <> "" Then
        txtDistance.Text = _
          AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtDistance.Text), _
          UN_LENGTH, PreviousUnits), UN_LENGTH))
      End If
      If txtDepth.Text <> "" Then
        txtDepth.Text = _
          AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtDepth.Text), _
          UN_LENGTH, PreviousUnits), UN_LENGTH))
      End If
      If txtActiveRate.Text <> "" Then
        txtActiveRate.Text = _
          AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtActiveRate.Text), _
          UN_RATEMASS, PreviousUnits), UN_RATEMASS))
      End If
    End If
    PreviousUnits = UP.Units 'save the new setting
  End If
End Sub

Private Sub Form_Load()
'This routine is executed only when the form is first loaded.
'Setting PreviousUnits to -1 assures that Form_Activate,
'which is executed after this routine, will update the units
'labels and perform an initial calc.
  
  CenterForm Me 'Center the form on the screen

  PropTakeAction = True         'Activate control reactions
  txtCalc(0).Text = 0.1         'Default deposition level
  optDepos(0).Value = True      'Default deposition type
  optPondType(0).Value = True   'Default pond type
  txtActiveRate.Text = AGFormat$(UnitsDisplay( _
    UD.SM.ACfrac * UD.SM.FlowRate * UD.SM.NonVGrav, UN_RATEMASS))
  PreviousUnits = -1
End Sub

Private Function GenFormData() As String
'Generate report text for this form to be used for printing
  
End Function

Private Sub optDepos_Click(Index As Integer)
  Dim c As Control
  If PropTakeAction Then
    'Update Pond Controls
    Select Case Index
    Case 0 'Deposition
      fraPond.Enabled = False
      For Each c In optPondType
        c.Enabled = False
      Next
      lblDistance.Enabled = False
      txtDistance.Enabled = False
      lblDistanceUnits.Enabled = False
      lblDepth.Enabled = False
      txtDepth.Enabled = False
      lblDepthUnits.Enabled = False
      lblConc.Enabled = False
      txtCalc(3).Enabled = False
      lblConcUnits.Enabled = False
      If CalcOutputMarker = 2 Then 'can't leave numbers in conc
        'if there is a number, leave it, otherwise default
        If txtCalc(0).Text = "" Then
          txtCalc(0) = 0.1 'default
        Else
          CalcOutputMarker = 0 'move marker to deposition
        End If
      End If
    Case 1 'Pond-integrated deposition
      fraPond.Enabled = True
      For Each c In optPondType
        c.Enabled = True
      Next
      lblDistance.Enabled = True
      txtDistance.Enabled = True
      lblDistanceUnits.Enabled = True
      lblDepth.Enabled = True
      txtDepth.Enabled = True
      lblDepthUnits.Enabled = True
      lblConc.Enabled = True
      txtCalc(3).Enabled = True
      lblConcUnits.Enabled = True
    End Select
    '
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
  End If
End Sub

Private Sub optPondType_Click(Index As Integer)
  If PropTakeAction Then
    Select Case Index
    Case 0  'EPA pond
      PropTakeAction = False
      txtDistance.Text = AGFormat$(UnitsDisplay(63.613, UN_LENGTH))
      txtDepth.Text = AGFormat$(UnitsDisplay(2, UN_LENGTH))
      PropTakeAction = True
    Case 1  'EPA wetland
      PropTakeAction = False
      txtDistance.Text = AGFormat$(UnitsDisplay(63.613, UN_LENGTH))
      txtDepth.Text = AGFormat$(UnitsDisplay(0.15, UN_LENGTH))
      PropTakeAction = True
    Case 2  'User-defined
    End Select
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
  End If
End Sub

Private Sub txtActiveRate_Change()
'When this control changes, clear the
'calc output.
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
  End If
End Sub

Private Sub txtDepth_Change()
  If PropTakeAction Then
    optPondType(2).Value = True 'set to user-defined
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
  End If
End Sub

Private Sub txtDistance_Change()
  If PropTakeAction Then
    optPondType(2).Value = True 'set to user-defined
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
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
    NeedCalcs = True
    ClearOutputFields
    PropTakeAction = True
  End If
End Sub

