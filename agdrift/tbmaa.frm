VERSION 5.00
Begin VB.Form frmTBMultiApp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multiple Application Assessment"
   ClientHeight    =   5400
   ClientLeft      =   2700
   ClientTop       =   1890
   ClientWidth     =   7395
   HelpContextID   =   1449
   Icon            =   "TBMAA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5400
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTerrestrial 
      Caption         =   "Terrestrial Assessment"
      Height          =   495
      HelpContextID   =   1449
      Left            =   1920
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdAquatic 
      Caption         =   "Aquatic Assessment"
      Height          =   495
      HelpContextID   =   1449
      Left            =   720
      TabIndex        =   6
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdExams 
      Caption         =   "E&XAMS"
      Height          =   495
      HelpContextID   =   1449
      Left            =   4800
      TabIndex        =   2
      Top             =   4800
      Width           =   735
   End
   Begin VB.Frame fraWindRose 
      Caption         =   "Wind Rose"
      Height          =   3135
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cmdPlotWindRose 
         Caption         =   "Plot Wind Rose Probability"
         Height          =   375
         HelpContextID   =   1449
         Left            =   2160
         TabIndex        =   13
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CommandButton cmdPlotMedian 
         Caption         =   "Plot "
         Height          =   315
         HelpContextID   =   1522
         Left            =   2640
         TabIndex        =   10
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox cboLibPlot 
         Height          =   315
         HelpContextID   =   1522
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Frame fraDataSel 
         Caption         =   "Data Selection"
         Height          =   2175
         Left            =   3360
         TabIndex        =   29
         Top             =   240
         Width           =   3735
         Begin VB.ComboBox cboFieldDir 
            Height          =   315
            HelpContextID   =   1526
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1680
            Width           =   735
         End
         Begin VB.ComboBox cboMaxWS 
            Height          =   315
            HelpContextID   =   1524
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   360
            Width           =   735
         End
         Begin VB.ComboBox cboMonth 
            Height          =   315
            HelpContextID   =   1525
            Index           =   0
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   840
            Width           =   735
         End
         Begin VB.ComboBox cboMonth 
            Height          =   315
            HelpContextID   =   1525
            Index           =   1
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblFieldDirUnits 
            AutoSize        =   -1  'True
            Caption         =   "deg"
            Height          =   195
            Left            =   3000
            TabIndex        =   47
            Top             =   1680
            Width           =   270
         End
         Begin VB.Label lblFieldDir 
            Alignment       =   2  'Center
            Caption         =   "Direction to Sensitive Area"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label lblMaxWS 
            Caption         =   "Max. Wind Speed"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblMonth 
            Caption         =   "Ending Month"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lblMonth 
            Caption         =   "Beginning Month"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblMaxWSUnits 
            AutoSize        =   -1  'True
            Caption         =   "m/s"
            Height          =   195
            Left            =   3000
            TabIndex        =   30
            Top             =   360
            Width           =   270
         End
      End
      Begin VB.OptionButton optType 
         Caption         =   "Library"
         Height          =   255
         HelpContextID   =   1522
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Open..."
         Height          =   375
         HelpContextID   =   1523
         Left            =   360
         TabIndex        =   12
         Top             =   2280
         Width           =   855
      End
      Begin VB.OptionButton optType 
         Caption         =   "User-defined"
         Height          =   255
         HelpContextID   =   1523
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1215
      End
      Begin VB.ComboBox cboLibrary 
         Height          =   315
         HelpContextID   =   1522
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "Plo&t"
      Height          =   495
      HelpContextID   =   1449
      Left            =   3120
      TabIndex        =   4
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   495
      HelpContextID   =   1449
      Left            =   3960
      TabIndex        =   3
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calc"
      Height          =   495
      HelpContextID   =   1449
      Left            =   5640
      TabIndex        =   1
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Clos&e"
      Height          =   495
      HelpContextID   =   1449
      Left            =   6480
      TabIndex        =   0
      Top             =   4800
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame fraControl 
      Caption         =   "Control"
      Height          =   1575
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   2655
      Begin VB.TextBox txtEvents 
         Height          =   285
         HelpContextID   =   1527
         Left            =   1560
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtYears 
         Height          =   285
         HelpContextID   =   1528
         Left            =   1560
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Events per Year"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Years"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame fraMet 
      Caption         =   "Meteorology"
      Height          =   1575
      Left            =   2880
      TabIndex        =   28
      Top             =   3120
      Width           =   4455
      Begin VB.TextBox txtUserRH 
         Height          =   285
         HelpContextID   =   1529
         Left            =   3120
         TabIndex        =   23
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtUserTemp 
         Height          =   285
         HelpContextID   =   1529
         Left            =   1560
         TabIndex        =   22
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton optMetSource 
         Caption         =   "User-defined"
         Height          =   255
         HelpContextID   =   1529
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optMetSource 
         Caption         =   "Generated"
         Height          =   255
         HelpContextID   =   1529
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblMetSource 
         Caption         =   "Current"
         Height          =   255
         Left            =   360
         TabIndex        =   48
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblMetUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   5
         Left            =   3960
         TabIndex        =   45
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label lblMetUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   4
         Left            =   3960
         TabIndex        =   44
         Top             =   840
         Width           =   120
      End
      Begin VB.Label lblMetUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   3
         Left            =   3960
         TabIndex        =   43
         Top             =   480
         Width           =   120
      End
      Begin VB.Label lblMetUnits 
         AutoSize        =   -1  'True
         Caption         =   "deg C"
         Height          =   195
         Index           =   2
         Left            =   2400
         TabIndex        =   42
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label lblMetUnits 
         AutoSize        =   -1  'True
         Caption         =   "deg C"
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   41
         Top             =   840
         Width           =   420
      End
      Begin VB.Label lblMetUnits 
         AutoSize        =   -1  'True
         Caption         =   "deg C"
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   40
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblRH 
         AutoSize        =   -1  'True
         Caption         =   "Relative Humidity"
         Height          =   195
         Left            =   2880
         TabIndex        =   39
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Caption         =   "Temperature"
         Height          =   195
         Left            =   1440
         TabIndex        =   38
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblCurrRH 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   37
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblGenRH 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   36
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblCurrTemp 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   35
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblGenTemp 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   840
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmTBMultiApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tbmaa.frm,v 1.12 2003/08/22 17:01:55 tom Exp $
Option Explicit

Dim PropTakeAction As Integer
Dim NeedCalcs As Boolean     'tracks calculation status
Dim NeedCheck As Boolean     'Determines need for data checking
Dim PreviousUnits As Integer 'Tracks units setting
'grid editing vars
Dim gRow As Integer
Dim gCol As Integer

Dim StartDate

Dim MaxCalcWS As Integer  'Max Windspeed calculated so far
Dim MaxSrcWS As Long      'Max windspeed available from source

'Median data from library
Dim MedianData(11, 2, 2) As Single '12 months, 3 25-50-75 percentile, 3 T,RH,WS

'raw data for generating probability tables
Dim WD(11) As Single
Dim temp(11) As Single
Dim rhum(11) As Single
Dim FREQ(35, 19, 11) As Integer '0-350 deg by 10, WS 1-20 m/s, Month 1-12
Dim PROB(35, 19) As Single      '0-350 deg by 10, WS 1-20 m/s

'calculated stuff
'Composite Downwind Deposition
Dim numc As Long
Dim comd(24) As Single
Dim comv(24, 1) As Single
'Deposition
Dim numd As Long
Dim depd(MAX_CALCDATA) As Single
Dim depv(MAX_CALCDATA) As Single
'Pond-Integrated Deposition
Dim nump As Long
Dim pidd(MAX_CALCDATA) As Single
Dim pidv(MAX_CALCDATA) As Single

'EXAMS export data
Dim NEXREC As Long
Dim NEXPTS As Long      'same for all recs
Dim EXDIST() As Single  '2D - save all recs
Dim EXDEP() As Single   '1D - same for all recs
Dim EXRAR As Single

Public Function GetLocationName() As String
'Return the current Location Name (blank if user-defined)
  If optType(0).Value Then 'User-defined
    GetLocationName = ""
  ElseIf optType(1).Value Then
    GetLocationName = cboLibrary.Text
  End If
End Function

Private Sub cboFieldDir_Click()
  If PropTakeAction Then GenProb
End Sub

Private Sub cboLibrary_Click()
  If PropTakeAction Then GenLibraryProb
End Sub

Private Sub cboMaxWS_Click()
  If PropTakeAction Then
    If Not optType(0).Value Then GenProb
  End If
End Sub

Private Sub cboMonth_Click(Index As Integer)
  If PropTakeAction Then
    'Make sure beginning month is less
    'than or equal to ending month
    Select Case Index
    Case 0 'beginning month
      If cboMonth(0).ListIndex > cboMonth(1).ListIndex Then
        PropTakeAction = False
        cboMonth(1).ListIndex = cboMonth(0).ListIndex 'adjust ending month
        PropTakeAction = True
      End If
    Case 1 'ending month
      If cboMonth(0).ListIndex > cboMonth(1).ListIndex Then
        PropTakeAction = False
        cboMonth(0).ListIndex = cboMonth(1).ListIndex 'adjust beginning month
        PropTakeAction = True
      End If
    End Select
    GenProb
  End If
End Sub

Private Sub cmdAquatic_Click()
  If NeedCalcs Then Calculate
  If Not NeedCalcs Then
    'Load the form, then its data, then show it.
    Load frmTBAquatic
    frmTBAquatic.LoadDeposition CInt(numd), depd(), depv(), _
                                CInt(nump), pidd(), pidv()
    frmTBAquatic.Show vbModal
  End If
End Sub

Private Sub cmdPlotMedian_Click()
'Plot the Median values
  Dim saveDataSource(4) As String
  Dim saveDataTitle(4) As String
  Dim iMedian As Integer
  Dim i As Integer, j As Integer
  Dim s As String
  Dim PVtmp As Long
  On Error GoTo cmdPlotMedianErrHand
  
  'Select the Median value to plot
  iMedian = cboLibPlot.ListIndex                  '0=T 1=RH 2=WS 3=WD
  Select Case iMedian
  Case 0
    PVtmp = PV_MAAT
  Case 1
    PVtmp = PV_MAARH
  Case 2
    PVtmp = PV_MAAWS
  Case 3
    PVtmp = PV_MAAWD
  End Select
  
  'Save sources and titles
  For i = 0 To 4: saveDataSource(i) = PlotGetDataSource(i): Next 'save existing
  For i = 0 To 4: saveDataTitle(i) = PlotGetDataTitle(i): Next 'save existing
  'Clear cources and titles
  For i = 0 To 4: PlotSetDataSource i, "": Next 'clear
  For i = 0 To 4: PlotSetDataTitle i, "": Next 'clear
  'Set sources
  Select Case iMedian
  Case 0, 1, 2
    PlotSetDataSource 0, "ToolboxData: 0": PlotSetDataTitle 0, "25 Percent"
    PlotSetDataSource 1, "ToolboxData: 1": PlotSetDataTitle 1, "50 Percent"
    PlotSetDataSource 2, "ToolboxData: 2": PlotSetDataTitle 2, "75 Percent"
  Case 3
    PlotSetDataSource 0, "ToolboxData: 0": PlotSetDataTitle 0, "Dominant Wind Direction"
  End Select
  
  'Set TPD data
  Select Case iMedian
  Case 0, 1, 2   'T, RH, WS
    TPD.X1D = True           'one X for all Y curves
    TPD.NC = 3               'three curves
    ReDim TPD.np(0)          '1D curves need only 1 np
    ReDim TPD.X(11)          'X values
    ReDim TPD.Y(11, 2)       'Y values
    TPD.np(0) = 12
    For i = 0 To 11: TPD.X(i) = CSng(i + 1): Next 'Jan to Dec
    CopyMemory TPD.Y(0, 0), MedianData(0, 0, iMedian), TPD.np(0) * Len(MedianData(0, 0, 0))
    CopyMemory TPD.Y(0, 1), MedianData(0, 1, iMedian), TPD.np(0) * Len(MedianData(0, 0, 0))
    CopyMemory TPD.Y(0, 2), MedianData(0, 2, iMedian), TPD.np(0) * Len(MedianData(0, 0, 0))
    'units conversion
    Select Case iMedian
    Case 0 'T
      For i = 0 To TPD.np(0) - 1
        TPD.Y(i, 0) = UnitsDisplay(TPD.Y(i, 0), UN_TEMP)
        TPD.Y(i, 1) = UnitsDisplay(TPD.Y(i, 1), UN_TEMP)
        TPD.Y(i, 2) = UnitsDisplay(TPD.Y(i, 2), UN_TEMP)
      Next
    Case 2 'WS
      For i = 0 To TPD.np(0) - 1
        TPD.Y(i, 0) = UnitsDisplay(TPD.Y(i, 0), UN_SPEED)
        TPD.Y(i, 1) = UnitsDisplay(TPD.Y(i, 1), UN_SPEED)
        TPD.Y(i, 2) = UnitsDisplay(TPD.Y(i, 2), UN_SPEED)
      Next
    End Select
  Case 3     'WD
    TPD.X1D = True           'one X for all Y curves
    TPD.NC = 1               'one curve
    ReDim TPD.np(0)          '1D curves need only 1 np
    ReDim TPD.X(11)          'X values
    ReDim TPD.Y(11, 0)         'Y values
    TPD.np(0) = 12
    For i = 0 To 11: TPD.X(i) = CSng(i + 1): Next 'Jan to Dec
    CopyMemory TPD.Y(0, 0), WD(0), TPD.np(0) * Len(WD(0))
  End Select
  
  'plot
  If SetupPlot(PVtmp) Then frmPlot.Show vbModal
   
  'restore plot settings
  For i = 0 To 4: PlotSetDataSource i, saveDataSource(i): Next
  For i = 0 To 4: PlotSetDataTitle i, saveDataTitle(i): Next
  Exit Sub

cmdPlotMedianErrHand:
  Select Case UnexpectedError("cmdPlotMedian,Click")
  Case vbAbort  'Abort - Stop the whole program
    End
  Case vbRetry  'Retry - Resume at the same line
    Resume
  Case vbIgnore 'Ignore - Resume at the next line
    Resume Next
  End Select
End Sub

Private Sub cmdPlotWindRose_Click()
  Dim np As Long
  Dim deg(36) As Single
  Dim p10(36) As Single
  Dim p30(36) As Single
  Dim p50(36) As Single
  Dim p70(36) As Single
  Dim p90(36) As Single
  
  Dim saveDataSource(4) As String
  Dim saveDataTitle(4) As String
  Dim i As Integer
  Dim j As Integer
  Dim s As String
  On Error GoTo cmdPlotWindRoseErrHand
  
  'Save sources and titles
  For i = 0 To 4: saveDataSource(i) = PlotGetDataSource(i): Next 'save existing
  For i = 0 To 4: saveDataTitle(i) = PlotGetDataTitle(i): Next 'save existing
  'Clear cources and titles
  For i = 0 To 4: PlotSetDataSource i, "": Next 'clear
  For i = 0 To 4: PlotSetDataTitle i, "": Next 'clear
  'Set sources and titles
  PlotSetDataSource 0, "ToolboxData: 0": PlotSetDataTitle 0, "90th Percentile"
  PlotSetDataSource 1, "ToolboxData: 1": PlotSetDataTitle 1, "70th Percentile"
  PlotSetDataSource 2, "ToolboxData: 2": PlotSetDataTitle 2, "50th Percentile"
  PlotSetDataSource 3, "ToolboxData: 3": PlotSetDataTitle 3, "30th Percentile"
  PlotSetDataSource 4, "ToolboxData: 4": PlotSetDataTitle 4, "10th Percentile"
    
  'plot
  GenWindRosePlotDataTPD
  If SetupPlot(PV_MAAROSE) Then frmPlot.Show vbModal
   
  'Restore sorces and titles
  For i = 0 To 4: PlotSetDataSource i, saveDataSource(i): Next
  For i = 0 To 4: PlotSetDataTitle i, saveDataTitle(i): Next
  Exit Sub

cmdPlotWindRoseErrHand:
  Select Case UnexpectedError("cmdPlotWindRose,Click")
  Case vbAbort  'Abort - Stop the whole program
    End
  Case vbRetry  'Retry - Resume at the same line
    Resume
  Case vbIgnore 'Ignore - Resume at the next line
    Resume Next
  End Select
End Sub

Private Sub cmdTerrestrial_Click()
  If NeedCalcs Then Calculate
  If Not NeedCalcs Then
    'Load the form, then its data, then show it.
    Load frmTBTerrestrial
    frmTBTerrestrial.LoadDeposition CInt(numd), depd(), depv(), _
                                    CInt(nump), pidd(), pidv()
    frmTBTerrestrial.Show vbModal
  End If
End Sub

Private Sub cmdCalc_Click()
  If NeedCalcs Then Calculate
End Sub

Private Sub cmdExams_Click()
  If NeedCalcs Then Calculate
  If Not NeedCalcs Then ExportExams
End Sub

Private Sub ExportExams()
'Export Multiple Application Assessment EXAMS data
'
  Dim fn As String
  Dim Msg As String
  Dim dlm As String    'multi-column delimiter
  Dim hdr As String    'temporary storage for header text
  Dim s As String      'workspace string
  Dim c1wid As Integer 'number of columns to allot for column 1
  Dim c2wid As Integer 'number of columns to allot for column 2
  Dim c1fmt As String  'format string for column 1
  Dim c2fmt As String  'format string for column 2
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim nlong As Long

  'Prompt for a file name
  If Not FileDialog(FD_SAVEAS, FD_TYPE_TEXT, fn) Then
    Exit Sub
  End If
  
  On Error GoTo ErrHandExportExams
  Open fn For Output As #1

  'Part 1: EXAMS header
  '
  'This part lists values in a two-column format.
  'Column 1 describes the value and its units, if applicable.
  'Column 2 lists the value
  
  'set up the formats for the columns
  c1wid = 42  'column 1
  c2wid = 37  'column 2
  c1fmt = "!" & String$(c1wid, "@") 'left-justified
  c2fmt = " " & String$(c2wid, "@") '1 space, right-justified

  hdr = "" 'start with a blank string
    
  'General data
  AppendStr hdr, Format$("Title:", c1fmt), False
  AppendStr hdr, Format$(UD.Title, c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Run ID:", c1fmt), False
  AppendStr hdr, Format$(GetRunID(), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, "", True

  AppendStr hdr, Format$("Drop Size Distribution Name:", c1fmt), False
  AppendStr hdr, Format$(ClipStr$(UD.DSD(0).Name, c2wid), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Number of Drop Size Categories:", c1fmt), False
  AppendStr hdr, Format$(AGFormat$(UD.DSD(0).NumDrop), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, "", True
  
  AppendStr hdr, Format$("Spray Material Name:", c1fmt), False
  AppendStr hdr, Format$(ClipStr$(UD.SM.Name, c2wid), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Nonvolatile Rate (" & UnitsName(UN_RATEMASS) & "):", c1fmt), False
  AppendStr hdr, Format$(AGFormat$(UnitsDisplay(UD.SM.NVfrac * UD.SM.FlowRate * UD.SM.NonVGrav, UN_RATEMASS)), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Active Rate (" & UnitsName(UN_RATEMASS) & "):", c1fmt), False
  AppendStr hdr, Format$(AGFormat$(UnitsDisplay(UD.SM.ACfrac * UD.SM.FlowRate * UD.SM.NonVGrav, UN_RATEMASS)), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Spray Volume Rate (" & UnitsName(UN_RATEVOL) & "):", c1fmt), False
  AppendStr hdr, Format$(AGFormat$(UnitsDisplay(UD.SM.FlowRate, UN_RATEVOL)), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, "", True
    
  AppendStr hdr, Format$("Aircraft Name:", c1fmt), False
  AppendStr hdr, Format$(ClipStr$(UD.AC.Name, c2wid), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Number of Nozzles:", c1fmt), False
  AppendStr hdr, Format$(Format$(UD.NZ.NumNoz), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, "", True
    
  AppendStr hdr, Format$("Temperature (" & UnitsName(UN_TEMP) & "):", c1fmt), False
  AppendStr hdr, Format$(AGFormat$(UnitsDisplay(UD.MET.temp, UN_TEMP)), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Relative Humidity (%):", c1fmt), False
  AppendStr hdr, Format$(AGFormat$(UD.MET.Humidity), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, "", True
       
  AppendStr hdr, Format$("Release Height (" & UnitsName(UN_LENGTH) & "):", c1fmt), False
  AppendStr hdr, Format$(AGFormat$(UnitsDisplay(UD.CTL.Height, UN_LENGTH)), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Number of Spray Lines:", c1fmt), False
  AppendStr hdr, Format$(Format$(UD.CTL.NumLines), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Swath Width:", c1fmt), False
  Select Case UD.CTL.SwathWidthType
    Case 0
      s = AGFormat$(UnitsDisplay(UD.CTL.SwathWidth, UN_LENGTH)) + " " & UnitsName(UN_LENGTH)
    Case 1, 2
      s = AGFormat$(UD.CTL.SwathWidth) + " x Wingspan"
  End Select
  AppendStr hdr, Format$(s, c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Swath Displacement:", c1fmt), False
  Select Case UD.CTL.SwathDispType
    Case 0
      s = AGFormat$(UD.CTL.SwathDisp) + " x Swath Width"
    Case 1
      s = AGFormat$(UD.CTL.SwathDisp) + " x Application Rate"
    Case 2
      s = AGFormat$(UnitsDisplay(UD.CTL.SwathDisp, UN_LENGTH)) + " " & UnitsName(UN_LENGTH)
    Case 3
      s = "Aircraft Centerline"
  End Select
  AppendStr hdr, Format$(s, c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, "", True
  
  AppendStr hdr, Format$("Start Month:", c1fmt), False
  AppendStr hdr, Format$(Format$(cboMonth(0).ListIndex + 1), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("End Month:", c1fmt), False
  AppendStr hdr, Format$(Format$(cboMonth(1).ListIndex + 1), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Maximum Wind Speed (m/s):", c1fmt), False
  AppendStr hdr, Format$(Format$(cboMaxWS.ListIndex + 2), c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, "", True
       
  AppendStr hdr, Format$("Number of Applications (Events) per Year:", c1fmt), False
  AppendStr hdr, Format$(txtEvents.Text, c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Number of Years:", c1fmt), False
  AppendStr hdr, Format$(txtYears.Text, c2fmt), False
  AppendStr hdr, "", True
    
  Print #1, hdr

  'Part II: Downwind distances
  'get EXAMS export data
  ReDim EXDIST(MAX_CALCDATA)
  ReDim EXDEP(MAX_CALCDATA)
  hdr = ""
  AppendStr hdr, Format$("Y Values (m):", c1fmt), False
  AppendStr hdr, Format$(Format$(NEXPTS), c2fmt), False
  AppendStr hdr, "", True
  Print #1, hdr;
  
  'This call is for the distances
  Call agsmex(CLng(1), NEXPTS, EXDIST(0), EXDEP(0), EXRAR)
  
  k = 0 'keeps track of output values per line
  For i = 0 To NEXPTS - 1
    Print #1, Format$(Format$(EXDIST(i), "#########0.0"), "    @@@@@@@@@@@@");
    k = k + 1
    If k = 5 Then Print #1,: k = 0
  Next
  If k > 0 Then Print #1,
  
  'Part III: Applications
  For i = 0 To NEXREC - 1
    nlong = i + 1
    Call agsmex(nlong, NEXPTS, EXDIST(0), EXDEP(0), EXRAR)
    
    Print #1,
    Print #1, Format$(Format$(i + 1), "    @@@@@@@@@@@@");
    Print #1, Format$(EXRAR, "    0.000000E+00")
    k = 0 'keeps track of output values per line
    For j = 0 To NEXPTS - 1
      Print #1, Format$(EXDEP(j), "    0.000000E+00");
      k = k + 1
      If k = 5 Then Print #1, "": k = 0
    Next
    If k > 0 Then Print #1,
  Next

  Close #1
  Exit Sub
  
ErrHandExportExams:
  Msg = "Error writing file: " + fn + Chr$(13) + Error$(Err)
  MsgBox Msg, vbCritical + vbOKOnly
  Exit Sub
End Sub

Private Sub cmdExport_Click()
'Export plot data
  If NeedCalcs Then Calculate
  If Not NeedCalcs Then
    If GenPlotTitles(PV_MAA, False) And GenPlotUnits(PV_MAA) Then
      CopyPlotDataToTPD
      frmExportToolbox.Show vbModal
    End If
  End If
End Sub

Private Sub cmdImport_Click()
'Select a probability file and process
  NeedCalcs = True
  GetFileProb
End Sub

Private Sub cmdOk_Click()
  'Display a warning message
  If UD.Tier > 1 And Not NeedCalcs Then
    If MsgBox("None of the information generated in the " + _
              "toolbox calculations will be saved on exit. " + _
              "Exit the toolbox?", _
              vbOKCancel + vbExclamation) = vbCancel Then Exit Sub
  End If
  'Reset calc flag, since we don't know if new
  'main calcs will be performed or loaded
  NeedCalcs = True
  MaxCalcWS = 0 'need to recalc all speeds
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
  PlotSetDataSource 0, "ToolboxData: 0": PlotSetDataTitle 0, "Controlled Sample"
  PlotSetDataSource 1, "ToolboxData: 1": PlotSetDataTitle 1, "Maximum Deposition"
    
  'plot
  CopyPlotDataToTPD
  If SetupPlot(PV_MAA) Then frmPlot.Show vbModal
   
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
'(including right after calcs)
  Dim oldindex As Integer
  
  'adjust controls to suit the tier
  If UD.Tier = 1 Then
    fraMet.Enabled = False
    lblMetSource.Visible = False
    optMetSource(1).Enabled = False
    optMetSource(2).Visible = False
    lblTemp.Enabled = False
    lblCurrTemp.Visible = False
    lblGenTemp.Enabled = False
    txtUserTemp.Visible = False
    lblRH.Enabled = False
    lblCurrRH.Visible = False
    lblGenRH.Enabled = False
    txtUserRH.Visible = False
    lblMetUnits(0).Visible = False 'CurrTemp
    lblMetUnits(1).Enabled = False 'GenTemp
    lblMetUnits(2).Visible = False 'UserTemp
    lblMetUnits(3).Visible = False 'CurrRH
    lblMetUnits(4).Enabled = False 'GenRH
    lblMetUnits(5).Visible = False 'UserRH
    lblFieldDir.Enabled = False
    cboFieldDir.Enabled = False
    lblFieldDirUnits.Enabled = False
  Else
    fraMet.Enabled = True
    lblMetSource.Visible = True
    optMetSource(1).Enabled = True
    optMetSource(2).Visible = True
    lblTemp.Enabled = True
    lblCurrTemp.Visible = True
    lblGenTemp.Enabled = True
    txtUserTemp.Visible = True
    lblRH.Enabled = True
    lblCurrRH.Visible = True
    lblGenRH.Enabled = True
    txtUserRH.Visible = True
    lblMetUnits(0).Visible = True 'CurrTemp
    lblMetUnits(1).Enabled = True 'GenTemp
    lblMetUnits(2).Visible = True 'UserTemp
    lblMetUnits(3).Visible = True 'CurrRH
    lblMetUnits(4).Enabled = True 'GenRH
    lblMetUnits(5).Visible = True 'UserRH
    If UD.Tier = 2 Then
      lblFieldDir.Enabled = False
      cboFieldDir.Enabled = False
      lblFieldDirUnits.Enabled = False
    Else
      lblFieldDir.Enabled = True
      cboFieldDir.Enabled = True
      lblFieldDirUnits.Enabled = True
    End If
  End If
  
  'The Tier may have changed, so we need to update the
  'MaxWS combo. Compare the ListIndex before/after to see
  'if we need to recalc the probability table.
  oldindex = cboMaxWS.ListIndex
  SetupMaxWS
  If cboMaxWS.ListIndex <> oldindex Then
    If Not optType(0).Value Then GenProb
  End If
  
  'Set up the data checking flag
  'For Tier 1, no checking is required.
  'For other Tiers, checking is only required if
  'calcs have not been performed
  NeedCheck = False
  If UD.Tier > TIER_1 Then
    If Not UC.Valid Then
      NeedCheck = True
    End If
  End If
  
  'Update user data values
  lblCurrTemp.Caption = AGFormat$(UD.MET.temp)
  lblCurrRH.Caption = AGFormat$(UD.MET.Humidity)
 
End Sub

Private Sub Form_Load()
'Initialize the controls on this form
  Dim DB As Database
  Dim DS As Recordset
  Dim i As Integer

  CenterForm Me 'Center the form on the screen
  PropTakeAction = False
  PreviousUnits = -1
  NeedCalcs = True
  
  'MaxCalcWS is a special value that
  'keeps track of what wind speeds have
  'already been calculated. Set it to zero
  'to cause calculation of all available wind
  'speeds.
  MaxCalcWS = 0 'need to recalc all speeds
    
  'MaxSrcWS keeps track of the maximum windspeed
  'available from the current samson file or library entry.
  MaxSrcWS = 0
  
  'init the Library combo box
  If Not LibOpenMAADB(DB) Then Exit Sub
  Set DS = DB.OpenRecordset("WindRose", dbOpenDynaset)
  cboLibrary.Clear
  DS.MoveFirst
  Do Until DS.EOF
    cboLibrary.AddItem DS.Fields("Name")
    DS.MoveNext
  Loop
  DS.Close
  DB.Close
  cboLibrary.ListIndex = 0
  
  'init the Median plot combo box
  cboLibPlot.Clear
  cboLibPlot.AddItem "Median Temperature"
  cboLibPlot.AddItem "Median Relative Humidity"
  cboLibPlot.AddItem "Median Wind Speed"
  cboLibPlot.AddItem "Dominant Wind Direction"
  cboLibPlot.ListIndex = 0
  
  'init the month boxes
  For i = 0 To 1
    cboMonth(i).Clear
    cboMonth(i).AddItem "Jan"
    cboMonth(i).AddItem "Feb"
    cboMonth(i).AddItem "Mar"
    cboMonth(i).AddItem "Apr"
    cboMonth(i).AddItem "May"
    cboMonth(i).AddItem "Jun"
    cboMonth(i).AddItem "Jul"
    cboMonth(i).AddItem "Aug"
    cboMonth(i).AddItem "Sep"
    cboMonth(i).AddItem "Oct"
    cboMonth(i).AddItem "Nov"
    cboMonth(i).AddItem "Dec"
  Next
  cboMonth(0).ListIndex = 0
  cboMonth(1).ListIndex = 11
  
  'init the Field Direction Angle box
  cboFieldDir.Clear
  cboFieldDir.AddItem "None"
  cboFieldDir.ItemData(cboFieldDir.NewIndex) = 0
  For i = 1 To 36
    cboFieldDir.AddItem i * 10
    cboFieldDir.ItemData(cboFieldDir.NewIndex) = i * 10
  Next
  cboFieldDir.ListIndex = 0
  
  'Control
  txtEvents.Text = "1"
  txtYears.Text = "1"
  
  'Met
  PropTakeAction = True
  optMetSource(1).Value = True
  PropTakeAction = False
  txtUserTemp.Text = UD.MET.temp
  txtUserRH.Text = UD.MET.Humidity
  
  'Activate form controls
  PropTakeAction = True
  
  'Set Type button to Library. This will also load
  'a probability array into the grid
  optType(1).Value = True 'Library
End Sub

Private Sub Calculate()
'Calculations for Multiple Application Assessment
'load the calcuations form and do the calcs on the current data
  Dim f As Form
  Dim g As Control
  
  Dim i As Integer
  Dim ispd As Integer
  Dim idir As Integer
  Dim Msg As String
  Dim fn As String
  Dim ndiam As Long
  ReDim Diam(MAX_DROPS - 1) As Single
  Dim Compl(MAX_DROPS - 1) As Single
  Dim NXY As Long
  Dim Ndum As Long
  Dim Xdum As Single
  Dim Ydum As Single
  Dim ThermCount As Single
  Dim ThermTarget As Single
  Dim ShortMsg As String
  Dim LongMsg As String
  
  Dim iunits As Long
  Dim istat As Long
  Dim idk As Long
  Dim itype As Long
  Dim adat(2) As Single
  Dim cdat As String * 40
  Dim clen As Long
  Dim ier As Long
  Dim realwd(2) As Single
  
  Dim TEMPA As Single
  Dim RHUMA As Single
  Dim NEVNTS As Long
  Dim NYEARS As Long
  Dim NTSPD As Long
  
  Dim NPlong As Long
  Dim ev(3) As Single
  
  'display a warning message
  If UD.Tier > 1 Then
    NTSPD = cboMaxWS.ListIndex + 2
    If (NTSPD - MaxCalcWS) * 7 > 0 Then 'The 7 accounts for the extra wind directions
      If MsgBox("To complete the toolbox calculations will " + _
                "require " + CStr((NTSPD - MaxCalcWS) * 7) + " AgDRIFT runs. " + _
                "Continue?", _
                vbOKCancel + vbExclamation) = vbCancel Then Exit Sub
    End If
  End If
  
  'set up the calculation form
  Set f = frmTBMAACalc
  f.Show
  CenterForm f 'center the form
  f.lblStatusMessage(0).Caption = "" 'clear out the calc message
  f.lblStatusMessage(1).Caption = ""
  
  'Change the form mouse pointer
  f.MousePointer = vbHourglass
  
  'If this flag ever goes false, stop calculating
  UI.OkToDoCalcs = True
  
  UI.CalcsInProgress = True  'the calcs have begun!
  f.lblStatusMessage(0).Caption = "Starting calculations..."
  
  'Check data
  If UD.Tier > TIER_1 And NeedCheck Then
    If CheckData(f.lstCalcStat) Then
      NeedCheck = False
    Else
      GoTo CalculateHalt
    End If
  End If
  
  'record the date and time locally
  StartDate = Now

  'Enable the elapsed timer
  Timer1.Interval = 1000 'milliseconds
  Timer1.Enabled = True

  'Init Thermometer bar variables
  ThermCount = 0
  ThermTarget = 1 'revise this when below
  UpdateTherm ThermCount, ThermTarget
  DoEvents: If Not UI.OkToDoCalcs Then GoTo CalculateHalt
  
  'Time stamp the calc log
  Msg = ""
  AppendStr Msg, CStr(StartDate), False
  AppendStr Msg, " Calculations starting", False
  AddToLog f.lstCalcStat, Msg
  
  'Get form data
  If optMetSource(1).Value Then
    TEMPA = Val(lblGenTemp)
    RHUMA = Val(lblGenRH)
  ElseIf optMetSource(2).Value Then
    TEMPA = Val(txtUserTemp)
    RHUMA = Val(txtUserRH)
  End If
  NEVNTS = Val(txtEvents)
  NYEARS = Val(txtYears)
  NTSPD = cboMaxWS.ListIndex + 2
  
  'Check the MAA data passed from the MAA form
  Call agsmck(CLng(UD.Tier), TEMPA, RHUMA, _
              NEVNTS, NYEARS, PROB(0, 0), _
              NTSPD, ier, realwd(0), cdat, clen)
  If ier > 0 Then
    FormatAgreadMessage 2, ier, realwd(), cdat, ShortMsg, LongMsg
    AddToLog f.lstCalcStat, ShortMsg
    MsgBox LongMsg, vbCritical + vbOKOnly
    GoTo CalculateHalt
  End If
  
  Select Case UD.Tier
  Case 1 'Tier I
    Call agsmti(CLng(UD.ApplMethod), CLng(UC.NumDep), UC.DepDist(0), UC.DepVal(0))

  Case 2, 3 'Tier II, III
    'Loop through all wind speeds
    'Note that MaxCalcWS tells how many wind
    'speeds were already computed.
    ThermTarget = -1 'this is a flag for the init code below
    For ispd = MaxCalcWS + 1 To NTSPD
      For idir = 0 To 6 'loop through a fixed number of wind directions
        DoEvents: If Not UI.OkToDoCalcs Then GoTo CalculateHalt
        Msg = ""
        AppendStr Msg, CStr(Now), False
        AppendStr Msg, " Wind Speed ", False
        AppendStr Msg, Format$(ispd) & "/" & Format$(NTSPD), False
        AppendStr Msg, ", Dir ", False
        AppendStr Msg, Format$(idir + 1) & "/7", False
        AddToLog f.lstCalcStat, Msg
      
        'Pass the User Data to the fortran
        Call aginit(UD, CLng(ispd + idir * 100))
  
        'check the data with agread
        iunits = UP.Units    'set the units flag
        Do While True
          DoEvents: If Not UI.OkToDoCalcs Then GoTo CalculateHalt
          Call agread(iunits, istat, idk, itype, adat(0), cdat, clen)
          Select Case ProcessAgreadResults(istat, idk, itype, adat, cdat, clen)
          Case 0: 'keep looping
          Case 1: 'normal end, proceed
            Exit Do
          Case -1: 'problem, stop calculating
            GoTo CalculateHalt
          End Select
        Loop
  
        'Process all the drop categories
        DoEvents: If Not UI.OkToDoCalcs Then GoTo CalculateHalt
        Call aglims(ndiam, Diam(0), Compl(0))  'get computed drop dist
        If ThermTarget < 0 Then
          ThermTarget = ndiam * (NTSPD - MaxCalcWS) * 7 'the 7 accounts for the 0 to 6 wind dir loop
        End If
        For i = 0 To ndiam - 1
          f.lblStatusMessage(0).Caption = "Initial Drop Size: " & AGFormat$(Diam(i)) & " µm"
          DoEvents: If Not UI.OkToDoCalcs Then GoTo CalculateHalt
          Call agdrop(i + 1)
          ThermCount = ThermCount + 1
          UpdateTherm ThermCount, ThermTarget
        Next
  
        DoEvents: If Not UI.OkToDoCalcs Then GoTo CalculateHalt
   
        'get deposition (note that n, x, y are not used in agends)
        Call agends(AGENDS_MULDEP, Ndum, Xdum, Ydum)
        DoEvents: If Not UI.OkToDoCalcs Then GoTo CalculateHalt
      Next idir
    
      'this wind speed is done. Update MaxCalcWS
      If ispd > MaxCalcWS Then MaxCalcWS = ispd
    Next ispd
  End Select
  
  'get controlled sampling and plot data
  f.lblStatusMessage(0).Caption = "Completing calculations..."
  ThermTarget = 1
  ThermCount = 1
  UpdateTherm ThermCount, ThermTarget
  DoEvents: If Not UI.OkToDoCalcs Then GoTo CalculateHalt

  'get plot data
  Call agsmpl(numc, comd(0), comv(0, 0), NEXREC, _
              numd, depd(0), depv(0), _
              nump, pidd(0), pidv(0))
  
  'Success! reset the calc flags and set return value
  f.MousePointer = vbDefault
  UI.CalcsInProgress = False
  Timer1.Enabled = False
  Unload f
  NeedCalcs = False
  Exit Sub
  
CalculateHalt:
'stop the calculations and exit
  f.MousePointer = vbDefault
  UI.CalcsInProgress = False
  Timer1.Enabled = False
  Unload f

End Sub

Private Sub AddToLog(c As Control, s As String)
'Add a string to the Log control
  On Error Resume Next
  Dim frm As Form
  Dim fsave As Control
  Set frm = frmTBMAACalc
  c.AddItem s
  c.Refresh
  Set fsave = frm.ActiveControl 'save current control
  c.SetFocus           'set focus to list box
  SendKeys "{END}"                  'send an END key to the list box
  DoEvents
  c.Selected(c.ListIndex) = False   'turn off highlight
  fsave.SetFocus                    'restore original focus
End Sub

Private Function ProcessAgreadResults(istat As Long, idk As Long, itype As Long, adat() As Single, cdat As String, clen As Long) As Integer
'Act on the results of an agread call
'
' returns: 0: if agread should be called again
'          1: if agread need not be called again
'          2: if there was a problem and calculations should not proceed
'
' agread info: istat 0=don't write data, keep reading
'                    1=write data, keep reading
'                    2=error: write data, STOP
'                    3=end of data
'                    4=DropKick message, allow continue
'                    5=DropKick message, do not allow continue
'              idk   0=density mismatch in DropKick
'                    1=speed mismatch in DropKick
'                    2=Q mismatch in DropKick
'                      Note: idk=2 is a special case. No re-
'                      calculation is possible, so we can't
'                      display a recompute button.
'              itype 0=char data only
'                    1=int and char data
'                    2=real and char data
'          adat/idat for istat=1-3: (0)=value (1)=min (2)=max
'                    for istat=4-5: (0)=new value for DropKick
'                                   (1)=AgDRIFT value
'                                   (2)=DropKick value
  Dim ShortMsg As String
  Dim LongMsg As String
  Dim DKAction As Integer
  Dim f As Form
  
  ProcessAgreadResults = 0 'default: okay, more to read
  
  Set f = frmTBMAACalc
  
  'Build messages from agread return string
  FormatAgreadMessage istat, itype, adat(), cdat, ShortMsg, LongMsg
  
  'Act on various flag settings
  Select Case istat  'error level
  Case 0: 'normal, keep reading
    If ShortMsg <> "" Then AddToLog f.lstCalcStat, ShortMsg
  Case 1: 'warning with string
    If ShortMsg <> "" Then AddToLog f.lstCalcStat, ShortMsg
    If LongMsg <> "" Then
      If Not UP.SuppressTier3Warn Then
        AppendStr LongMsg, "", True
        AppendStr LongMsg, "Continue with calculations?", True
        If MsgBox(LongMsg, vbExclamation + vbYesNo) = vbNo Then
          ProcessAgreadResults = -1 'trouble
        End If
      End If
    End If
  Case 2: 'error with string; stop
    If ShortMsg <> "" Then AddToLog f.lstCalcStat, ShortMsg
    If LongMsg <> "" Then
      If Not UP.SuppressTier3Warn Then
        AppendStr LongMsg, "", True
        AppendStr LongMsg, "Continue with calculations?", True
        MsgBox LongMsg, vbCritical + vbOKOnly
        ProcessAgreadResults = -1 'trouble
      End If
    End If
  Case 3: 'normal end of data
    ProcessAgreadResults = 1 'normal end
  Case 4, 5: 'Special DropKick message: allow/do not allow continue
    'If certain of the input data to DropKick do not match the
    'corresponding input data in DropKick, there is a problem.
    'The user may
    'continue with everything as it is (only if ier=4), or
    'cancel the whole thing. The value that does not match is
    'indicated by IDK. If IDK=2 (Q), do not allow the option of
    'recomputing.
    If ShortMsg <> "" Then AddToLog f.lstCalcStat, ShortMsg
    If LongMsg <> "" Then
      'DropKick may not be rerun in this toolbox.
      'Set DKAction to:
      ' 2=do not recompute dropkick calcs, continue
      ' 3=abort entire calculation
      AppendStr LongMsg, "", True
      Select Case istat
      Case 4 'allow calcs to continue
        Select Case idk
        Case 0, 1, 2 'density, speed, Q
          AppendStr LongMsg, "Continue?", True
          Select Case MsgBox(LongMsg, vbExclamation + vbOKCancel)
          Case vbOK
            DKAction = 2 'don't recompute
          Case vbCancel
            DKAction = 3 'halt
          End Select
        End Select
      Case 5 'do not allow calcs to continue
        Select Case idk
        Case 0, 1, 2 'density, speed, Q
          AppendStr LongMsg, "Cannot continue calculations.", True
          Select Case MsgBox(LongMsg, vbExclamation + vbOKOnly)
          Case vbOK
            DKAction = 3 'halt
          End Select
        End Select
      End Select
      'Take the appropriate DropKick action
      Select Case DKAction
      Case 2 'don't recompute
        AddToLog f.lstCalcStat, "Not recomputing DropKick"
      Case 3 'halt
        ProcessAgreadResults = -1 'trouble
      End Select
    End If
  End Select
End Function

Private Sub UpdateTherm(curr_val As Single, max_val As Single)
'Update the Percent Completed Thermometer bar
  Dim f As Form
  Dim frac As Single
  Dim s As String
  Dim X As Single
  Dim SaveDrawMode As Integer
  
  Set f = frmTBMAACalc
  If max_val = 0 Then
    frac = 0
  Else
    frac = curr_val / max_val
  End If
  s = Format$(Int(frac * 100 + 0.5)) + "%"
  X = f.picTherm.Width * frac
  f.picTherm.Cls
  f.picTherm.CurrentX = (f.picTherm.Width - f.picTherm.TextWidth(s)) / 2
  f.picTherm.CurrentY = (f.picTherm.Height - f.picTherm.TextHeight(s)) / 2
  f.picTherm.Print s
  f.picTherm.Line (0, 0)-(X, f.picTherm.Height), RGB(255, 0, 0), BF
  f.picTherm.Refresh
'
' If the window is iconized, treat the whole form as a thermometer bar
  If f.WindowState = 1 Then
    X = f.Width * frac
    f.Cls
    f.CurrentX = (f.Width - f.TextWidth(s)) / 2
    f.CurrentY = (f.Height - f.TextHeight(s)) / 2
    SaveDrawMode = f.DrawMode
    f.DrawMode = 14 'Merge Pen Not
    f.Print s
    f.Line (0, 0)-(X, f.Height), QBColor(12), BF
    f.DrawMode = SaveDrawMode
    DoEvents                    ' Allow other events.
  End If
End Sub

Private Sub optMetSource_Click(Index As Integer)
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    MaxCalcWS = 0 'need to recalc all speeds
    If Index = 2 Then
      txtUserTemp.Enabled = True
      txtUserRH.Enabled = True
    Else
      txtUserTemp.Enabled = False
      txtUserRH.Enabled = False
    End If
  End If
End Sub

Private Sub optType_Click(Index As Integer)
'This method is executed only when the the option
'changes.
  If PropTakeAction Then
    UpdateTypeControls
    Select Case Index
    Case 0:  'user-defined
      'remove generated temp and rh
      lblGenTemp = ""
      lblGenRH = ""
      MaxSrcWS = 20
      SetupMaxWS
    Case 1:  'Library
      GenLibraryProb
    End Select
  End If
End Sub

Private Sub Timer1_Timer()
  Dim f As Form
  Set f = frmTBMAACalc
  f.lblStatusMessage(1).Caption = "Elapsed Time: " & Format$(CDbl(Now) - CDbl(StartDate), "hh:mm:ss")
End Sub

Private Sub GenLibraryProb()
'Retrieve frequency, temperature, and
'humidity tables from the library, and
'generate a probability table
  Dim DB As Database
  Dim DS As Recordset
  Dim bb() As Byte
  Dim nb As Long
  
  Dim str As String
  Dim maxspd As Long
  
  Me.MousePointer = vbHourglass
  
  'Retrieve the data from the library
  If Not LibOpenMAADB(DB) Then Exit Sub
  Set DS = DB.OpenRecordset("WindRose", dbOpenDynaset)
  
  DS.FindFirst "Name='" & cboLibrary.List(cboLibrary.ListIndex) & "'"
  If DS.NoMatch Then
    ClearProbControls
    DS.Close
    DB.Close
    Exit Sub
  End If
  FieldToArray DS.Fields("WD"), WD()
  FieldToArray DS.Fields("Temperature"), temp()
  FieldToArray DS.Fields("Humidity"), rhum()
  MaxSrcWS = DS.Fields("MaxSpeed")
  'Frequency is a special case because it is multidimensional
  'and is an integer array, so we can't use FieldToArray
  'directly.
  With DS.Fields("Frequency")
    nb = .FieldSize - 1   'bytes, less trailing null
    bb = .GetChunk(0, nb) 'retrieve raw data
    CopyMemory FREQ(0, 0, 0), bb(0), nb
  End With
  
  'Retrieve Median values while we're at it.
  '(Same deal here as for Frequency)
  With DS.Fields("T255075")
    nb = .FieldSize - 1   'bytes, less trailing null
    bb = .GetChunk(0, nb) 'retrieve raw data
    CopyMemory MedianData(0, 0, 0), bb(0), nb
  End With
  With DS.Fields("RH255075")
    nb = .FieldSize - 1   'bytes, less trailing null
    bb = .GetChunk(0, nb) 'retrieve raw data
    CopyMemory MedianData(0, 0, 1), bb(0), nb
  End With
  With DS.Fields("WS255075")
    nb = .FieldSize - 1   'bytes, less trailing null
    bb = .GetChunk(0, nb) 'retrieve raw data
    CopyMemory MedianData(0, 0, 2), bb(0), nb
  End With
  
  DS.Close
  DB.Close
  
  'Update the MaxWS combo box
  SetupMaxWS
  
  'generate probabilities
  GenProb
  
  Me.MousePointer = vbDefault
End Sub

Private Sub SetupMaxWS()
'configure the cboMaxWS combo box with a
'new max windspeed value.
'In Tier I the Max WS is restricted to 10 m/s
'In Tier II the Max WS is restricted to 12 m/s
  Dim PTAsave As Integer
  Dim MaxWS As Integer
  Dim oldindex As Integer
  Dim i As Integer
  
  PTAsave = PropTakeAction
  PropTakeAction = False

  'MaxSrcWS is a global that contains the max possible WS
  'from the current source
  MaxWS = MaxSrcWS

  'perform a basic limit check
  If MaxWS < 2 Then MaxWS = 2
  If MaxWS > 20 Then MaxWS = 20
  
  'Restrict windspeed for Tier I/II
  Select Case UD.Tier
  Case TIER_1
    If MaxWS > 10 Then MaxWS = 10
  Case TIER_2
    If MaxWS > 12 Then MaxWS = 12
  End Select
  
  'fill the combo box with windspeed choices starting from 2.
  'Try to retain any previous selection, but if none has
  'been made yet (ListIndex < 0), then select the maximum value
  oldindex = cboMaxWS.ListIndex 'save previous index
  If oldindex < 0 Then oldindex = MaxWS - 2 '
  'fill the combo box with items
  cboMaxWS.Clear
  For i = 2 To MaxWS
    cboMaxWS.AddItem Format$(i)
  Next
  'before restoring the previously selected value,
  'make sure it's in range
  If oldindex > cboMaxWS.ListCount - 1 Then
    oldindex = cboMaxWS.ListCount - 1
  End If
  cboMaxWS.ListIndex = oldindex
  PropTakeAction = PTAsave
End Sub

Private Sub ClearProbControls()
'Reset the maxWS, and clear the generated met
'controls in response to a bad library or samson file load
  MaxSrcWS = 0
  SetupMaxWS
  
  lblGenTemp = ""
  lblGenRH = ""
End Sub

Private Sub txtEvents_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
  End If
End Sub

Private Sub txtUserRH_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    MaxCalcWS = 0 'need to recalc all speeds
  End If
End Sub

Private Sub txtUserTemp_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    MaxCalcWS = 0 'need to recalc all speeds
  End If
End Sub

Private Sub txtYears_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
  End If
End Sub

Private Sub GetFileProb()
'Read a text file containing a probability table
  Dim fname As String
  Dim lun As Integer
  Dim num_entries As Integer
  Dim iws As Integer
  Dim tempg As Single
  Dim rhumg As Single
  Dim i As Integer
  Dim j As Integer
  
  'prompt for a file name
  If Not FileDialog(FD_OPEN, FD_TYPE_TEXT, fname) Then Exit Sub
  
  On Error GoTo GetFileProbEH1
  lun = FreeFile
  Open fname For Input As #lun
  
  On Error GoTo GetFileProbEH2
  
  Me.MousePointer = vbHourglass
  
  'frequency file format
  '=====================
  'temperature humidity
  'windspeed (first entry is always 2, entries always ascend by 1's)
  '(36 probabilities, one per line)
  'windspeed
  '(36 probabilities, one per line)
  'etc. (20 wind speeds max)
  '
  'will read until end-of-file
  Input #lun, tempg, rhumg
  On Error GoTo GetFileProbEH3
  For i = 0 To 19
    Input #lun, iws
    For j = 0 To 35
      Input #lun, PROB(j, iws - 1)
    Next
  Next
GetFileProbAfterReadLoop:
  MaxSrcWS = iws 'last wind speed in file is the max
  'clear the rest of the probability array (agsmck counts nonzero entries)
  For i = MaxSrcWS To 19
    For j = 0 To 35
      PROB(j, i) = 0
    Next
  Next
  
  Close #lun
  
  MaxCalcWS = 0 'reset calcs for all windpeeds
  
  SetupMaxWS 'redo the MaxWS combo
  
  'deliver the results to the form
  lblGenTemp = AGFormat$(tempg)
  lblGenRH = AGFormat$(rhumg)
  
  Me.MousePointer = vbDefault
  Exit Sub
  
GetFileProbEH1:
  MsgBox "Error opening file: " + fname + vbCrLf + Error$(Err), vbCritical + vbOKOnly
  Me.MousePointer = vbDefault
  Exit Sub

GetFileProbEH2:
  MsgBox "Error reading file: " + fname + vbCrLf + Error$(Err), vbCritical + vbOKOnly
  Close #lun
  Me.MousePointer = vbDefault
  Exit Sub

GetFileProbEH3:
  'this error handler is active when the probabilities
  'are read. An end-of-file here is the normal read
  'termination. Anything else is an error.
  If Err.Number = 62 Then 'input past end of file
    Resume GetFileProbAfterReadLoop 'Normal read termination
  Else
    Resume GetFileProbEH2           'Something unexpected
  End If
End Sub

Private Sub CopyPlotDataToTPD()
  'set up temporary plot data for plotting and for export
  'We must do this here because this toolbox can show other
  'toolboxes, which could change the shared TPD structure
  
  Dim i As Integer
  
  TPD.X1D = True           'one X for all Y curves
  TPD.NC = 2               'two curves
  ReDim TPD.np(0)          '1D curves need only 1 np
  ReDim TPD.X(numc - 1)    'X values
  ReDim TPD.Y(numc - 1, 1) 'Y values
  TPD.np(0) = numc
  For i = 0 To numc - 1
    TPD.X(i) = UnitsDisplay(comd(i), UN_LENGTH)
    TPD.Y(i, 0) = comv(i, 0)
    TPD.Y(i, 1) = comv(i, 1)
  Next
'  CopyMemory TPD.X(0), comd(0), numc * Len(comd(0))
'  CopyMemory TPD.Y(0, 0), comv(0, 0), numc * Len(comv(0, 0))
'  CopyMemory TPD.Y(0, 1), comv(0, 1), numc * Len(comv(0, 0))
End Sub

Private Sub GenWindRosePlotDataTPD()
  Dim np As Long
  Dim ic As Integer
  Dim ip As Integer
  'set up temporary plot data for plotting
  'We must do this here because this toolbox can show other
  'toolboxes, which could change the shared TPD structure
  TPD.X1D = True           'one X for all Y curves
  TPD.NC = 5               'five curves
  ReDim TPD.np(0)          '1D curves need only 1 np
  ReDim TPD.X(36)          'X values
  ReDim TPD.Y(36, TPD.NC - 1) 'Y values
  
  Call agwplt(PROB(0, 0), np, TPD.X(0), _
    TPD.Y(0, 4), TPD.Y(0, 3), TPD.Y(0, 2), TPD.Y(0, 1), TPD.Y(0, 0))
  
  TPD.np(0) = CInt(np)
  
  'units conversion
  For ic = 0 To TPD.NC - 1
    For ip = 0 To TPD.np(0) - 1
      TPD.Y(ip, ic) = UnitsDisplay(TPD.Y(ip, ic), UN_SPEED)
    Next
  Next
End Sub

Private Sub GenProb()
'Given a frequency table, temperature and humidity arrays,
'use agwdrs and form data to generate a probability table
'and populate the grid control
  
  'temp, rhum, and FREQ are all form-level public
  
  Dim user_max As Long
  Dim monb As Long
  Dim mone As Long
  Dim flddir As Long
  
  Dim tempg As Single
  Dim rhumg As Single
  
  Me.MousePointer = vbHourglass
  
  'Retrieve form data
  user_max = cboMaxWS.ListIndex + 2
  monb = cboMonth(0).ListIndex + 1
  mone = cboMonth(1).ListIndex + 1
  If monb > mone Then
    monb = mone
    mone = cboMonth(0).ListIndex + 1
  End If
  'guard against uninitialized combo box
  If cboFieldDir.ListIndex >= 0 Then
    flddir = cboFieldDir.ItemData(cboFieldDir.ListIndex)
  Else
    Exit Sub
  End If
  
  'generate probability table
  DoEvents
  Call agwdrs(flddir, temp(0), rhum(0), _
              user_max, FREQ(0, 0, 0), monb, mone, _
              tempg, rhumg, PROB(0, 0))

  'deliver the results to the form
  lblGenTemp = AGFormat$(tempg)
  lblGenRH = AGFormat$(rhumg)
  
  NeedCalcs = True 'Value changed, need to recalc
  
  Me.MousePointer = vbDefault
End Sub

Private Sub UpdateTypeControls()
'Update the type-related controls to reflect the
'currently selected type
  If optType(0).Value Then        'user-def
    cboLibrary.Enabled = False
    cboLibPlot.Enabled = False
    cmdPlotMedian.Enabled = False
    cmdImport.Enabled = True
    
    lblMonth(0).Enabled = False
    cboMonth(0).Enabled = False
    lblMonth(1).Enabled = False
    cboMonth(1).Enabled = False
  Else                            'library
    cboLibrary.Enabled = True
    cboLibPlot.Enabled = True
    cmdPlotMedian.Enabled = True
    cmdImport.Enabled = False
    
    lblMonth(0).Enabled = True
    cboMonth(0).Enabled = True
    lblMonth(1).Enabled = True
    cboMonth(1).Enabled = True
  End If
End Sub
