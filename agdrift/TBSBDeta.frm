VERSION 5.00
Begin VB.Form frmTBSprayBlockDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spray Block Details"
   ClientHeight    =   4665
   ClientLeft      =   1650
   ClientTop       =   1575
   ClientWidth     =   8760
   HelpContextID   =   1494
   Icon            =   "TBSBDETA.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4665
   ScaleWidth      =   8760
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      HelpContextID   =   1494
      Left            =   6240
      TabIndex        =   3
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   375
      HelpContextID   =   1494
      Left            =   7080
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calc"
      Height          =   375
      HelpContextID   =   1494
      Left            =   7080
      TabIndex        =   1
      Top             =   4200
      Width           =   735
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      DrawMode        =   14  'Copy Pen
      Height          =   255
      Left            =   5400
      ScaleHeight     =   195
      ScaleWidth      =   3195
      TabIndex        =   26
      Top             =   3600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Clos&e"
      Height          =   375
      HelpContextID   =   1494
      Left            =   7920
      TabIndex        =   0
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "&Plot"
      Height          =   375
      HelpContextID   =   1494
      Left            =   5400
      TabIndex        =   4
      Top             =   4200
      Width           =   735
   End
   Begin VB.Frame fraDefine 
      Caption         =   "Define"
      Height          =   3015
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   8535
      Begin VB.Frame fraContourLevels 
         Caption         =   "Deposition"
         Height          =   2775
         Left            =   5040
         TabIndex        =   36
         Top             =   120
         Width           =   3375
         Begin VB.Frame fraComponent 
            Caption         =   "Component"
            Height          =   975
            Left            =   1560
            TabIndex        =   39
            Top             =   1320
            Width           =   1695
            Begin VB.ComboBox cboDepComp 
               Height          =   315
               HelpContextID   =   1534
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Contour Levels"
            Height          =   2055
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1335
            Begin VB.CheckBox cbxAuto 
               Caption         =   "Auto"
               Height          =   195
               HelpContextID   =   1534
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtContourLevel 
               Height          =   285
               HelpContextID   =   1534
               Index           =   4
               Left            =   120
               TabIndex        =   16
               Top             =   1720
               Width           =   1095
            End
            Begin VB.TextBox txtContourLevel 
               Height          =   285
               HelpContextID   =   1534
               Index           =   3
               Left            =   120
               TabIndex        =   15
               Top             =   1410
               Width           =   1095
            End
            Begin VB.TextBox txtContourLevel 
               Height          =   285
               HelpContextID   =   1534
               Index           =   2
               Left            =   120
               TabIndex        =   14
               Top             =   1100
               Width           =   1095
            End
            Begin VB.TextBox txtContourLevel 
               Height          =   285
               HelpContextID   =   1534
               Index           =   1
               Left            =   120
               TabIndex        =   13
               Top             =   790
               Width           =   1095
            End
            Begin VB.TextBox txtContourLevel 
               Height          =   285
               HelpContextID   =   1534
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Top             =   480
               Width           =   1095
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Units"
            Height          =   1095
            Left            =   1560
            TabIndex        =   37
            Top             =   240
            Width           =   1695
            Begin VB.ComboBox cboUnitsDen 
               Height          =   315
               HelpContextID   =   1534
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   720
               Width           =   1215
            End
            Begin VB.ComboBox cboUnitsNum 
               Height          =   315
               HelpContextID   =   1534
               Left            =   240
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   240
               Width           =   1215
            End
            Begin VB.Line Line1 
               BorderWidth     =   3
               X1              =   240
               X2              =   1440
               Y1              =   645
               Y2              =   645
            End
         End
         Begin VB.Label lblDepMax 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1680
            TabIndex        =   41
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Maximum Deposition"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   2400
            Width           =   1455
         End
      End
      Begin VB.Frame fraFlightDir 
         Caption         =   "Flight Direction"
         Height          =   2775
         Left            =   2160
         TabIndex        =   28
         Top             =   120
         Width           =   2775
         Begin VB.HScrollBar hscFlightDir 
            Height          =   255
            HelpContextID   =   1533
            LargeChange     =   10
            Left            =   135
            Max             =   360
            TabIndex        =   10
            Top             =   2400
            Width           =   2535
         End
         Begin VB.PictureBox picFlightDir 
            Height          =   1815
            HelpContextID   =   1533
            Left            =   135
            ScaleHeight     =   1755
            ScaleWidth      =   2475
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txtFlightDir 
            Height          =   285
            HelpContextID   =   1533
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   8
            Top             =   270
            Width           =   855
         End
         Begin VB.Label lblFlightPath2 
            Caption         =   "deg"
            Height          =   255
            Left            =   2280
            TabIndex        =   29
            Top             =   300
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdACBound 
         Caption         =   "Area Coverage Boundary"
         Height          =   375
         HelpContextID   =   1497
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton cmdReceptors 
         Caption         =   "Discrete Receptors"
         Height          =   375
         HelpContextID   =   1496
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton cmdSBBound 
         Caption         =   "Spray Block Boundary"
         Height          =   375
         HelpContextID   =   1495
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Plotting Control"
      Height          =   1575
      Left            =   120
      TabIndex        =   30
      Top             =   3000
      Width           =   2415
      Begin VB.CheckBox cbxPlotIncludes 
         Caption         =   "Area Coverage Boundary"
         Height          =   255
         HelpContextID   =   1494
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox cbxPlotIncludes 
         Caption         =   "Contour Lines"
         Height          =   255
         HelpContextID   =   1494
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox cbxPlotIncludes 
         Caption         =   "Grid"
         Height          =   255
         HelpContextID   =   1494
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox cbxPlotIncludes 
         Caption         =   "Flight Lines"
         Height          =   255
         HelpContextID   =   1494
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   2175
      End
      Begin VB.CheckBox cbxPlotIncludes 
         Caption         =   "Spray Block Boundary"
         Height          =   255
         HelpContextID   =   1494
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraAreaCoverage 
      Caption         =   "Area Coverage"
      Height          =   1575
      Left            =   2640
      TabIndex        =   31
      Top             =   3000
      Width           =   2655
      Begin VB.Label lblACDepUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   1920
         TabIndex        =   43
         Top             =   720
         Width           =   330
      End
      Begin VB.Label lblACAreaUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   1920
         TabIndex        =   42
         Top             =   360
         Width           =   330
      End
      Begin VB.Label lblACResults 
         Caption         =   "Area"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   330
      End
      Begin VB.Label lblACResults 
         Alignment       =   2  'Center
         Caption         =   "Deposition Level"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblACArea 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblACAmount 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   32
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      Caption         =   "Calculating"
      Height          =   195
      Left            =   5400
      TabIndex        =   27
      Top             =   3360
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "frmTBSprayBlockDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TBDType
  BoundFlag As Boolean
  FltLinFlag As Boolean
  GridFlag As Boolean
  ContourFlag As Boolean
  ACBoundFlag As Boolean
  DepNum As Long         'Deposition Numerator Units
  DepDen As Long         'Deposition Denominator Units
  AutoCont As Long       'Automatic Contouring 0=off 1=on
  NumCont As Long        'Number of Contour Levels
  ValCont(4) As Single   'Contour Levels
  FlightDir As Single    'Flight Direction (deg)
  SMComp As Long         'Spray Material Component 0=unevap 1=nonvol 2=active
  NumBound As Long       'Number of Area Coverage Boundary Points
  BoundX(99) As Single   'Spray Block Boundary X (m)
  BoundY(99) As Single   'Spray Block Boundary Y (m)
  NumACBound As Long     'Number of Area Coverage Boundary Points
  ACBoundX(99) As Single 'Area Coverage Boundary X (m)
  ACBoundY(99) As Single 'Area Coverage Boundary Y (m)
  NumDisc As Long        'Number of Discrete Receptors
  DiscType(99) As Single 'Receptor Type 0=
  DiscX(99) As Single    'Receptor X Position (m)
  DiscY(99) As Single    'Receptor Y Position (m)
  DiscZ(99) As Single    'Receptor Z Position (m)
  DiscI(99) As Single    'Receptor Normal X component
  DiscJ(99) As Single    'Receptor Normal Y component
  DiscK(99) As Single    'Receptor Normal Z component
  DiscSize(99) As Single 'Receptor size (m)
  DiscDep(99) As Single  'Receptor Deposition
  NumFltLin As Long      'Number of Flight Lines
  FltLinXB(199) As Single 'Flight Line X begin (m)
  FltLinXE(199) As Single 'FLight Line X end (m)
  FltLinYB(199) As Single 'Flight Line Y begin (m)
  FltLinYE(199) As Single 'FLight Line Y end (m)
  NXgrid As Long         'Number of Grid X locations
  NYgrid As Long         'Number of Grid Y locations
  Xgrid() As Single      'Grid X locations (m)
  Ygrid() As Single      'Grid Y locations (m)
  Vgrid() As Single      'Grid values
  DepMax As Single       'Max Deposition in user units
  ACArea As Single       'Area Coverage Area (m2)
  ACCover As Single      'Avg Depos within area (user units)
End Type

Private PropTakeAction As Integer
Private OkToCalculate As Boolean  'used to cancel calcs in progress
Private CalcValid As Boolean    'flag for general calcs
Private CalcValidAC As Boolean  'flag for just the area coverage part of calcs
Private PlotValid As Boolean

Private TBD As TBDType

Private Sub Calculate()
'Perform calculations
  Dim ndiam As Long
  Dim iflg As Long
  ReDim Diam(3 * MAX_DROPS - 1) As Single
  ReDim Compl(3 * MAX_DROPS - 1) As Single
  Dim i As Integer
  Dim j As Integer
  Dim lngTmp() As Long
  
  If CalcValid And CalcValidAC Then Exit Sub 'nothing to do
  
  If Not CalcValid Then CalcValidAC = False 'force Area Coverage with regular calcs
  
  Me.MousePointer = vbHourglass
  cmdPlot.Enabled = False
  cmdExport.Enabled = False
  cmdCalc.Visible = False
  With cmdStop
    .Top = cmdCalc.Top
    .Left = cmdCalc.Left
    .Visible = True
  End With
   
  OkToCalculate = True  'stop if this goes false
    
  'init progress bar
  lblProgress.Caption = "Calculating"
  lblProgress.Visible = True
  picProgress.Visible = True
  UpdateProgressBar 0, 0
  
  If Not CalcValid Then
    'Sanity check on the spray block boundary
    If TBD.NumBound < 3 Then
      MsgBox "The Spray Block Boundary must contain at least " _
           & "three corner points for calculations to proceed.", vbCritical
      GoTo CalculateHalt
    End If
    
    'Check for nonsimple boundaries for spray block
    Call agsbck(CLng(TBD.NumBound), TBD.BoundX(0), TBD.BoundY(0), iflg)
    If iflg = 1 Then
      MsgBox "The Spray Block Boundary crosses itself " _
           & "and the calculations cannot proceed.", vbCritical
      GoTo CalculateHalt
    End If
  End If
  
  If Not CalcValidAC Then
    'Sanity check on the area coverage boundary
    If TBD.NumACBound > 0 And TBD.NumACBound < 3 Then
      MsgBox "The Area Coverage Boundary must contain at least " _
           & "three corner points for calculations to proceed.", vbCritical
      GoTo CalculateHalt
    End If
    
    'Check for nonsimple boundaries for area coverage
    If TBD.NumACBound > 0 Then
      Call agsbck(CLng(TBD.NumACBound), TBD.ACBoundX(0), TBD.ACBoundY(0), iflg)
      If iflg = 1 Then
        MsgBox "The Area Coverage Boundary crosses itself " _
             & "and the calculations cannot proceed.", vbCritical
        GoTo CalculateHalt
      End If
    End If
  End If
    
  'regular calcs
  If Not CalcValid Then
    'harvest the contour values
    'no need to order them
    For i = 0 To 4
      TBD.ValCont(i) = Val(txtContourLevel(i).Text)
    Next
   
    'Initialize calcs
    DoEvents: If Not OkToCalculate Then GoTo CalculateHalt
    Call agsbin(UD, _
      TBD.NumBound, TBD.BoundX(0), TBD.BoundY(0), _
      TBD.FlightDir, TBD.AutoCont, TBD.ValCont(0), _
      TBD.DepNum, TBD.DepDen, TBD.SMComp, _
      ndiam, Diam(0), Compl(0))
  
    'send discrete receptors
    '(convert types to long. they were single to keep EditGrid happy)
    ReDim lngTmp(TBD.NumDisc) 'don't do "- 1" in case there are zero receptors
    For i = 0 To TBD.NumDisc - 1: lngTmp(i) = TBD.DiscType(i): Next
    Call agdrin(TBD.NumDisc, lngTmp(0), _
                TBD.DiscX(0), TBD.DiscY(0), TBD.DiscZ(0), _
                TBD.DiscI(0), TBD.DiscJ(0), TBD.DiscK(0), _
                TBD.DiscSize(0))
    
    'Process all the drop categories
    'The compl array tells how much work has been done after
    'each drop category is complete.
    UpdateProgressBar 0, 1
    For i = 0 To ndiam - 1
      lblProgress.Caption = "Initial Drop Size: " & AGFormat$(Diam(i)) & " µm"
      DoEvents: If Not OkToCalculate Then GoTo CalculateHalt
      Call agdrop(i + 1)
      UpdateProgressBar Compl(i), Compl(ndiam - 1)
    Next
    
    lblProgress.Caption = "Completing calculations..."
    
    'recover discrete receptor deposition
    Call agdrot(TBD.DiscDep(0))
    
    'Retrieve the grid mesh dimensions and flight lines
    DoEvents: If Not OkToCalculate Then GoTo CalculateHalt
    Call agsgrd(TBD.NXgrid, TBD.NYgrid, _
      TBD.NumFltLin, TBD.FltLinYB(0), TBD.FltLinXB(0), TBD.FltLinXE(0))
    DoEvents
    CopyMemory TBD.FltLinYE(0), TBD.FltLinYB(0), TBD.NumFltLin * 4
    'Transform the flight lines for display
    For i = 0 To TBD.NumFltLin - 1
      agrtrn TBD.FltLinXB(i), TBD.FltLinYB(i)
      agrtrn TBD.FltLinXE(i), TBD.FltLinYE(i)
    Next
    
    'Recover Grid and contour values
    lblProgress.Caption = "Recovering deposition..."
    DoEvents: If Not OkToCalculate Then GoTo CalculateHalt
'tbc fix the indices
    ReDim TBD.Xgrid(1 To TBD.NXgrid)
    ReDim TBD.Ygrid(1 To TBD.NYgrid)
    ReDim TBD.Vgrid(1 To TBD.NXgrid, 1 To TBD.NYgrid)
'tbc fix indices
    Call agsend(TBD.NXgrid, TBD.NYgrid, _
      TBD.Xgrid(1), TBD.Ygrid(1), TBD.Vgrid(1, 1), TBD.ValCont(0), TBD.DepMax)
    
    'recover generated contour values and max deposition
    DoEvents: If Not OkToCalculate Then GoTo CalculateHalt
    PropTakeAction = False
    TBD.NumCont = 0
    For i = 0 To 4
      If TBD.ValCont(i) > 0 Then
        TBD.ValCont(TBD.NumCont) = TBD.ValCont(i)
        txtContourLevel(i).Text = AGFormat$(TBD.ValCont(i))
        TBD.NumCont = TBD.NumCont + 1
      Else
        txtContourLevel(i).Text = ""
      End If
    Next
    lblDepMax.Caption = AGFormat$(TBD.DepMax)
    PropTakeAction = True
  End If
  
  If Not CalcValidAC Then
    'Area Coverage
    If TBD.NumACBound > 0 Then
      lblProgress.Caption = "Area Coverage..."
      DoEvents: If Not OkToCalculate Then GoTo CalculateHalt
'tbc
'Open "C:\windows\desktop\vbarea.txt" For Output As #1
'Print #1, TBD.NXgrid; TBD.NYgrid
'For j = 1 To TBD.NYgrid
'  For i = 1 To TBD.NXgrid
'    Print #1, TBD.Xgrid(i); TBD.Ygrid(j); TBD.Vgrid(i, j)
'  Next
'Next
'Close #1
'tbc
      Call agarea(TBD.NXgrid, TBD.NYgrid, _
        TBD.Xgrid(1), TBD.Ygrid(1), TBD.Vgrid(1, 1), _
        TBD.NumACBound, TBD.ACBoundX(0), TBD.ACBoundY(0), _
        TBD.ACArea, TBD.ACCover)
      lblACArea = AGFormat$(UnitsDisplay(TBD.ACArea, UN_AREA))
      lblACAmount = AGFormat$(TBD.ACCover)
    Else
      lblACArea = ""
      lblACAmount = ""
    End If
  End If
  
  CalcValid = True  'Made it through the calcs.
  CalcValidAC = True
  PlotValid = False 'Plot geometry needs regeneration
    
CalculateHalt:
  'turn off progress bar
  lblProgress.Visible = False
  picProgress.Visible = False
    
  Me.MousePointer = vbDefault
  cmdPlot.Enabled = True
  cmdExport.Enabled = True
  cmdCalc.Visible = True
  cmdStop.Visible = False
End Sub

Private Sub ResetCalcs()
'Remove calculations and update the form
  CalcValid = False
  PlotValid = False
  lblDepMax.Caption = ""
  TBD.NumFltLin = 0 'erase flight lines
  ResetCalcsAC
End Sub

Private Sub ResetCalcsAC()
  CalcValidAC = False
  lblACArea = ""
  lblACAmount = ""
End Sub

Private Sub UpdateDirectionPicture(ctrl As Control)
'Update the graphical representation of Flight Direction
'and Wind Direction
  Dim FDrad As Integer
  Dim FDx As Integer
  Dim FDy As Integer
  Dim FDsin As Single
  Dim FDcos As Single
  Dim FDang As Single
  Dim FDtext As String
  Dim textX As Integer
  Dim textY As Integer
  
  FDrad = 400                     'circle/arrow radius
  FDx = ctrl.ScaleWidth / 2 + 100 'circle center X adjusted for text key
  FDy = ctrl.ScaleHeight / 2      'circle center Y
  
  'Setup drawing area
  ctrl.Cls
  ctrl.FontSize = 6
  ctrl.DrawWidth = 2
  
  'Color key
  ctrl.CurrentX = 0
  ctrl.CurrentY = 0
  ctrl.ForeColor = vbBlack
  ctrl.Print "North"
  ctrl.ForeColor = vbBlue
  ctrl.Print "Flight Direction"
  ctrl.ForeColor = vbRed
  ctrl.Print "Wind"

  'Flight Direction
  FDang = (TBD.FlightDir - 90) * 3.141592654 / 180
  FDsin = Sin(FDang)
  FDcos = Cos(FDang)
  ctrl.ForeColor = vbBlue
  GoSub DrawInsideArrow

  'North
  FDrad = 250
  FDang = (0 - 90) * 3.141592654 / 180
  FDsin = Sin(FDang)
  FDcos = Cos(FDang)
  ctrl.ForeColor = vbBlack
  GoSub DrawInsideArrow
  
  'Wind Direction
  FDrad = 400
  FDtext = "Wind"
  ctrl.ForeColor = vbRed
  FDang = (TBD.FlightDir + UD.MET.WD - 90) * 3.141592654 / 180
  FDsin = Sin(FDang)
  FDcos = Cos(FDang)
  GoSub DrawOutsideArrow

  Exit Sub

DrawInsideArrow:
  'draw an arrow inside the circle pointing out
  ctrl.Line (FDx - (FDrad * FDcos), FDy - (FDrad * FDsin))-(FDx + (FDrad * FDcos), FDy + (FDrad * FDsin))
  ctrl.Line (FDx + (FDrad * FDcos), FDy + (FDrad * FDsin))-(FDx + (FDrad * 0.5 * FDcos) + (FDrad * 0.25 * FDsin), FDy + (FDrad * 0.5 * FDsin) + (FDrad * 0.25 * -FDcos))
  ctrl.Line (FDx + (FDrad * FDcos), FDy + (FDrad * FDsin))-(FDx + (FDrad * 0.5 * FDcos) + (FDrad * 0.25 * -FDsin), FDy + (FDrad * 0.5 * FDsin) + (FDrad * 0.25 * FDcos))
  Return

DrawOutsideArrow:
  'draw an arrow outside the circle pointing in of length 200
  ctrl.Line (FDx + ((FDrad + 200) * FDcos), FDy + ((FDrad + 200) * FDsin))-(FDx + (FDrad * FDcos), FDy + (FDrad * FDsin))
  ctrl.Line (FDx + (FDrad * FDcos), FDy + (FDrad * FDsin))-(FDx + (FDrad * FDcos) + (200 * 0.5 * FDcos) + (200 * 0.25 * FDsin), FDy + (FDrad * FDsin) + (200 * 0.5 * FDsin) + (200 * 0.25 * -FDcos))
  ctrl.Line (FDx + (FDrad * FDcos), FDy + (FDrad * FDsin))-(FDx + (FDrad * FDcos) + (200 * 0.5 * FDcos) + (200 * 0.25 * -FDsin), FDy + (FDrad * FDsin) + (200 * 0.5 * FDsin) + (200 * 0.25 * FDcos))
  Return

DrawText:
  'draw text outside the circle
  If FDcos > 0.087 Then
    textX = FDx + (FDrad * FDcos)
  ElseIf FDcos < -0.087 Then
    textX = FDx + (FDrad * FDcos) - ctrl.TextWidth(FDtext)
  Else
    textX = FDx + (FDrad * FDcos) - 0.5 * ctrl.TextWidth(FDtext)
  End If
  If FDsin > 0.087 Then
    textY = FDy + (FDrad * FDsin)
  ElseIf FDsin < -0.087 Then
    textY = FDy + (FDrad * FDsin) - ctrl.TextHeight(FDtext)
  Else
    textY = FDy + (FDrad * FDsin) - 0.5 * ctrl.TextHeight(FDtext)
  End If
  ctrl.CurrentX = textX
  ctrl.CurrentY = textY
  ctrl.Print FDtext
  Return

End Sub

Private Sub UpdateProgressBar(curr_val As Single, max_val As Single)
'Update the Percent Completed progress bar
  Dim frac As Single
  Dim s As String
  Dim X As Single
  Dim pic As Control
  Dim SaveDrawMode As Integer
  
  Set pic = picProgress
  
  If max_val = 0 Then
    frac = 0
  Else
    frac = curr_val / max_val
  End If
  s = Format$(Int(frac * 100 + 0.5)) + "%"
  X = pic.Width * frac
  pic.Cls
  pic.DrawMode = 14 'Merge Pen Not
  pic.CurrentX = (pic.Width - pic.TextWidth(s)) / 2
  pic.CurrentY = (pic.Height - pic.TextHeight(s)) / 2
  pic.Print s
  pic.Line (0, 0)-(X, pic.Height), RGB(255, 0, 0), BF
  pic.Refresh
'
' If the window is iconized, treat the whole form as a progress bar
  If Me.WindowState = 1 Then
    X = Me.Width * frac
    Me.Cls
    Me.CurrentX = (Me.Width - Me.TextWidth(s$)) / 2
    Me.CurrentY = (Me.Height - Me.TextHeight(s$)) / 2
    SaveDrawMode = Me.DrawMode
    Me.DrawMode = 14 'Merge Pen Not
    Me.Print s$
    Me.Line (0, 0)-(X, Me.Height), QBColor(12), BF
    Me.DrawMode = SaveDrawMode
    DoEvents                    ' Allow other events.
  End If
End Sub

Private Sub UpdateACDepUnits()
'Update the Area Coverage Deposition units
  lblACDepUnits.Caption = cboUnitsNum.Text + "/" + cboUnitsDen.Text
End Sub

Private Sub cboDepComp_Click()
  TBD.SMComp = cboDepComp.ListIndex
  ResetCalcs
End Sub

Private Sub cboUnitsDen_Click()
  TBD.DepDen = cboUnitsDen.ListIndex
  UpdateACDepUnits
  ResetCalcs
End Sub

Private Sub cboUnitsNum_Click()
  TBD.DepNum = cboUnitsNum.ListIndex
  UpdateACDepUnits
  'If the numerator is drops, dim the component
  If cboUnitsNum.ListIndex = 0 Then 'drops
    cboDepComp.ListIndex = 2 'Unevaporated
    cboDepComp.Enabled = False
  Else
    cboDepComp.Enabled = True
  End If
  ResetCalcs
End Sub

Private Sub cbxAuto_Click()
  TBD.AutoCont = cbxAuto.Value
  ResetCalcs
End Sub

Private Sub cbxPlotIncludes_Click(Index As Integer)
  Select Case Index
  Case 0
    TBD.BoundFlag = CBool(cbxPlotIncludes(Index).Value)
  Case 1
    TBD.FltLinFlag = CBool(cbxPlotIncludes(Index).Value)
  Case 2
    TBD.GridFlag = CBool(cbxPlotIncludes(Index).Value)
  Case 3
    TBD.ContourFlag = CBool(cbxPlotIncludes(Index).Value)
  Case 4
    TBD.ACBoundFlag = CBool(cbxPlotIncludes(Index).Value)
  End Select
  PlotValid = False
End Sub

Private Sub cmdACBound_Click()
  Dim n As Integer
  
  Const MAX_BOUND = 100
  
  Load frmTBSBDBoundary
  With frmTBSBDBoundary
    .Caption = "Area Coverage Boundary"
    .SetHelpContextID 1497 'Area Coverage Boundary
    .eg.Setup .grdTable, .txtEditGrid, MAX_BOUND
    .eg.AddColumn "Corner X (" + UnitsName(UN_LENGTH) + ")"
    .eg.AddColumn "Corner Y (" + UnitsName(UN_LENGTH) + ")"
    .eg.Resize
    'populate the table with current values
    n = CInt(TBD.NumACBound)
    .eg.ArrayToGrid 1, n, TBD.ACBoundX(), UN_LENGTH
    .eg.ArrayToGrid 2, n, TBD.ACBoundY(), UN_LENGTH
    .Show vbModal
    If Not .Cancelled Then
      .eg.GridToArray 1, n, TBD.ACBoundX(), UN_LENGTH
      .eg.GridToArray 2, n, TBD.ACBoundY(), UN_LENGTH
      TBD.NumACBound = CLng(n)
      ResetCalcsAC
      If TBD.ACBoundFlag Then PlotValid = False 'Plot geometry has changed
    End If
  End With
  Unload frmTBSBDBoundary
End Sub

Private Sub cmdCalc_Click()
  Calculate
End Sub

Private Sub cmdExport_Click()
'Export Stream Assesment EXAMS data
'
  Dim fn As String
  Dim Msg As String
  Dim X As Single
  Dim Y As Single
  Dim i As Integer
  Dim j As Integer

  If Not CalcValid Then
    Select Case MsgBox( _
      "Calculations must be performed in order to export the Grid. " _
      & "Perform them now?", _
      vbYesNoCancel Or vbQuestion)
    Case vbYes
      Calculate
    End Select
  End If
  If Not CalcValid Then Exit Sub
  
  'Prompt for a file name
  If Not FileDialog(FD_SAVEAS, FD_TYPE_TEXT, fn) Then
    Exit Sub
  End If
  
  On Error GoTo ErrHandExportSprayBlockGrid
  Open fn For Output As #1

  'Print out the grid. Remember that the grid points are transformed,
  'so we need to send them through agtrn to make them normal again.
  Print #1, TBD.NXgrid; TBD.NYgrid
  For j = 1 To TBD.NYgrid
    For i = 1 To TBD.NXgrid
      X = TBD.Xgrid(i)
      Y = TBD.Ygrid(j)
      agrtrn X, Y
      Print #1, X; Y, TBD.Vgrid(i, j)
    Next
  Next
  
  Close #1
  Exit Sub
  
ErrHandExportSprayBlockGrid:
  Msg = "Error writing file: " + fn + vbCrLf + Error$(Err)
  MsgBox Msg, vbCritical + vbOKOnly
  Close #1
  Exit Sub
End Sub

Private Sub cmdPlot_Click()
'generate plot geometry and display
  Dim saveDataSource(4) As String
  Dim saveDataTitle(4) As String
  Dim i As Integer
  Dim j As Integer
  Dim slot As Integer
  Dim X As Single
  Dim Y As Single
  Dim XG() As Single
  Dim YG() As Single
  
  'Save sources and titles
  For i = 0 To 4: saveDataSource(i) = PlotGetDataSource(i): Next 'save existing
  For i = 0 To 4: saveDataTitle(i) = PlotGetDataTitle(i): Next 'save existing
  
  'See if Calcs need to be done
  If (TBD.GridFlag Or TBD.ContourFlag) And Not CalcValid Then
    Select Case MsgBox( _
      "Calculations must be performed in order to plot the Grid or " _
      & "the Contour Lines. Perform these calculations before continuing?", _
      vbYesNoCancel Or vbQuestion)
    Case vbYes
      Calculate
    Case vbNo
      'just keep going
    Case vbCancel
      Exit Sub
    End Select
  End If
  
  'Generate the plot geometry
  If Not PlotValid Then
    Me.MousePointer = vbHourglass
    cmdPlot.Enabled = False
    'init progress bar
    lblProgress.Caption = "Generating plot geometry..."
    lblProgress.Visible = True
    picProgress.Visible = True
    UpdateProgressBar 0, 0
  
    DoEvents
    PlotXYDataReset
    'Clear cources and titles
    For i = 0 To 4: PlotSetDataSource i, "": Next 'clear
    For i = 0 To 4: PlotSetDataTitle i, "": Next 'clear
    
    'Spray Block Boundary
    If TBD.BoundFlag And TBD.NumBound > 1 Then
      slot = 0
      PlotDataAddPoint slot, UnitsDisplay(TBD.BoundX(0), UN_LENGTH), _
                             UnitsDisplay(TBD.BoundY(0), UN_LENGTH), GQ_MOVETO
      For i = 1 To TBD.NumBound - 1
        PlotDataAddPoint slot, UnitsDisplay(TBD.BoundX(i), UN_LENGTH), _
                               UnitsDisplay(TBD.BoundY(i), UN_LENGTH), GQ_DRAWTO
      Next
      PlotDataAddPoint slot, UnitsDisplay(TBD.BoundX(0), UN_LENGTH), _
                             UnitsDisplay(TBD.BoundY(0), UN_LENGTH), GQ_DRAWTO
    End If
    
    'Flight Lines
    If TBD.FltLinFlag Then
      slot = 0
      For i = 0 To TBD.NumFltLin - 1
        PlotDataAddPoint slot, UnitsDisplay(TBD.FltLinXB(i), UN_LENGTH), _
                               UnitsDisplay(TBD.FltLinYB(i), UN_LENGTH), GQ_MOVETO
        PlotDataAddPoint slot, UnitsDisplay(TBD.FltLinXE(i), UN_LENGTH), _
                               UnitsDisplay(TBD.FltLinYE(i), UN_LENGTH), GQ_DRAWTO
      Next
    End If
    
    'Grid
    If TBD.GridFlag And CalcValid Then
      slot = 0
      For i = 1 To TBD.NXgrid
        X = TBD.Xgrid(i)
        Y = TBD.Ygrid(1)
        agrtrn X, Y
        PlotDataAddPoint slot, UnitsDisplay(X, UN_LENGTH), _
                               UnitsDisplay(Y, UN_LENGTH), GQ_MOVETO
        X = TBD.Xgrid(i)
        Y = TBD.Ygrid(TBD.NYgrid)
        agrtrn X, Y
        PlotDataAddPoint slot, UnitsDisplay(X, UN_LENGTH), _
                               UnitsDisplay(Y, UN_LENGTH), GQ_DRAWTO
      Next
      For i = 1 To TBD.NYgrid
        X = TBD.Xgrid(1)
        Y = TBD.Ygrid(i)
        agrtrn X, Y
        PlotDataAddPoint slot, UnitsDisplay(X, UN_LENGTH), _
                               UnitsDisplay(Y, UN_LENGTH), GQ_MOVETO
        X = TBD.Xgrid(TBD.NXgrid)
        Y = TBD.Ygrid(i)
        agrtrn X, Y
        PlotDataAddPoint slot, UnitsDisplay(X, UN_LENGTH), _
                               UnitsDisplay(Y, UN_LENGTH), GQ_DRAWTO
      Next
    End If
    
    'Contours
    If TBD.ContourFlag And CalcValid Then
      'harvest the contour values
      TBD.NumCont = 0
      For i = 0 To 4
        If Trim$(txtContourLevel(i).Text) <> "" Then
          TBD.ValCont(TBD.NumCont) = Val(txtContourLevel(i).Text)
          TBD.NumCont = TBD.NumCont + 1
        End If
      Next
      'units conversion
      ReDim XG(1 To TBD.NXgrid)
      ReDim YG(1 To TBD.NYgrid)
      For i = 1 To TBD.NXgrid
        XG(i) = UnitsDisplay(TBD.Xgrid(i), UN_LENGTH)
      Next
      For i = 1 To TBD.NYgrid
        YG(i) = UnitsDisplay(TBD.Ygrid(i), UN_LENGTH)
      Next
      'isoplq does the tranforations
      isoplq TBD.NumCont, TBD.ValCont(), CInt(TBD.NXgrid), CInt(TBD.NYgrid), _
        XG(), YG(), TBD.Vgrid()
    End If
    
    'Area Coverage Boundary
    If TBD.ACBoundFlag And TBD.NumACBound > 1 Then
      slot = 0
      PlotDataAddPoint slot, UnitsDisplay(TBD.ACBoundX(0), UN_LENGTH), _
                             UnitsDisplay(TBD.ACBoundY(0), UN_LENGTH), GQ_MOVETO
      For i = 1 To TBD.NumACBound - 1
        PlotDataAddPoint slot, UnitsDisplay(TBD.ACBoundX(i), UN_LENGTH), _
                               UnitsDisplay(TBD.ACBoundY(i), UN_LENGTH), GQ_DRAWTO
      Next
      PlotDataAddPoint slot, UnitsDisplay(TBD.ACBoundX(0), UN_LENGTH), _
                             UnitsDisplay(TBD.ACBoundY(0), UN_LENGTH), GQ_DRAWTO
    End If
    
    'turn off progress bar
    lblProgress.Visible = False
    picProgress.Visible = False
    
    PlotValid = True
    Me.MousePointer = vbDefault
    cmdPlot.Enabled = True
  End If

  'Display the completed plot
  If PlotValid Then
    ShowPlot PV_SBDET
  End If
  'Restore sorces and titles
  For i = 0 To 4: PlotSetDataSource i, saveDataSource(i): Next
  For i = 0 To 4: PlotSetDataTitle i, saveDataTitle(i): Next
End Sub

Private Sub cmdOk_Click()
  'Display a warning message
  If CalcValid Or CalcValidAC Then
    If MsgBox("None of the information generated in the " + _
              "toolbox calculations will be saved on exit. " + _
              "Exit the toolbox?", _
              vbOKCancel + vbExclamation) = vbCancel Then Exit Sub
  End If
  
  OkToCalculate = False 'set the flag that stops the calcs
  Unload Me
End Sub

Private Sub cmdReceptors_Click()
  Dim n As Integer
  Dim Xtmp(99)
  
  Const MAX_RECEP = 100
  
  Load frmTBSBDReceptors
  With frmTBSBDReceptors
    'set up the table and its columns
    .eg.Setup .grdTable, .txtEditGrid, MAX_RECEP
    .eg.AddColumn "Type"
    .eg.AddColumn "X Location (" + UnitsName(UN_LENGTH) + ")"
    .eg.AddColumn "Y Location (" + UnitsName(UN_LENGTH) + ")"
    .eg.AddColumn "Z Location (" + UnitsName(UN_LENGTH) + ")"
    .eg.AddColumn "X Normal"
    .eg.AddColumn "Y Normal"
    .eg.AddColumn "Z Normal"
    .eg.AddColumn "Size (" + UnitsName(UN_SMLENGTH) + ")"
    .eg.AddColumn "Deposition (" + _
                  cboUnitsNum.Text + "/" + cboUnitsDen.Text + _
                  ")", False
    .eg.Resize
    'populate the table with current values
    n = CInt(TBD.NumDisc)
    .eg.ArrayToGrid 1, n, TBD.DiscType()
    .eg.ArrayToGrid 2, n, TBD.DiscX(), UN_LENGTH
    .eg.ArrayToGrid 3, n, TBD.DiscY(), UN_LENGTH
    .eg.ArrayToGrid 4, n, TBD.DiscZ(), UN_LENGTH
    .eg.ArrayToGrid 5, n, TBD.DiscI()
    .eg.ArrayToGrid 6, n, TBD.DiscJ()
    .eg.ArrayToGrid 7, n, TBD.DiscK()
    .eg.ArrayToGrid 8, n, TBD.DiscSize(), UN_SMLENGTH
    'The last column displays Deposition
    If CalcValid Then
      .eg.ArrayToGrid 9, n, TBD.DiscDep()
    End If
    'show the Discrete Receptor form
    .Show vbModal
    'Harvest the new values
    If Not .Cancelled Then
      .eg.GridToArray 1, n, TBD.DiscType()
      .eg.GridToArray 2, n, TBD.DiscX(), UN_LENGTH
      .eg.GridToArray 3, n, TBD.DiscY(), UN_LENGTH
      .eg.GridToArray 4, n, TBD.DiscZ(), UN_LENGTH
      .eg.GridToArray 5, n, TBD.DiscI()
      .eg.GridToArray 6, n, TBD.DiscJ()
      .eg.GridToArray 7, n, TBD.DiscK()
      .eg.GridToArray 8, n, TBD.DiscSize(), UN_SMLENGTH
      TBD.NumDisc = CLng(n)
      ResetCalcs
    End If
  End With
  Unload frmTBSBDReceptors
End Sub

Private Sub cmdSBBound_Click()
  Dim n As Integer
  
  Const MAX_BOUND = 100
  
  Load frmTBSBDBoundary
  With frmTBSBDBoundary
    .Caption = "Spray Block Boundary"
    .SetHelpContextID 1495 'Spray Block Boundary
    .eg.Setup .grdTable, .txtEditGrid, MAX_BOUND
    .eg.AddColumn "Corner X (" + UnitsName(UN_LENGTH) + ")"
    .eg.AddColumn "Corner Y (" + UnitsName(UN_LENGTH) + ")"
    .eg.Resize
    'populate the table with current values
    n = CInt(TBD.NumBound)
    .eg.ArrayToGrid 1, n, TBD.BoundX(), UN_LENGTH
    .eg.ArrayToGrid 2, n, TBD.BoundY(), UN_LENGTH
    'show the form
    .Show vbModal
    'Harvest the edited values
    If Not .Cancelled Then
      .eg.GridToArray 1, n, TBD.BoundX(), UN_LENGTH
      .eg.GridToArray 2, n, TBD.BoundY(), UN_LENGTH
      TBD.NumBound = CLng(n)
      ResetCalcs
    End If
  End With
  Unload frmTBSBDBoundary
End Sub

Private Sub cmdStop_Click()
  OkToCalculate = False 'set the flag that stops the calcs
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim c As Control
  
  CenterForm Me
  PropTakeAction = False
  
  cboUnitsNum.AddItem "drops"
  cboUnitsNum.AddItem "ozf"
  cboUnitsNum.AddItem "gal"
  cboUnitsNum.AddItem "lbm"
  cboUnitsNum.AddItem "l"
  cboUnitsNum.AddItem "g"
  cboUnitsNum.AddItem "kg"

  cboUnitsDen.AddItem "in²"
  cboUnitsDen.AddItem "ft²"
  cboUnitsDen.AddItem "ac"
  cboUnitsDen.AddItem "cm²"
  cboUnitsDen.AddItem "m²"
  cboUnitsDen.AddItem "ha"

  cboDepComp.AddItem "Active"
  cboDepComp.AddItem "Nonvolatile"
  cboDepComp.AddItem "Unevaporated"
  
  'default settings
  With TBD
    .BoundFlag = True
    .FltLinFlag = True
    .ContourFlag = True
    .AutoCont = True
    .DepNum = 1
    .DepDen = 1
    .SMComp = 2
    .NumBound = 4
    .BoundX(0) = 0: .BoundY(0) = 0
    .BoundX(1) = 100: .BoundY(1) = 0
    .BoundX(2) = 100: .BoundY(2) = 100
    .BoundX(3) = 0: .BoundY(3) = 100
  End With
  
  'Match Controls to Data
  'Flight Dir scroll bar
  hscFlightDir.Value = TBD.FlightDir
  'Flight Dir text
  txtFlightDir.Text = CStr(TBD.FlightDir)
  'Auto Contours
  cbxAuto.Value = -TBD.AutoCont
  'Contours
  For i = 0 To 4
    If TBD.ValCont(i) > 0 Then
      txtContourLevel(i).Text = AGFormat$(TBD.ValCont(i))
    Else
      txtContourLevel(i).Text = ""
    End If
  Next
  'Units
  cboUnitsNum.ListIndex = TBD.DepNum
  cboUnitsDen.ListIndex = TBD.DepDen
  cboDepComp.ListIndex = TBD.SMComp
  UpdateACDepUnits
  'Plotting control
  cbxPlotIncludes(0).Value = -TBD.BoundFlag
  cbxPlotIncludes(1).Value = -TBD.FltLinFlag
  cbxPlotIncludes(2).Value = -TBD.GridFlag
  cbxPlotIncludes(3).Value = -TBD.ContourFlag
  cbxPlotIncludes(4).Value = -TBD.ACBoundFlag
  'Area Coverage results
  lblACAreaUnits = UnitsName(UN_AREA)
  
  PropTakeAction = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  OkToCalculate = False 'set the flag that stops the calcs
End Sub

Private Sub hscFlightDir_Change()
  If PropTakeAction Then
    PropTakeAction = False
    TBD.FlightDir = hscFlightDir.Value
    'redraw the picture box
    UpdateDirectionPicture picFlightDir
    'update the text
    If Not (TBD.FlightDir = 0 And txtFlightDir.Text = "") Then
      txtFlightDir.Text = TBD.FlightDir
    End If
    ResetCalcs
    PropTakeAction = True
  End If
End Sub

Private Sub hscFlightDir_Scroll()
  If PropTakeAction Then
    PropTakeAction = False
    TBD.FlightDir = hscFlightDir.Value
    'redraw the picture box
    UpdateDirectionPicture picFlightDir
    'update the text
    If Not (TBD.FlightDir = 0 And txtFlightDir.Text = "") Then
      txtFlightDir.Text = TBD.FlightDir
    End If
    ResetCalcs
    PropTakeAction = True
  End If
End Sub

Private Sub picFlightDir_Paint()
  'redraw the picture box
  UpdateDirectionPicture picFlightDir
End Sub

Private Sub txtContourLevel_Change(Index As Integer)
  If PropTakeAction Then
    PlotValid = False
  End If
End Sub

Private Sub txtFlightDir_KeyPress(KeyAscii As Integer)
  If KeyAscii = Asc(vbCr) Then
    txtFlightDir_LostFocus 'Update the flight direction
    KeyAscii = 0
  End If
End Sub

Private Sub txtFlightDir_LostFocus()
  If PropTakeAction Then
    PropTakeAction = False
    If Val(txtFlightDir.Text) < 0 Then txtFlightDir.Text = "0"
    If Val(txtFlightDir.Text) > 360 Then txtFlightDir.Text = "360)"
    TBD.FlightDir = Val(txtFlightDir.Text)
    'redraw the picture box
    UpdateDirectionPicture picFlightDir
    'adjust the scroll bar
    hscFlightDir.Value = TBD.FlightDir
    ResetCalcs
    PropTakeAction = True
  End If
End Sub
