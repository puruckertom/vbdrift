VERSION 5.00
Begin VB.Form frmTBStream 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stream Assessment"
   ClientHeight    =   6210
   ClientLeft      =   1155
   ClientTop       =   1710
   ClientWidth     =   7980
   HelpContextID   =   1447
   Icon            =   "TBSTREAM.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExams 
      Caption         =   "E&XAMS"
      Height          =   375
      HelpContextID   =   1447
      Left            =   5400
      TabIndex        =   2
      Top             =   5760
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Geometry"
      Height          =   3735
      Left            =   120
      TabIndex        =   34
      Top             =   0
      Width           =   7815
      Begin VB.Frame fraSprayLines 
         Caption         =   "Spray Block"
         Height          =   1335
         Left            =   120
         TabIndex        =   84
         Top             =   840
         Width           =   2415
         Begin VB.TextBox txtSprayLine 
            Height          =   285
            HelpContextID   =   1513
            Left            =   1200
            TabIndex        =   5
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtTurnTime 
            Height          =   285
            HelpContextID   =   1514
            Left            =   1200
            TabIndex        =   6
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Spray Line Length"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   88
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblSprayLineUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Left            =   2040
            TabIndex        =   87
            Top             =   360
            Width           =   330
         End
         Begin VB.Label lblTurnTimeUnits 
            AutoSize        =   -1  'True
            Caption         =   "sec"
            Height          =   195
            Left            =   2040
            TabIndex        =   86
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label 
            Caption         =   "Turn-Around Time"
            Height          =   375
            Index           =   22
            Left            =   120
            TabIndex        =   85
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame fraStream 
         Caption         =   "Stream"
         Height          =   1815
         Left            =   5280
         TabIndex        =   74
         Top             =   360
         Width           =   2415
         Begin VB.TextBox txtStreamWidth 
            Height          =   285
            HelpContextID   =   1515
            Left            =   1080
            TabIndex        =   7
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtStreamDepth 
            Height          =   285
            HelpContextID   =   1516
            Left            =   1080
            TabIndex        =   8
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtStreamRate 
            Height          =   285
            HelpContextID   =   1517
            Left            =   1080
            TabIndex        =   9
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label 
            Caption         =   "Width"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   83
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblStreamWidthUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Left            =   1920
            TabIndex        =   82
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label 
            Caption         =   "Depth"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   81
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblStreamDepthUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Left            =   1920
            TabIndex        =   80
            Top             =   720
            Width           =   330
         End
         Begin VB.Label Label 
            Caption         =   "Flow Rate"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   79
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblStreamRateUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Left            =   1920
            TabIndex        =   78
            Top             =   1080
            Width           =   330
         End
         Begin VB.Label lblStreamSpeedUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Left            =   1920
            TabIndex        =   77
            Top             =   1440
            Width           =   330
         End
         Begin VB.Label Label 
            Caption         =   "Flow Speed"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   76
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblStreamSpeed 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1080
            TabIndex        =   75
            Top             =   1440
            Width           =   735
         End
      End
      Begin VB.TextBox txtRechargeRate 
         Height          =   285
         HelpContextID   =   1521
         Left            =   2400
         TabIndex        =   12
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtDecayFactor 
         Height          =   285
         HelpContextID   =   1520
         Left            =   2400
         TabIndex        =   11
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtRiparian 
         Height          =   285
         HelpContextID   =   1519
         Left            =   2400
         TabIndex        =   10
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtStreamDist 
         Height          =   285
         HelpContextID   =   1518
         Left            =   6600
         TabIndex        =   13
         Top             =   2640
         Width           =   735
      End
      Begin VB.Line Line12 
         X1              =   2640
         X2              =   2760
         Y1              =   1200
         Y2              =   1320
      End
      Begin VB.Line Line11 
         X1              =   2760
         X2              =   2640
         Y1              =   1320
         Y2              =   1440
      End
      Begin VB.Line Line10 
         X1              =   2760
         X2              =   2520
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line9 
         X1              =   5280
         X2              =   5160
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line8 
         X1              =   5160
         X2              =   5160
         Y1              =   600
         Y2              =   840
      End
      Begin VB.Line Line7 
         X1              =   4200
         X2              =   4080
         Y1              =   2400
         Y2              =   2280
      End
      Begin VB.Line Line6 
         X1              =   4080
         X2              =   3960
         Y1              =   2280
         Y2              =   2400
      End
      Begin VB.Line Line5 
         X1              =   4080
         X2              =   4080
         Y1              =   2280
         Y2              =   2640
      End
      Begin VB.Line Line4 
         X1              =   4680
         X2              =   4080
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label 
         Caption         =   "Recharge Rate"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   47
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label lblRechargeRateUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3240
         TabIndex        =   46
         Top             =   3240
         Width           =   330
      End
      Begin VB.Label lblDecayFactorUnits 
         AutoSize        =   -1  'True
         Caption         =   "1/day"
         Height          =   195
         Left            =   3240
         TabIndex        =   42
         Top             =   2880
         Width           =   420
      End
      Begin VB.Label Label 
         Caption         =   "Instream Chemical Decay Rate"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   41
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label 
         Caption         =   "Riparian Interception Factor"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   40
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label lblStreamDistUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   7440
         TabIndex        =   38
         Top             =   2640
         Width           =   330
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Distance from edge of application area to center of stream"
         Height          =   615
         Index           =   6
         Left            =   4800
         TabIndex        =   37
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Line Line3 
         Index           =   23
         X1              =   4560
         X2              =   4680
         Y1              =   2040
         Y2              =   2160
      End
      Begin VB.Line Line3 
         Index           =   22
         X1              =   3600
         X2              =   3720
         Y1              =   2160
         Y2              =   2280
      End
      Begin VB.Line Line3 
         Index           =   21
         X1              =   3720
         X2              =   3600
         Y1              =   2040
         Y2              =   2160
      End
      Begin VB.Line Line3 
         Index           =   20
         X1              =   4680
         X2              =   4560
         Y1              =   2160
         Y2              =   2280
      End
      Begin VB.Line Line3 
         Index           =   18
         X1              =   4680
         X2              =   3600
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line3 
         Index           =   17
         X1              =   4680
         X2              =   4680
         Y1              =   2040
         Y2              =   2280
      End
      Begin VB.Line Line3 
         Index           =   16
         X1              =   3600
         X2              =   3600
         Y1              =   2040
         Y2              =   2280
      End
      Begin VB.Label Label 
         Caption         =   "Stream"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label 
         Caption         =   "Spray Block"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   35
         Top             =   600
         Width           =   855
      End
      Begin VB.Line Line3 
         Index           =   15
         X1              =   2760
         X2              =   2880
         Y1              =   1800
         Y2              =   1920
      End
      Begin VB.Line Line3 
         Index           =   14
         X1              =   2880
         X2              =   3000
         Y1              =   960
         Y2              =   1080
      End
      Begin VB.Line Line3 
         Index           =   13
         X1              =   2880
         X2              =   2760
         Y1              =   960
         Y2              =   1080
      End
      Begin VB.Line Line3 
         Index           =   12
         X1              =   3000
         X2              =   2880
         Y1              =   1800
         Y2              =   1920
      End
      Begin VB.Line Line3 
         Index           =   10
         X1              =   2880
         X2              =   2880
         Y1              =   1920
         Y2              =   960
      End
      Begin VB.Line Line3 
         Index           =   9
         X1              =   2760
         X2              =   3000
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line3 
         Index           =   8
         X1              =   2760
         X2              =   3000
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line3 
         Index           =   7
         X1              =   4440
         X2              =   4560
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Line Line3 
         Index           =   6
         X1              =   4800
         X2              =   4920
         Y1              =   600
         Y2              =   720
      End
      Begin VB.Line Line3 
         Index           =   5
         X1              =   4920
         X2              =   4800
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Line Line3 
         Index           =   4
         X1              =   4560
         X2              =   4440
         Y1              =   600
         Y2              =   720
      End
      Begin VB.Line Line3 
         Index           =   3
         X1              =   5160
         X2              =   4800
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   4560
         X2              =   4200
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   1
         X1              =   4800
         X2              =   4800
         Y1              =   480
         Y2              =   2400
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         Index           =   0
         X1              =   4560
         X2              =   4560
         Y1              =   480
         Y2              =   2400
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   4
         X1              =   3600
         X2              =   3600
         Y1              =   960
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   3
         X1              =   3480
         X2              =   3480
         Y1              =   960
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   2
         X1              =   3360
         X2              =   3360
         Y1              =   960
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   3240
         X2              =   3240
         Y1              =   960
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   3120
         X2              =   3120
         Y1              =   960
         Y2              =   1920
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   3120
         Top             =   960
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   1935
         Left            =   4560
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Clos&e"
      Height          =   375
      HelpContextID   =   1447
      Left            =   7080
      TabIndex        =   0
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calc"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1447
      Left            =   6240
      TabIndex        =   1
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      HelpContextID   =   1447
      Left            =   4560
      TabIndex        =   3
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "Plo&t"
      Height          =   375
      HelpContextID   =   1447
      Left            =   3720
      TabIndex        =   4
      Top             =   5760
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control"
      Height          =   1815
      Left            =   120
      TabIndex        =   39
      Top             =   3720
      Width           =   7815
      Begin VB.PictureBox picCalcType 
         Height          =   1215
         Index           =   2
         Left            =   2160
         ScaleHeight     =   1155
         ScaleWidth      =   5835
         TabIndex        =   51
         Top             =   480
         Width           =   5895
         Begin VB.CheckBox cbxAuto 
            Caption         =   "Automatically set distance values"
            Height          =   255
            HelpContextID   =   1447
            Index           =   2
            Left            =   2040
            TabIndex        =   32
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox txtDistTime 
            Height          =   285
            HelpContextID   =   1447
            Index           =   1
            Left            =   600
            TabIndex        =   27
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtDistTime 
            Height          =   285
            HelpContextID   =   1447
            Index           =   0
            Left            =   600
            TabIndex        =   26
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtDistDist 
            Height          =   285
            HelpContextID   =   1447
            Index           =   0
            Left            =   2040
            TabIndex        =   28
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtDistDist 
            Height          =   285
            HelpContextID   =   1447
            Index           =   1
            Left            =   2880
            TabIndex        =   29
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtDistDist 
            Height          =   285
            HelpContextID   =   1447
            Index           =   2
            Left            =   3720
            TabIndex        =   30
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtDistDist 
            Height          =   285
            HelpContextID   =   1447
            Index           =   3
            Left            =   4560
            TabIndex        =   31
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "End:"
            Height          =   195
            Index           =   18
            Left            =   240
            TabIndex        =   73
            Top             =   720
            Width           =   330
         End
         Begin VB.Label lblDistTimeUnits 
            AutoSize        =   -1  'True
            Caption         =   "sec"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   72
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Begin:"
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   71
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label 
            Caption         =   "Time"
            Height          =   255
            Index           =   16
            Left            =   840
            TabIndex        =   70
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblDistTimeUnits 
            AutoSize        =   -1  'True
            Caption         =   "sec"
            Height          =   195
            Index           =   0
            Left            =   1440
            TabIndex        =   69
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label 
            Caption         =   "Downstream Distance(s)"
            Height          =   255
            Index           =   19
            Left            =   2040
            TabIndex        =   68
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label lblDistDistUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Left            =   5400
            TabIndex        =   67
            Top             =   540
            Width           =   330
         End
      End
      Begin VB.PictureBox picCalcType 
         Height          =   1215
         Index           =   1
         Left            =   840
         ScaleHeight     =   1155
         ScaleWidth      =   5835
         TabIndex        =   50
         Top             =   480
         Width           =   5895
         Begin VB.CheckBox cbxAuto 
            Caption         =   "Automatically set time values"
            Height          =   255
            HelpContextID   =   1447
            Index           =   1
            Left            =   2040
            TabIndex        =   25
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox txtTimeDist 
            Height          =   285
            HelpContextID   =   1447
            Index           =   1
            Left            =   600
            TabIndex        =   20
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtTimeTime 
            Height          =   285
            HelpContextID   =   1447
            Index           =   3
            Left            =   4560
            TabIndex        =   24
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtTimeTime 
            Height          =   285
            HelpContextID   =   1447
            Index           =   2
            Left            =   3720
            TabIndex        =   23
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtTimeTime 
            Height          =   285
            HelpContextID   =   1447
            Index           =   1
            Left            =   2880
            TabIndex        =   22
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtTimeTime 
            Height          =   285
            HelpContextID   =   1447
            Index           =   0
            Left            =   2040
            TabIndex        =   21
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtTimeDist 
            Height          =   285
            HelpContextID   =   1447
            Index           =   0
            Left            =   600
            TabIndex        =   19
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "End:"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   66
            Top             =   720
            Width           =   330
         End
         Begin VB.Label lblTimeDistUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   65
            Top             =   720
            Width           =   330
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Begin:"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblTImeTimeUnits 
            AutoSize        =   -1  'True
            Caption         =   "sec"
            Height          =   195
            Left            =   5400
            TabIndex        =   63
            Top             =   540
            Width           =   255
         End
         Begin VB.Label Label 
            Caption         =   "Time(s) "
            Height          =   255
            Index           =   15
            Left            =   2040
            TabIndex        =   62
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lblTimeDistUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Index           =   0
            Left            =   1440
            TabIndex        =   61
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label 
            Caption         =   "Downstream Distance"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   60
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.PictureBox picCalcType 
         Height          =   975
         Index           =   0
         Left            =   240
         ScaleHeight     =   915
         ScaleWidth      =   6795
         TabIndex        =   49
         Top             =   480
         Width           =   6855
         Begin VB.TextBox txtSingle 
            Height          =   285
            HelpContextID   =   1447
            Index           =   1
            Left            =   2760
            TabIndex        =   18
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtSingle 
            Height          =   285
            HelpContextID   =   1447
            Index           =   0
            Left            =   600
            TabIndex        =   17
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblSingleConc 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5160
            TabIndex        =   59
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblSingleConcUnits 
            AutoSize        =   -1  'True
            Caption         =   "ng/L (ppt)"
            Height          =   195
            Left            =   6000
            TabIndex        =   58
            Top             =   480
            Width           =   705
         End
         Begin VB.Label Label 
            Caption         =   "Peak Conc.:"
            Height          =   255
            Index           =   11
            Left            =   4200
            TabIndex        =   57
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblSingleDistUnits 
            AutoSize        =   -1  'True
            Caption         =   "units"
            Height          =   195
            Left            =   3600
            TabIndex        =   56
            Top             =   480
            Width           =   330
         End
         Begin VB.Label Label 
            Caption         =   "Distance:"
            Height          =   255
            Index           =   10
            Left            =   2040
            TabIndex        =   55
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblSingleTimeUnits 
            AutoSize        =   -1  'True
            Caption         =   "sec"
            Height          =   195
            Left            =   1440
            TabIndex        =   54
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label 
            Caption         =   "Time:"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   53
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label 
            Caption         =   "Provide one value and the others will be calculated."
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.OptionButton optCalcType 
         Caption         =   "a single point."
         Height          =   255
         HelpContextID   =   1447
         Index           =   0
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optCalcType 
         Caption         =   "given distance(s)"
         Height          =   255
         HelpContextID   =   1447
         Index           =   2
         Left            =   4320
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optCalcType 
         Caption         =   "given time(s)"
         Height          =   255
         HelpContextID   =   1447
         Index           =   1
         Left            =   3000
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Calculate results at:"
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   48
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraTier1 
      Caption         =   "Tier I Settings"
      Height          =   615
      Left            =   120
      TabIndex        =   43
      Top             =   5520
      Width           =   3135
      Begin VB.TextBox txtActiveRate 
         Height          =   285
         HelpContextID   =   1010
         Left            =   1440
         TabIndex        =   33
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lblActiveRateUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2385
         TabIndex        =   45
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lblActiveRate 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Active Rate:"
         Height          =   195
         Left            =   360
         TabIndex        =   44
         Top             =   240
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmTBStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tbstream.frm,v 1.11 2002/08/06 20:00:59 tom Exp $
Private PropTakeAction As Integer
Private NeedCalcs As Integer     'tracks calculation status
Private CalcOutputMarker As Integer 'tracks output selection
Private PreviousUnits As Integer 'Tracks units setting

Private NEX As Long          'EXAMS array size
Private EXTIME(49) As Single 'EXAMS Time array
Private EXDIST(49) As Single 'EXAMS Distance array
Private EXCONC(49) As Single 'EXAMS Initial Concentration array
Private Nauto As Long
Private Xauto(3) As Single
Private Rauto(3) As Single

Private Sub cbxAuto_Click(Index As Integer)
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
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
'Export Stream Assesment EXAMS data
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
       
  AppendStr hdr, Format$("Stream Width (" & lblStreamWidthUnits & "):", c1fmt), False
  AppendStr hdr, Format$(txtStreamWidth, c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Stream Depth (" & lblStreamDepthUnits & "):", c1fmt), False
  AppendStr hdr, Format$(txtStreamDepth, c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Stream Flow Rate (" & lblStreamRateUnits & "):", c1fmt), False
  AppendStr hdr, Format$(txtStreamRate, c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Spray Line Length (" & lblSprayLineUnits & "):", c1fmt), False
  AppendStr hdr, Format$(txtSprayLine, c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Instream Chemical Decay Rate (" & lblDecayFactorUnits & "):", c1fmt), False
  AppendStr hdr, Format$(txtDecayFactor, c2fmt), False
  AppendStr hdr, "", True
  AppendStr hdr, Format$("Recharge Rate (" & lblRechargeRateUnits & "):", c1fmt), False
  AppendStr hdr, Format$(txtRechargeRate, c2fmt), False
  AppendStr hdr, "", True
    
  Print #1, hdr

  'Part II: Initial Conditions
  '
  'set up the formats for the columns
  c1wid = 3  'column 1
  c2wid = 15  'column 2
  c1fmt = String$(c1wid, "@") 'right-justified
  c2fmt = " " & String$(c2wid, "@") '1 space, right-justified
  Print #1, "  #      Time (sec)    Distance (m) Initial Concentration (ng/L)(ppt)"
  For i = 0 To NEX - 1
    Print #1, Format$(i + 1, c1fmt) + _
              Format$(EXTIME(i), c2fmt) + _
              Format$(EXDIST(i), c2fmt) + _
              Format$(EXCONC(i), c2fmt)
  Next
  Print #1,

  'Part III: actual discharge (given distances only)
  If optCalcType(2).Value Then
    c1wid = 42  'column 1
    c2wid = 37  'column 2
    c1fmt = "!" & String$(c1wid, "@") 'left-justified
    c2fmt = " " & String$(c2wid, "@") '1 space, right-justified

    hdr = "" 'start with a blank string
  
    AppendStr hdr, Format$("Number of Distances:", c1fmt), False
    AppendStr hdr, Format$(Format$(Nauto), c2fmt), False
  
    Print #1, hdr
  
    For i = 0 To Nauto - 1 'number of distances
      c1wid = 42  'column 1
      c2wid = 37  'column 2
      c1fmt = "!" & String$(c1wid, "@") 'left-justified
      c2fmt = " " & String$(c2wid, "@") '1 space, right-justified
  
      hdr = "" 'start with a blank string
  
      AppendStr hdr, "", True
    
      AppendStr hdr, Format$("Distance (" & UnitsName(UN_LENGTH) & "):", c1fmt), False
      AppendStr hdr, Format$(AGFormat(UnitsDisplay(Xauto(i), UN_LENGTH)), c2fmt), False
      AppendStr hdr, "", True
      AppendStr hdr, Format$("Discharge Rate (m3/s):", c1fmt), False
      AppendStr hdr, Format$(AGFormat(Rauto(i)), c2fmt), False
      AppendStr hdr, "", True
      AppendStr hdr, Format$("Number of Times:", c1fmt), False
      AppendStr hdr, Format$(Format$(TPD.np(i)), c2fmt), False 'tbc
      Print #1, hdr
  
      c1wid = 5   'column 1
      c2wid = 15  'column 2
      c1fmt = String$(c1wid, "@") 'right-justified
      c2fmt = " " & String$(c2wid, "@") '1 space, right-justified
      Print #1, "    #      Time (sec)    Concentration (ng/L)(ppt)"
      For j = 0 To TPD.np(i) - 1
        Print #1, Format$(j + 1, c1fmt) + _
                  Format$(TPD.X(j, i), c2fmt) + _
                  Format$(TPD.Y(j, i), c2fmt)
      Next
    Next
  End If
  
  Close #1
  Exit Sub
  
ErrHandExportExams:
  Msg = "Error writing file: " + fn + vbCrLf + Error$(Err)
  MsgBox Msg, vbCritical + vbOKOnly
  Close #1
  Exit Sub
End Sub


Private Sub cmdExport_Click()
'Export plot data
  If optCalcType(0).Value Then
    MsgBox "No Export for Single-Point Calculations", _
      vbCritical + vbOKOnly
  ElseIf optCalcType(1).Value Then
    If NeedCalcs Then Calculate
    If Not NeedCalcs Then
      If GenPlotTitles(PV_SATM, False) And GenPlotUnits(PV_SATM) Then
        frmExportToolbox.Show vbModal
      End If
    End If
  ElseIf optCalcType(2).Value Then
    If NeedCalcs Then Calculate
    If Not NeedCalcs Then
      If GenPlotTitles(PV_SADI, False) And GenPlotUnits(PV_SADI) Then
        frmExportToolbox.Show vbModal
      End If
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
  If optCalcType(0).Value Then
    MsgBox "No Plots for Single-Point Calculations", _
      vbCritical + vbOKOnly
    Exit Sub
  End If
  If NeedCalcs Then Calculate
  If NeedCalcs Then Exit Sub
  
  'Save sources and titles
  For i = 0 To 4: saveDataSource(i) = PlotGetDataSource(i): Next 'save existing
  For i = 0 To 4: saveDataTitle(i) = PlotGetDataTitle(i): Next 'save existing
  'Clear cources and titles
  For i = 0 To 4: PlotSetDataSource i, "": Next 'clear
  For i = 0 To 4: PlotSetDataTitle i, "": Next 'clear
    
  If optCalcType(1).Value Then
    'Set sources and titles
    j = 0
    For i = 0 To 3
      If txtTimeTime(i) <> "" Then
        PlotSetDataSource j, "ToolboxData: " + Format$(j)
        PlotSetDataTitle j, txtTimeTime(i) + " sec"
        j = j + 1
      End If
    Next
    'plot
    If SetupPlot(PV_SATM) Then frmPlot.Show vbModal
  ElseIf optCalcType(2).Value Then
    'Set sources and titles
    j = 0
    For i = 0 To 3
      If txtDistDist(i) <> "" Then
        PlotSetDataSource j, "ToolboxData: " + Format$(j)
        PlotSetDataTitle j, txtDistDist(i) + " " + UnitsName(UN_LENGTH)
        j = j + 1
      End If
    Next
    'plot
    If SetupPlot(PV_SADI) Then frmPlot.Show vbModal
  End If
   
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
    lblSprayLineUnits = UnitsName(UN_LENGTH)
    lblStreamWidthUnits = UnitsName(UN_LENGTH)
    lblStreamDepthUnits = UnitsName(UN_LENGTH)
    lblStreamRateUnits = UnitsName(UN_BIGFLOWRATE)
    lblStreamSpeedUnits = UnitsName(UN_SPEED)
    lblStreamDistUnits = UnitsName(UN_LENGTH)
    lblRechargeRateUnits = UnitsName(UN_RECHARGERATE)
    lblSingleDistUnits = UnitsName(UN_LENGTH)
    lblTimeDistUnits(0) = UnitsName(UN_LENGTH)
    lblTimeDistUnits(1) = UnitsName(UN_LENGTH)
    lblDistDistUnits = UnitsName(UN_LENGTH)
    lblActiveRateUnits = UnitsName(UN_RATEMASS)
    
    'Convert any existing user-defined values
    '(if this is not the first time this form has been shown)
    If PreviousUnits <> -1 Then
      If txtSprayLine.Text <> "" Then
        txtSprayLine.Text = _
          AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtSprayLine.Text), _
          UN_LENGTH, PreviousUnits), UN_LENGTH))
      End If
      If txtStreamWidth.Text <> "" Then
        txtStreamWidth.Text = _
          AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtStreamWidth.Text), _
          UN_LENGTH, PreviousUnits), UN_LENGTH))
      End If
      If txtStreamDepth.Text <> "" Then
        txtStreamDepth.Text = _
          AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtStreamDepth.Text), _
          UN_LENGTH, PreviousUnits), UN_LENGTH))
      End If
      If txtStreamRate.Text <> "" Then
        txtStreamRate.Text = _
          AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtStreamRate.Text), _
          UN_BIGFLOWRATE, PreviousUnits), UN_BIGFLOWRATE))
      End If
      If txtStreamDist.Text <> "" Then
        txtStreamDist.Text = _
          AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtStreamDist.Text), _
          UN_LENGTH, PreviousUnits), UN_LENGTH))
      End If
      If txtRechargeRate.Text <> "" Then
        txtRechargeRate.Text = _
          AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtRechargeRate.Text), _
          UN_RECHARGERATE, PreviousUnits), UN_RECHARGERATE))
      End If
      If txtSingle(1).Text <> "" Then
        txtSingle(1).Text = _
          AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtSingle(1).Text), _
          UN_LENGTH, PreviousUnits), UN_LENGTH))
      End If
      For i = 0 To 1
        If txtTimeDist(i).Text <> "" Then
          txtTimeDist(i).Text = _
            AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtTimeDist(i).Text), _
            UN_LENGTH, PreviousUnits), UN_LENGTH))
        End If
      Next
      For i = 0 To 3
        If txtDistDist(i).Text <> "" Then
          txtDistDist(i).Text = _
            AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtDistDist(i).Text), _
            UN_LENGTH, PreviousUnits), UN_LENGTH))
        End If
      Next
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
  optCalcType(0).Value = True   'Select default calc type
  txtSprayLine.Text = AGFormat$(UnitsDisplay(100, UN_LENGTH))
  txtStreamWidth.Text = AGFormat$(UnitsDisplay(3, UN_LENGTH))
  txtStreamDepth.Text = AGFormat$(UnitsDisplay(0.5, UN_LENGTH))
  txtTurnTime.Text = AGFormat$(30)
  txtStreamRate.Text = AGFormat$(UnitsDisplay(1.5, UN_BIGFLOWRATE))
  txtStreamDist.Text = AGFormat$(UnitsDisplay(50, UN_LENGTH))
  txtRiparian.Text = AGFormat$(0)
  txtDecayFactor.Text = AGFormat$(0)
  txtRechargeRate.Text = AGFormat$(UnitsDisplay(0, UN_RECHARGERATE))
  txtSingle(0).Text = AGFormat(0)
  txtActiveRate.Text = AGFormat$(UnitsDisplay( _
    UD.SM.ACfrac * UD.SM.FlowRate * UD.SM.NonVGrav, UN_RATEMASS))
  PreviousUnits = -1
End Sub

Private Sub optCalcType_Click(Index As Integer)
  Dim c As Control
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    For Each c In picCalcType
      If c.Index = Index Then
        c.Left = 120
        c.Top = 480
        c.BorderStyle = 0
        c.Visible = True
      Else
        c.Visible = False
      End If
    Next
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

Private Sub Calculate()
'Calculations for Stream Toolbox
  
  'Arguments for agstrm
  Dim ISTYPE As Long       'Calculation type 0=Single point 1=Time 2=Dist
  Dim INTYPE As Long       'Input Type:  istype=0: 0=time 1=dist 2=conc
                           '             istype=1,2: 0=auto >0=# inputs
  Dim XWIDE As Single      'Stream Width (m)
  Dim XDEEP As Single      'Stream Depth (m)
  Dim XACT As Single       'Active Rate (kg/ha) for Tier 1
  Dim XDIST As Single      'Distance to Stream centerline (m)
  Dim XSRATE As Single     'Stream Flow Rate (m3/s)
  Dim XSLENG As Single     'Spray Line Length (m)
  Dim XSTURN As Single     'Turn-Around Time (sec)
  Dim XRIPAR As Single     'Riparian removal fraction
  Dim XDECAY As Single     'Stream Decay Rate (%/km)
  Dim XCHARG As Single     'Recharge Rate m3/s/km
  Dim XINPTS(5) As Single  'Input values: istype=0: xinpts(0)=input value
                           '              istype=1,2: xinpts(0-1)=range
                           '                          xinpts(2-5)=points
  Dim iunit As Long        'Units flag
  Dim lfl As Long          'Operation flag: set to 0 to init
  Dim xsngl(2) As Single   '(0)=time (1)=dist (2)=conc, for istype=0
  Dim ier As Long          'error flag: 0=done 1=warn 2=err 4=err,string only
  Dim realwd(2) As Single  'error data array
  Dim cdat As String       'error message
  Dim clen As Long         'length of error message
  
  'other local vars
  Dim AutoRange As Integer
  Dim CalcType As Integer
  Dim Msg As String
  Dim NPlong(3) As Long
  Dim tmpval(3) As Single
  Dim ntmp As Integer
  Dim OkToDoCalcs As Integer
  Dim Success As Integer

  ' Change the form mouse pointer
  Me.MousePointer = vbHourglass 'hourglass
  
  OkToDoCalcs = True 'set to false to prevent call to agstrm
  
  'Extract the input data from the form controls
  
  'geometry
  XSLENG = UnitsInternal(Val(txtSprayLine.Text), UN_LENGTH)
  XWIDE = UnitsInternal(Val(txtStreamWidth.Text), UN_LENGTH)
  XDEEP = UnitsInternal(Val(txtStreamDepth.Text), UN_LENGTH)
  XSTURN = Val(txtTurnTime.Text)
  XSRATE = UnitsInternal(Val(txtStreamRate.Text), UN_BIGFLOWRATE)
  XDIST = UnitsInternal(Val(txtStreamDist.Text), UN_LENGTH)
  XRIPAR = Val(txtRiparian.Text)
  XDECAY = Val(txtDecayFactor.Text)
  XCHARG = UnitsInternal(Val(txtRechargeRate.Text), UN_RECHARGERATE)
  
  'Tier 1
  XACT = UnitsInternal(Val(txtActiveRate.Text) / (UD.SM.FlowRate * UD.SM.NonVGrav), UN_RATEMASS)
  
  'control
  For i = 0 To 2
    If optCalcType(i).Value Then CalcType = i
  Next
  Select Case CalcType
  Case 0: 'Single Point
    If CalcOutputMarker >= 0 Then
      ISTYPE = 0 'Single Point
      INTYPE = CalcOutputMarker
      TPD.NC = 0 'no plot data will be generated
      Select Case CalcOutputMarker
      Case 0, 2: 'time, conc
        XINPTS(0) = Val(txtSingle(CalcOutputMarker).Text)
      Case 1:    'distance
        XINPTS(0) = UnitsInternal(Val(txtSingle(CalcOutputMarker).Text), UN_LENGTH)
      End Select
    Else
      Msg = "No control inputs defined!"
      MsgBox Msg, vbCritical + vbOKOnly
      OkToDoCalcs = False
    End If
  
  Case 1: 'Times
    AutoRange = (cbxAuto(CalcType).Value = 1)
    ISTYPE = 1 'Times
    INTYPE = 0 'for autorange, >0 for manual
    TPD.NC = 4 '4 data sets will be generated for autorange
    'get start, end
    ntmp = 0
    For i = 0 To 1
      AddToArray UnitsInternal(Val(txtTimeDist(i).Text), _
        UN_LENGTH), ntmp, XINPTS()
    Next
    'get entered values
    If Not AutoRange Then
      ntmp = 0
      For i = 0 To 3
        If Len(Trim$(txtTimeTime(i).Text)) > 0 Then
          AddToArray Val(txtTimeTime(i).Text), ntmp, tmpval()
        End If
      Next
      If ntmp > 0 Then
        INTYPE = ntmp
        For i = 0 To INTYPE - 1
          XINPTS(i + 2) = tmpval(i)
        Next
        TPD.NC = INTYPE
      'if the user entered no values, make autorange
      Else
        cbxAuto(CalcType).Value = 1
        AutoRange = True
      End If
    End If
  
  Case 2: 'Distances
    AutoRange = (cbxAuto(CalcType).Value = 1)
    ISTYPE = 2 'Distances
    INTYPE = 0 'for autorange, >0 for manual
    TPD.NC = 4 '4 data sets will be generated for autorange
    'get start, end
    ntmp = 0
    For i = 0 To 1
      AddToArray Val(txtDistTime(i).Text), ntmp, XINPTS()
    Next
    'get entered values
    If Not AutoRange Then
      ntmp = 0
      For i = 0 To 3
        If Len(Trim$(txtDistDist(i).Text)) > 0 Then
          AddToArray UnitsInternal(Val(txtDistDist(i).Text), _
            UN_LENGTH), ntmp, tmpval()
        End If
      Next
      If ntmp > 0 Then
        INTYPE = ntmp
        For i = 0 To INTYPE - 1
          XINPTS(i + 2) = tmpval(i)
        Next
        TPD.NC = INTYPE
      'if the user entered no values, make autorange
      Else
        cbxAuto(CalcType).Value = 1
        AutoRange = True
      End If
    End If
    
  End Select
  
  'Perform the calculations
  Success = False 'default value for flag
  If OkToDoCalcs Then
  
    'Set up the global area to store the calculated results
    'TPD.nc has been set above
    TPD.X1D = False 'separate X col for each Y col
    If TPD.NC = 0 Then
      ReDim TPD.np(0)
      ReDim TPD.X(0, 0)
      ReDim TPD.Y(0, 0)
    Else
      ReDim TPD.np(TPD.NC - 1)
      ReDim TPD.X(1199, TPD.NC - 1)
      ReDim TPD.Y(1199, TPD.NC - 1)
    End If
    
    'Do the calculations
    iunit = UP.Units          'get units from preferences
    lfl = 0                   'reset flag
    cdat = Space$(40)         'allocate string space
    Do While OkToDoCalcs
      Call agstrm(UD, _
                  CLng(UC.NumSgl), UC.SglDist(0), UC.SglVal(0), UC.HalfVal(0), _
                  ISTYPE, INTYPE, _
                  XWIDE, XDEEP, XACT, XDIST, XSRATE, XSLENG, _
                  XSTURN, XRIPAR, XDECAY, XCHARG, XINPTS(0), iunit, _
                  lfl, xsngl(0), Nauto, Xauto(0), Rauto(0), _
                  NPlong(0), TPD.X(0, 0), TPD.Y(0, 0), _
                  NEX, EXTIME(0), EXDIST(0), EXCONC(0), _
                  ier, realwd(0), cdat, clen)
      Select Case CheckAgstrmStatus(ier, realwd(), cdat, clen)
      Case 0 'success
        Success = True
        OkToDoCalcs = False
      Case 1 'retry
        'loop around and try again
      Case 2 'abort
        Success = False
        OkToDoCalcs = False
      End Select
    Loop
  End If

  If Success Then
    'reset calc flag
    NeedCalcs = False
    
    'finish up plot data
    'We can convert the data here, because we know the units will
    'not change within the life of the data
    If CalcType > 0 Then 'no plot data for single-point calcs
      For ic = 0 To TPD.NC - 1 'loop through all curves
        TPD.np(ic) = NPlong(ic)
        'titles and units conversion for distance
        If CalcType = 1 Then 'Time calcs produce a range of distances
          For i = 0 To TPD.np(ic) - 1
            TPD.X(i, ic) = UnitsDisplay(TPD.X(i, ic), UN_LENGTH)
          Next
        Else 'Distance calcs produce a range of times
        End If
      Next
    End If
    
    'Place calculated values back in the controls
    PropTakeAction = False
    Select Case ISTYPE
    Case 0 'Single
      Select Case CalcOutputMarker
      Case 0: 'time
        txtSingle(1).Text = AGFormat$(UnitsDisplay(xsngl(1), UN_LENGTH))
        lblSingleConc.Caption = AGFormat$(xsngl(2))
      Case 1:    'distance
        txtSingle(0).Text = AGFormat$(xsngl(0))
        lblSingleConc.Caption = AGFormat$(xsngl(2))
      End Select
      
    Case 1 'Time
      If AutoRange Then
        For i = 0 To Nauto - 1
          txtTimeTime(i).Text = AGFormat$(Xauto(i))
        Next
        For i = Nauto To 3
          txtTimeTime(i).Text = ""
        Next
      End If
    
    Case 2 'Distance
      If AutoRange Then
        For i = 0 To Nauto - 1
          txtDistDist(i).Text = AGFormat$(UnitsDisplay(Xauto(i), UN_LENGTH))
        Next
        For i = Nauto To 3
          txtDistDist(i).Text = ""
        Next
      End If
    End Select
    PropTakeAction = True
  End If
  
  Me.MousePointer = vbDefault 'default
End Sub

Private Sub ClearOutputFields()
'clear single-point output fields
'Don't clear the one pointed to by CalcOutputMarker
  Dim PTAsave As Integer
  
  PTAsave = PropTakeAction
  PropTakeAction = False 'desensitize controls
  
  If optCalcType(0).Value = True Then
    For i = 0 To 1
      If i = CalcOutputMarker Then
        txtSingle(i).ForeColor = vbRed
      Else
        txtSingle(i).Text = ""
        txtSingle(i).ForeColor = vbBlack
      End If
    Next
    lblSingleConc.Caption = ""
  End If
  
  PropTakeAction = PTAsave 'restore control sensitivity
End Sub

Private Sub txtDecayFactor_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
  End If
End Sub

Private Sub txtDistDist_Change(Index As Integer)
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    optCalcType(2).Value = True
  End If
End Sub

Private Sub txtDistTime_Change(Index As Integer)
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    optCalcType(2).Value = True
  End If
End Sub

Private Sub txtRechargeRate_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
  End If
End Sub

Private Sub txtRiparian_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
  End If
End Sub

Private Sub txtSingle_Change(Index As Integer)
'When this control changes, clear the other
'members of the array
  If PropTakeAction Then
    PropTakeAction = False
    optCalcType(0).Value = True
    If Trim$(txtSingle(Index).Text) = "" Then
      CalcOutputMarker = -1
    Else
      CalcOutputMarker = Index
    End If
    NeedCalcs = True
    ClearOutputFields
    PropTakeAction = True
  End If
End Sub


Private Sub txtSprayLine_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
  End If
End Sub


Private Sub txtStreamDepth_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
    CalcStreamSpeed
  End If
End Sub

Private Sub txtStreamDist_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
  End If
End Sub

Private Sub txtStreamRate_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
    CalcStreamSpeed
  End If
End Sub


Private Sub txtStreamWidth_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
    CalcStreamSpeed
  End If
End Sub


Private Sub txtTimeDist_Change(Index As Integer)
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    optCalcType(1).Value = True
  End If
End Sub


Private Sub txtTimeTime_Change(Index As Integer)
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    optCalcType(1).Value = True
  End If
End Sub



Private Function CheckAgstrmStatus(ier As Long, realwd() As Single, _
  cdat As String, clen As Long) As Integer
'Check and interpret the status value returned by agstrm.
'Some status values require querying the user
'
'Input:
' ier        - status returned by agstrm
' realwd     - array of three values returned by agstrm
' cdat, clen - message string and length returned by agstrm
'
'Returns: 0: success, calcs complete
'         1: warning, try again
'         2: error, abort
'
  Dim Msg As String
  Dim outstr As String
  Dim minstr As String
  Dim maxstr As String
  
  Const IER_OK = 0
  Const IER_REDO = 1
  Const IER_ABORT = 2
  
  
  Select Case ier
    Case 0  'success, calcs are done
      CheckAgstrmStatus = IER_OK

    Case 1 'warning with msg and data
      outstr = AGFormat$(realwd(0))
      minstr = AGFormat$(realwd(1))
      maxstr = AGFormat$(realwd(2))
  
      Msg = "Warning!" + Chr$(13)
      Msg = Msg + Chr$(34) + Trim$(cdat) + Chr$(34) + Chr$(13)
      Msg = Msg + "is out of the suggested range. The limits are:" + Chr$(13)
      Msg = Msg + Chr$(13)
      Msg = Msg + "Min: " + minstr + Chr$(13)
      Msg = Msg + "Val: " + outstr + Chr$(13)
      Msg = Msg + "Max: " + maxstr + Chr$(13)
      Msg = Msg + Chr$(13)
      Msg = Msg + "Continue with calculations?"
      If MsgBox(Msg, vbExclamation + vbYesNo) = vbNo Then
        CheckAgstrmStatus = IER_ABORT
      Else
        CheckAgstrmStatus = IER_REDO
      End If
    Case 2 'error with msg and data
      outstr = AGFormat$(realwd(0))
      minstr = AGFormat$(realwd(1))
      maxstr = AGFormat$(realwd(2))
  
      Msg = "Error!" + Chr$(13)
      Msg = Msg + Chr$(34) + Trim$(cdat) + Chr$(34) + Chr$(13)
      Msg = Msg + "is out of range. The limits are:" + Chr$(13)
      Msg = Msg + Chr$(13)
      Msg = Msg + "Min: " + minstr + Chr$(13)
      Msg = Msg + "Val: " + outstr + Chr$(13)
      Msg = Msg + "Max: " + maxstr
      MsgBox Msg, vbCritical + vbOKOnly
      CheckAgstrmStatus = IER_ABORT

    Case 3 'warning with msg
      Msg = "Warning! "
      Msg = Msg & Left$(cdat, clen)
      Msg = Msg + Chr$(13)
      Msg = Msg + "Continue with calculations?"
      If MsgBox(Msg, vbExclamation + vbYesNo) = vbNo Then
        CheckAgstrmStatus = IER_ABORT
      Else
        CheckAgstrmStatus = IER_REDO
      End If
    
    Case 4 'error with msg
      Msg = "Error! "
      Msg = Msg & Left$(cdat, clen)
      MsgBox Msg, vbCritical + vbOKOnly
      CheckAgstrmStatus = IER_ABORT
    
    Case Else 'just in case...
      CheckAgstrmStatus = IER_ABORT
  End Select
End Function

Private Sub txtTurnTime_Change()
  If PropTakeAction Then
    NeedCalcs = True 'Value changed, need to recalc
    ClearOutputFields
  End If
End Sub



Private Sub CalcStreamSpeed()
'Calculate the Stream Flow Speed from the Flow Rate,
'the Width, and the Depth, assuming a rectangular cross section
  Dim wid As Single
  Dim dep As Single
  Dim flo As Single
  Dim spd As Single
  
  lblStreamSpeed.Caption = ""
  wid = UnitsInternal(Val(txtStreamWidth.Text), UN_LENGTH)
  If wid = 0 Then Exit Sub
  dep = UnitsInternal(Val(txtStreamDepth.Text), UN_LENGTH)
  If dep = 0 Then Exit Sub
  flo = UnitsInternal(Val(txtStreamRate.Text), UN_BIGFLOWRATE)
  spd = flo / wid / dep
  lblStreamSpeed.Caption = AGFormat$(UnitsDisplay(spd, UN_SPEED))
End Sub

