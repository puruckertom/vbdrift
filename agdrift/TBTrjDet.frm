VERSION 5.00
Begin VB.Form frmTBTrajDetails 
   Caption         =   "Trajectory Details"
   ClientHeight    =   6675
   ClientLeft      =   1140
   ClientTop       =   1500
   ClientWidth     =   9480
   HelpContextID   =   1493
   Icon            =   "TBTRJDET.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6675
   ScaleWidth      =   9480
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Cop&y"
      Height          =   375
      HelpContextID   =   1493
      Left            =   6840
      TabIndex        =   3
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "&Abort"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1493
      Left            =   7680
      TabIndex        =   2
      Top             =   5640
      Width           =   735
   End
   Begin VB.Frame fraDrawArea 
      Caption         =   "Trajectories"
      Height          =   5775
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9375
      Begin VB.PictureBox picDrawArea 
         AutoRedraw      =   -1  'True
         Height          =   5415
         HelpContextID   =   1493
         Left            =   120
         ScaleHeight     =   5355
         ScaleWidth      =   9075
         TabIndex        =   8
         Top             =   240
         Width           =   9135
      End
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "&Plot"
      Height          =   375
      HelpContextID   =   1493
      Left            =   7680
      TabIndex        =   1
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Clos&e"
      Height          =   375
      HelpContextID   =   1493
      Left            =   8520
      TabIndex        =   0
      Top             =   6000
      Width           =   855
   End
   Begin VB.Frame fraControl 
      Caption         =   "Control"
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   5760
      Width           =   6495
      Begin VB.OptionButton optTransform 
         Caption         =   "Aircraft Coordinates"
         Height          =   255
         HelpContextID   =   1493
         Index           =   1
         Left            =   4680
         TabIndex        =   7
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optTransform 
         Caption         =   "Terrain Coordinates"
         Height          =   255
         HelpContextID   =   1493
         Index           =   0
         Left            =   4680
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cboView 
         Height          =   315
         HelpContextID   =   1493
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtDiam 
         Height          =   285
         HelpContextID   =   1363
         Left            =   1200
         TabIndex        =   4
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "View"
         Height          =   195
         Left            =   2835
         TabIndex        =   13
         Top             =   405
         Width           =   345
      End
      Begin VB.Label lblInput0 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Drop Size:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   405
         Width           =   900
      End
      Begin VB.Label lblUnits0 
         AutoSize        =   -1  'True
         Caption         =   "µm"
         Height          =   195
         Left            =   2130
         TabIndex        =   11
         Top             =   405
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmTBTrajDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OkToContinue As Boolean

Private Sub cboView_Click()
  picDrawArea.Cls
End Sub

Private Sub cmdAbort_Click()
  OkToContinue = False
End Sub

Private Sub cmdCopy_Click()
  Clipboard.Clear
  Clipboard.SetData picDrawArea.Image
End Sub

Private Sub cmdPlot_Click()
'Calculate and plot the trajectories
'
'X points to the right wing
'Y points up
'Z points to the tail
'
'This routine does all the drawing with GRAPHQ calls only.
'
  Dim Diam As Single
  Dim ViewDir As Integer
  Dim NTR As Long
  Dim ScaleLims(5) As Single 'Xmin, Xmax, Ymin, Ymax, Zmin, Zmax
  Dim Xmin As Single, Xmax As Single, Xinc As Single
  Dim Ymin As Single, Ymax As Single, Yinc As Single
  Dim Zmin As Single, Zmax As Single, Zinc As Single
  ReDim OldPos(2, MAX_NOZZLES - 1) As Single
  ReDim newpos(2, MAX_NOZZLES - 1) As Single
  ReDim DropStat(MAX_NOZZLES - 1) As Long
  Dim NumAct As Long
  Dim i As Integer
  
  'Constants for view selection
  Const ViewRear = 0   'Y Horiz, Z Vert
  Const ViewTop = 1    'Y Horiz, X Vert
  Const ViewRight = 2  'X Horiz, Z Vert
  
  Const PLOTMARGIN = 500 'distance between pic border and plot box
  
  picDrawArea.Cls
  initq picDrawArea
  Me.MousePointer = vbHourglass 'Hourglass
  
  'Switch Calc/Abort Buttons
  OkToContinue = True
  cmdPlot.Visible = False
  cmdAbort.Visible = True
  
  'Recover Drop Diameter and View direction
  Diam = Val(txtDiam.Text)
  ViewDir = cboView.ListIndex
  For i = 0 To 1
    If optTransform(i).Value Then
      NTR = i
      Exit For
    End If
  Next
  
  'Calc setup
  For i = 0 To UD.NZ.NumNoz - 1
    DropStat(i) = 1 'Initialize status flags
  Next
  
  'Init the calcs
  'Get the scale limits, and the initial drop positions
  'DropPos(1, 2, n)
  '              |- number of nozzles
  '           |---- X, Y, Z
  '        |------- 0 = new positions 1 = previous positions
  Call agtraj(UD, Diam, NTR, ScaleLims(0), newpos(0, 0), DropStat(0))
  'Extract minima and maxima in display units
  Xmin = UnitsDisplay(ScaleLims(0), UN_LENGTH): Xmax = UnitsDisplay(ScaleLims(1), UN_LENGTH)
  Ymin = UnitsDisplay(ScaleLims(2), UN_LENGTH): Ymax = UnitsDisplay(ScaleLims(3), UN_LENGTH)
  Zmin = UnitsDisplay(ScaleLims(4), UN_LENGTH): Zmax = UnitsDisplay(ScaleLims(5), UN_LENGTH)
'Debug.Print "Scale Limits"
'Debug.Print Xmin; Xmax
'Debug.Print Ymin; Ymax
'Debug.Print Zmin; Zmax

  'Clean up the scale limits
  autoscq GQ_LINEAR, Xmin, Xmax, Xinc
  autoscq GQ_LINEAR, Ymin, Ymax, Yinc
  autoscq GQ_LINEAR, Zmin, Zmax, Zinc
  
  'Set the plot scales
  With picDrawArea
    Select Case ViewDir
    Case ViewRear 'Y horiz, Z vert
      viewq PLOTMARGIN, .Width - PLOTMARGIN, .Height - PLOTMARGIN, PLOTMARGIN, _
           Ymin, Ymax, Zmin, Zmax, GQ_BOX
      axisq GQ_XAXIS, GQ_LINEAR, GQ_NOGRID, Yinc, 1
      axisq GQ_YAXIS, GQ_LINEAR, GQ_NOGRID, Zinc, 1
    Case ViewTop  'Y horiz, X vert
      viewq PLOTMARGIN, .Width - PLOTMARGIN, .Height - PLOTMARGIN, PLOTMARGIN, _
           Ymin, Ymax, Xmin, Xmax, GQ_BOX
      axisq GQ_XAXIS, GQ_LINEAR, GQ_NOGRID, Yinc, 1
      axisq GQ_YAXIS, GQ_LINEAR, GQ_NOGRID, Xinc, 1
    Case ViewRight 'X Horiz, Z vert
      viewq PLOTMARGIN, .Width - PLOTMARGIN, .Height - PLOTMARGIN, PLOTMARGIN, _
           Xmin, Xmax, Zmin, Zmax, GQ_BOX
      axisq GQ_XAXIS, GQ_LINEAR, GQ_NOGRID, Xinc, 1
      axisq GQ_YAXIS, GQ_LINEAR, GQ_NOGRID, Zinc, 1
    End Select
  End With
  
  'Iterate
  NumAct = UD.NZ.NumNoz 'All the nozzles are active at first
  While OkToContinue And NumAct > 0
    DoEvents
    'Move the 'current' positions to 'previous' storage
    CopyMemory OldPos(0, 0), newpos(0, 0), 3 * UD.NZ.NumNoz * Len(OldPos(0, 0))
    'Get new 'current' positions
    Call agtrgo(NumAct, newpos(0, 0), DropStat(0))
    Select Case ViewDir
    Case ViewRear
      For i = 0 To UD.NZ.NumNoz - 1
        If DropStat(i) > 0 Then
          moveq GQ_WORLD, UnitsDisplay(OldPos(1, i), UN_LENGTH), UnitsDisplay(OldPos(2, i), UN_LENGTH)
          drawq GQ_WORLD, UnitsDisplay(newpos(1, i), UN_LENGTH), UnitsDisplay(newpos(2, i), UN_LENGTH)
        End If
      Next
    Case ViewTop
      For i = 0 To UD.NZ.NumNoz - 1
        If DropStat(i) > 0 Then
          moveq GQ_WORLD, UnitsDisplay(OldPos(1, i), UN_LENGTH), UnitsDisplay(OldPos(0, i), UN_LENGTH)
          drawq GQ_WORLD, UnitsDisplay(newpos(1, i), UN_LENGTH), UnitsDisplay(newpos(0, i), UN_LENGTH)
        End If
      Next
    Case ViewRight
      For i = 0 To UD.NZ.NumNoz - 1
        If DropStat(i) > 0 Then
          moveq GQ_WORLD, UnitsDisplay(OldPos(0, i), UN_LENGTH), UnitsDisplay(OldPos(2, i), UN_LENGTH)
          drawq GQ_WORLD, UnitsDisplay(newpos(0, i), UN_LENGTH), UnitsDisplay(newpos(2, i), UN_LENGTH)
        End If
      Next
    End Select
  Wend
  
  'Switch Calc/Abort buttons back
  cmdPlot.Visible = True
  cmdAbort.Visible = False
  Me.MousePointer = vbDefault  'default
End Sub

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  CenterForm Me

  With cboView
    .AddItem "Rear"
    .AddItem "Top"
    .AddItem "Right"
    .ListIndex = 0
  End With
  
  cmdAbort.Visible = False
End Sub

Private Sub Form_Resize()
  'Adjust control positions and sizes to fit the form
  Dim minwidth As Single
  Dim minheight As Single
  Const MRGN = 100
  
  If Me.WindowState = vbMinimized Then Exit Sub
  
  'prevent the form from getting too small
  minwidth = 9600 '640 pixels
  minheight = 7200 '480 pixels
  If Me.Width < minwidth Then
    Me.Width = minwidth
    Exit Sub
  End If
  If Me.Height < minheight Then
    Me.Height = minheight
    Exit Sub
  End If
  
  'Ok/Calc/Plot buttons
  cmdOK.Top = Me.ScaleHeight - cmdOK.Height - MRGN
  cmdOK.Left = Me.ScaleWidth - cmdOK.Width - MRGN
  cmdPlot.Top = cmdOK.Top
  cmdPlot.Left = cmdOK.Left - cmdPlot.Width - MRGN
  cmdAbort.Top = cmdPlot.Top
  cmdAbort.Left = cmdPlot.Left
  cmdCopy.Top = cmdAbort.Top
  cmdCopy.Left = cmdAbort.Left - cmdAbort.Width - MRGN
  
  fraControl.Top = Me.ScaleHeight - fraControl.Height - MRGN
  
  fraDrawArea.Height = fraControl.Top - fraDrawArea.Top
  fraDrawArea.Width = Me.ScaleWidth - fraDrawArea.Left - MRGN
  picDrawArea.Top = MRGN * 2
  picDrawArea.Left = MRGN
  picDrawArea.Height = fraDrawArea.Height - MRGN * 3
  picDrawArea.Width = fraDrawArea.Width - MRGN * 2
End Sub

Private Sub optTransform_Click(Index As Integer)
  picDrawArea.Cls
End Sub

Private Sub txtDiam_Change()
  picDrawArea.Cls
End Sub
