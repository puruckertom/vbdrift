VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlot 
   Caption         =   "Plot"
   ClientHeight    =   5760
   ClientLeft      =   1815
   ClientTop       =   1845
   ClientWidth     =   7365
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "PLOT.frx":0000
   LinkTopic       =   "Graph"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5760
   ScaleWidth      =   7365
   Tag             =   "plot"
   Begin VB.PictureBox picArea 
      AutoRedraw      =   -1  'True
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   7275
      TabIndex        =   12
      Top             =   600
      Width           =   7335
      Begin VB.PictureBox picLegend 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4680
         ScaleHeight     =   465
         ScaleWidth      =   1065
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.PictureBox picPlot 
         ClipControls    =   0   'False
         Height          =   3015
         Left            =   1080
         ScaleHeight     =   2955
         ScaleWidth      =   4875
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label lblPlotTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "PlotTitle"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3420
         TabIndex        =   6
         Top             =   480
         Width           =   690
      End
      Begin VB.Label lblRunTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "RunTitle"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3435
         TabIndex        =   5
         Top             =   120
         Width           =   690
      End
      Begin VB.Label lblXaxis 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "X axis"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3735
         TabIndex        =   8
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label lblYaxis 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Y axis"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   375
         TabIndex        =   7
         Top             =   1800
         Width           =   495
      End
   End
   Begin VB.PictureBox picControlBar 
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   7275
      TabIndex        =   11
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Cop&y"
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin VB.ComboBox cboPlotVar 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   150
         Width           =   3090
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "&Options"
         Height          =   375
         HelpContextID   =   1220
         Left            =   960
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   0
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblRemoteControl 
      AutoSize        =   -1  'True
      Caption         =   "This label is used to control this form remotely"
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   3915
   End
End
Attribute VB_Name = "frmPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: plot.frm,v 1.11 2001/04/26 16:21:58 tom Exp $
Option Explicit

'Drag offsets to help reposition controls after dragging
Dim DragOffsetX As Single
Dim DragOffsetY As Single
Dim PropTakeAction As Integer
'The following keeps track of the previous successful
'ListIndex of the Plot combo box
Dim PreviousPlotVarListIndex As Integer

Private Sub ComboPlotVarAddItem(PlotVar As Long)
'Test a PlotVar for availability and possibly add it to the combo
  Dim PlotTitle As String
  Dim XTitle As String
  Dim YTitle As String
  Dim HelpID As Long
  If PlotIsAvailable(PlotVar) Then
    GenPlotTitleStrings PlotVar, False, PlotTitle, XTitle, YTitle, HelpID
    cboPlotVar.AddItem PlotTitle
    cboPlotVar.ItemData(cboPlotVar.NewIndex) = PlotVar
  End If
End Sub

Private Sub InitComboPlotVar()
'This is the only sub in the form that is not generic plotting.
'It sets up the plot selection combo according to the same rules
'that the View menu uses, but only if UI.PlotVar is not based
'on toolbox or temporary data
  
  cboPlotVar.Clear
  
  If (UI.PlotVar And PVA_SOURCE_MASK) = PVA_TB Then
    cboPlotVar.AddItem "Toolbox"
    cboPlotVar.ItemData(cboPlotVar.NewIndex) = -1
    cboPlotVar.ListIndex = cboPlotVar.NewIndex
  Else
    ComboPlotVarAddItem PV_VFINC
    ComboPlotVarAddItem PV_VFINC0
    ComboPlotVarAddItem PV_VFINC1
    ComboPlotVarAddItem PV_VFINC2
    ComboPlotVarAddItem PV_VFCUM
    ComboPlotVarAddItem PV_VFCUM0
    ComboPlotVarAddItem PV_VFCUM1
    ComboPlotVarAddItem PV_VFCUM2
    ComboPlotVarAddItem PV_DWDSDINC
    ComboPlotVarAddItem PV_DWDSDCUM
    ComboPlotVarAddItem PV_FXDSDINC
    ComboPlotVarAddItem PV_FXDSDCUM
    ComboPlotVarAddItem PV_SBDSDINC
    ComboPlotVarAddItem PV_SBDSDCUM
    ComboPlotVarAddItem PV_CNDSDINC
    ComboPlotVarAddItem PV_CNDSDCUM
    ComboPlotVarAddItem PV_DEP
    ComboPlotVarAddItem PV_PID
    ComboPlotVarAddItem PV_SV
    ComboPlotVarAddItem PV_VERT
    ComboPlotVarAddItem PV_CONC
    ComboPlotVarAddItem PV_LAYOUT
    ComboPlotVarAddItem PV_COV
    ComboPlotVarAddItem PV_MEAN
    ComboPlotVarAddItem PV_FA
    ComboPlotVarAddItem PV_SBDEP
    ComboPlotVarAddItem PV_SBCOVER
    ComboPlotVarAddItem PV_CANDEP
    ComboPlotVarAddItem PV_TA
    ComboPlotVarAddItem PV_DA
    ComboPlotVarAddItem PV_HA
    ComboPlotVarAddItem PV_TAB
'tbc    ComboPlotVarAddItem PV_TAP
  End If
    
  Exit Sub
End Sub

Private Sub cboPlotVar_Click()
  Dim PlotVar As Long
 
  If PropTakeAction Then
    On Error Resume Next 'in case the previous plotvar index is bad
    PlotVar = CLng(cboPlotVar.ItemData(cboPlotVar.ListIndex))
    If SetupPlot(PlotVar) Then 'try to setup the plot
      PreviousPlotVarListIndex = cboPlotVar.ListIndex 'remember the index
      lblRemoteControl.Caption = "replot" 'trigger a plot redraw
    Else
      'setup was unsuccessful, restore previous plot
      PropTakeAction = False 'turn off control reactions
      cboPlotVar.ListIndex = PreviousPlotVarListIndex  'restore this control
      PlotVar = CLng(cboPlotVar.ItemData(cboPlotVar.ListIndex)) 'get prev PlotVar
      SetupPlot PlotVar 'regen the previous plot
      lblRemoteControl.Caption = "replot" 'trigger a plot redraw
      PropTakeAction = True
    End If
  End If
  Exit Sub
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdCopy_Click()
  Clipboard.Clear
  Clipboard.SetData picArea.Image
End Sub

Private Sub cmdOptions_Click()
  PlotOptionsDialog
End Sub

Private Sub cmdPrint_Click()
  If PrinterExists() Then PlotPrintDialog
End Sub

Private Sub DataToForm()
'set up form according to saved data
  Dim i As Integer
  On Error GoTo DataToFormErrHand

  Dim PTAsave As Integer
  PTAsave = PropTakeAction 'save flag state
  PropTakeAction = False   'disable form controls

  'set the combo box index to match the current plot var
  For i = 0 To cboPlotVar.ListCount - 1
    If cboPlotVar.ItemData(i) = UI.PlotVar Then
      cboPlotVar.ListIndex = i
      PreviousPlotVarListIndex = i
      Exit For
    End If
  Next
  'Caption
  Me.Caption = PS.Caption
  'set the help context ID
  Me.HelpContextID = PS.HelpID
  'Init titles
  PlotSetFont PS.RunTitle.Font, lblRunTitle
  PlotSetFont PS.PlotTitle.Font, lblPlotTitle
  PlotSetFont PS.XTitle.Font, lblXaxis
  PlotSetFont PS.YTitle.Font, lblYaxis
  lblRunTitle.Caption = PS.RunTitle.Text
  lblPlotTitle.Caption = PS.PlotTitle.Text
  lblXaxis.Caption = PS.XTitle.Text & PS.Xunits
  lblYaxis.Caption = PS.YTitle.Text & PS.Yunits
  'Init Legend
  PlotSetFont PS.Legend.Font, picLegend

  'set label positions
  SetLabelPositions
  
'
' The picPlot object defines the plot area
'
  PS.XDS = picPlot.Left
  PS.XDE = picPlot.Left + picPlot.Width
  PS.YDS = picPlot.Top + picPlot.Height
  PS.YDE = picPlot.Top
  
  'replot on the screen
  Replot
  
  PropTakeAction = PTAsave 'restore flag value
  Exit Sub

DataToFormErrHand:
  Select Case UnexpectedError("Plot,DataToForm")
  Case vbAbort  'Abort - Stop the whole program
    End
  Case vbRetry  'Retry - Resume at the same line
    Resume
  Case vbIgnore 'Ignore - Resume at the next line
    Resume Next
  End Select
End Sub

Private Function EditTitle(TD As TitleData, Title As String) As Integer
'
' Call upon frmPlotEditTitle to alow the user to adjust the
' appearance of a Graph Title.
'
'returns true is successful, false if cancelled
'
  Dim s As String
  
  On Error GoTo ErrHandlerET
  'Create an instance of the form
  Dim f As New frmPlotEditTitle
  f.Caption = Title   'Title the window
  'Load the text attributes into the form's data control
  f!lblData.Caption = TD.Text
  f!lblData.FontName = TD.Font.Name
  f!lblData.FontSize = TD.Font.Size
  f!lblData.FontBold = TD.Font.Bold
  f!lblData.FontItalic = TD.Font.Italic
  f!lblData.FontUnderline = TD.Font.Underline
  f!lblData.FontStrikethru = TD.Font.Strikethru
  f!lblData.ForeColor = TD.Font.Color
  'Show the form as modal to collect the changes
  f.Show vbModal
  'check status
  If f.Tag = "False" Then
    EditTitle = False
    Exit Function
  End If
  'recover the new settigs
  TD.Text = f!lblData.Caption
  TD.Font.Name = f!lblData.FontName
  TD.Font.Size = f!lblData.FontSize
  TD.Font.Bold = f!lblData.FontBold
  TD.Font.Italic = f!lblData.FontItalic
  TD.Font.Underline = f!lblData.FontUnderline
  TD.Font.Strikethru = f!lblData.FontStrikethru
  TD.Font.Color = f!lblData.ForeColor
  'Dump the form
  Unload f
  PS.Changed = True 'user has changed plot settings
  EditTitle = True
  Exit Function

ErrHandlerET:
  s = "Could not set label font to " + Chr$(34) + f!lblData.Caption + Chr$(34) + "." + Chr$(13)
  s = s + "Try selecting a TrueType font instead."
  MsgBox s, vbExclamation + vbOKOnly
  Resume Next

End Function

Private Sub Form_Load()
  On Error GoTo FormLoadErrHand
  
  'Center the form
  Me.Left = (Screen.Width / 2) - (Me.Width / 2)
  Me.Top = (Screen.Height / 2) - (Me.Height / 2)
  
  'initialize the PlotVar combo box
  InitComboPlotVar

  PropTakeAction = True 'enable form controls

  DataToForm
  Exit Sub

FormLoadErrHand:
  Select Case UnexpectedError("Plot,Form,Load")
  Case vbAbort  'Abort - Stop the whole program
    End
  Case vbRetry  'Retry - Resume at the same line
    Resume
  Case vbIgnore 'Ignore - Resume at the next line
    Resume Next
  End Select
End Sub

Private Sub Form_Resize()
'Adjust the size of the Plot form elements to fit the form
  On Error GoTo FormResizeErrHand

  Dim newwidth As Integer
  Dim newheight As Integer
  'define the right and bottom margins
  Const RIGHTMARGIN = 500
  Const BOTTOMMARGIN = 1000
  Const TOPMARGIN = 750
  Const LEFTMARGIN = 1500
  'define minimum picturebox dimensions
  Const MINFORMWIDTH = 5000
  Const MINFORMHEIGHT = 5000
  'check out the new form size
  If Me.Width < MINFORMWIDTH Then
    Me.Width = MINFORMWIDTH  'will trigger a resize
    Exit Sub
  End If
  If Me.Height < MINFORMHEIGHT Then
    Me.Height = MINFORMHEIGHT 'will trigger a resize
    Exit Sub
  End If

  'fit the control bar picturebox to the form
  picControlBar.Width = Me.ScaleWidth
  'fit the plot area picture box to the form,
  'below the control bar
  picLegend.Visible = False 'turn the legend off for now
  picArea.Cls
  picArea.Width = Me.ScaleWidth
  picArea.Height = Me.ScaleHeight - picControlBar.Height
  'adjust the Plot picturebox
  newwidth = picArea.Width - LEFTMARGIN - RIGHTMARGIN
  newheight = picArea.Height - TOPMARGIN - BOTTOMMARGIN
  picPlot.Top = TOPMARGIN
  picPlot.Left = LEFTMARGIN
  picPlot.Width = newwidth
  picPlot.Height = newheight
'
' The picPlot object defines the plot area
'
  PS.XDS = picPlot.Left
  PS.XDE = picPlot.Left + picPlot.Width
  PS.YDS = picPlot.Top + picPlot.Height
  PS.YDE = picPlot.Top
  
  'Reposition the labels
  SetLabelPositions
  'Plot the graph
  Replot
  Exit Sub

FormResizeErrHand:
  Select Case UnexpectedError("Plot,Form,Resize")
  Case vbAbort  'Abort - Stop the whole program
    End
  Case vbRetry  'Retry - Resume at the same line
    Resume
  Case vbIgnore 'Ignore - Resume at the next line
    Resume Next
  End Select
End Sub

Private Sub FormToData()
'save form data in PS area
  PS.RunTitle.Text = lblRunTitle.Caption
  PS.RunTitle.Font.Name = lblRunTitle.FontName
  PS.RunTitle.Font.Size = lblRunTitle.FontSize
  PS.RunTitle.Font.Bold = lblRunTitle.FontBold
  PS.RunTitle.Font.Italic = lblRunTitle.FontItalic
  PS.RunTitle.Font.Underline = lblRunTitle.FontUnderline
  PS.RunTitle.Font.Strikethru = lblRunTitle.FontStrikethru
  PS.RunTitle.Font.Color = lblRunTitle.ForeColor
  
  PS.PlotTitle.Text = lblPlotTitle.Caption
  PS.PlotTitle.Font.Name = lblPlotTitle.FontName
  PS.PlotTitle.Font.Size = lblPlotTitle.FontSize
  PS.PlotTitle.Font.Bold = lblPlotTitle.FontBold
  PS.PlotTitle.Font.Italic = lblPlotTitle.FontItalic
  PS.PlotTitle.Font.Underline = lblPlotTitle.FontUnderline
  PS.PlotTitle.Font.Strikethru = lblPlotTitle.FontStrikethru
  PS.PlotTitle.Font.Color = lblPlotTitle.ForeColor
  'save X title
  PS.XTitle.Text = lblXaxis.Caption
  PS.XTitle.Font.Name = lblXaxis.FontName
  PS.XTitle.Font.Size = lblXaxis.FontSize
  PS.XTitle.Font.Bold = lblXaxis.FontBold
  PS.XTitle.Font.Italic = lblXaxis.FontItalic
  PS.XTitle.Font.Underline = lblXaxis.FontUnderline
  PS.XTitle.Font.Strikethru = lblXaxis.FontStrikethru
  PS.XTitle.Font.Color = lblXaxis.ForeColor
  'save Y title
  PS.YTitle.Text = lblYaxis.Caption
  PS.YTitle.Font.Name = lblYaxis.FontName
  PS.YTitle.Font.Size = lblYaxis.FontSize
  PS.YTitle.Font.Bold = lblYaxis.FontBold
  PS.YTitle.Font.Italic = lblYaxis.FontItalic
  PS.YTitle.Font.Underline = lblYaxis.FontUnderline
  PS.YTitle.Font.Strikethru = lblYaxis.FontStrikethru
  PS.YTitle.Font.Color = lblYaxis.ForeColor
  'save label positions
  SaveLabelPositions
End Sub

Private Sub lblPlotTitle_DblClick()
  If EditTitle(PS.PlotTitle, "Plot Title") Then DataToForm
End Sub

Private Sub lblPlotTitle_DragDrop(Source As Control, X As Single, Y As Single)
  'Reposition dragged controls on the form
  Dim thisctl As Control
  Set thisctl = lblPlotTitle
  Source.Move thisctl.Left + X - DragOffsetX, thisctl.Top + Y - DragOffsetY
  DragOffsetX = 0
  DragOffsetY = 0
  'Save the new postions
  SaveLabelPositions
  'replot
  Replot
End Sub

Private Sub lblPlotTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Start a Drag Event. Remember the mouse positions for
  'dropping later.
  If (Button And 1) > 0 Then 'Left button
    PS.Changed = True 'user has changed plot settings
    DragOffsetX = X
    DragOffsetY = Y
    lblPlotTitle.Drag vbBeginDrag  'start the drag operation
  End If
End Sub

Private Sub lblRemoteControl_Change()
'By changing the caption of this label, this form
'may be controlled
  Select Case lblRemoteControl.Caption
    Case ""   'this is the reset case, do nothing
    Case "regen"       'regenerate plot data
      GenPlotData UI.PlotVar
      DataToForm
    Case "replot"      'trigger a replot
      DataToForm
    Case "printdialog" 'execute the print dialog
      PrintDialog
  End Select
  lblRemoteControl.Caption = "" 'reset to ensure a change next time
End Sub

Private Sub lblRunTitle_DblClick()
  If EditTitle(PS.RunTitle, "Run Title") Then DataToForm
End Sub

Private Sub lblRunTitle_DragDrop(Source As Control, X As Single, Y As Single)
  'Reposition dragged controls on the form
  Dim thisctl As Control
  Set thisctl = lblRunTitle
  Source.Move thisctl.Left + X - DragOffsetX, thisctl.Top + Y - DragOffsetY
  DragOffsetX = 0
  DragOffsetY = 0
  'Save the new postions
  SaveLabelPositions
  'replot
  Replot
End Sub

Private Sub lblRunTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Start a Drag Event. Remember the mouse positions for
  'dropping later.
  If (Button And 1) > 0 Then 'Left button
    PS.Changed = True 'user has changed plot settings
    DragOffsetX = X
    DragOffsetY = Y
    lblRunTitle.Drag vbBeginDrag  'start the drag operation
  End If
End Sub

Private Sub lblXaxis_DblClick()
  If EditTitle(PS.XTitle, "X axis title") Then DataToForm
End Sub

Private Sub lblXaxis_DragDrop(Source As Control, X As Single, Y As Single)
  'Reposition dragged controls on the form
  Dim thisctl As Control
  Set thisctl = lblXaxis
  Source.Move thisctl.Left + X - DragOffsetX, thisctl.Top + Y - DragOffsetY
  DragOffsetX = 0
  DragOffsetY = 0
  'Save the new postions
  SaveLabelPositions
  'replot
  Replot
End Sub

Private Sub lblXaxis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Start a Drag Event. Remember the mouse positions for
  'dropping later.
  If (Button And 1) > 0 Then 'Left button
    PS.Changed = True 'user has changed plot settings
    DragOffsetX = X
    DragOffsetY = Y
    lblXaxis.Drag vbBeginDrag  'start the drag operation
  End If
End Sub

Private Sub lblYaxis_DblClick()
  If EditTitle(PS.YTitle, "Y axis title") Then DataToForm
End Sub

Private Sub lblYaxis_DragDrop(Source As Control, X As Single, Y As Single)
  'Reposition dragged controls on the form
  Dim thisctl As Control
  Set thisctl = lblYaxis
  Source.Move thisctl.Left + X - DragOffsetX, thisctl.Top + Y - DragOffsetY
  DragOffsetX = 0
  DragOffsetY = 0
  'Save the new postions
  SaveLabelPositions
  'replot
  Replot
End Sub

Private Sub lblYaxis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Start a Drag Event. Remember the mouse positions for
  'dropping later.
  If (Button And 1) > 0 Then 'Left button
    PS.Changed = True 'user has changed plot settings
    DragOffsetX = X
    DragOffsetY = Y
    lblYaxis.Drag vbBeginDrag  'start the drag operation
  End If
End Sub

Private Sub picArea_DblClick()
'bring up the plot options dialog
  PlotOptionsDialog
End Sub

Private Sub picArea_DragDrop(Source As Control, X As Single, Y As Single)
  'Reposition dragged controls on the form
  Source.Move X - DragOffsetX, Y - DragOffsetY
  DragOffsetX = 0
  DragOffsetY = 0
  'Save the new postions
  SaveLabelPositions
  'replot
  Replot
End Sub

Private Sub picLegend_DragDrop(Source As Control, X As Single, Y As Single)
  'Reposition dragged controls on the form
  Dim thisctl As Control
  Set thisctl = picLegend
  Source.Move thisctl.Left + X - DragOffsetX, thisctl.Top + Y - DragOffsetY
  DragOffsetX = 0
  DragOffsetY = 0
  'Save the new postions
  SaveLabelPositions
  'replot
  Replot
End Sub

Private Sub picLegend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Start a Drag Event. Remember the mouse positions for
  'dropping later.
  If (Button And 1) > 0 Then 'Left button
    PS.Changed = True 'user has changed plot settings
    DragOffsetX = X
    DragOffsetY = Y
    picLegend.Drag vbBeginDrag  'start the drag operation
  End If
End Sub

Private Sub picPlot_DblClick()
'bring up the plot options dialog
  PlotOptionsDialog
End Sub

Private Sub PrintDialog()
  Dim BeginPage, EndPage, NumCopies, i
  
  'Set Cancel to True
  CMDialog1.CancelError = True
  On Error GoTo ErrHandlerPrint

  'Set dialog box flags
  CMDialog1.Flags = cdlPDHidePrintToFile

  'Display the Print Dialog box
  CMDialog1.ShowPrinter

  'Get user-selected values
  BeginPage = CMDialog1.FromPage
  EndPage = CMDialog1.ToPage
  NumCopies = CMDialog1.Copies

  For i = 1 To NumCopies
    'printing code goes here
    PlotGraph Printer
    Printer.EndDoc
  Next

  Exit Sub

ErrHandlerPrint:
  'User pressed Cancel button
  Exit Sub

End Sub

Private Sub SaveLabelPositions()
'Record label positions relative to the plot area
  'RunTitle - Center at constant relative horizontal position
  '           Top edge constant distance from plot area top
  PS.RunTitle.PosX = picArea.ScaleLeft + (lblRunTitle.Left + lblRunTitle.Width * 0.5) / picArea.ScaleWidth
  PS.RunTitle.PosY = picArea.ScaleTop + lblRunTitle.Top
  'PlotTitle - Center at constant relative horizontal position
  '            Top edge constant distance from plot area top
  PS.PlotTitle.PosX = picArea.ScaleLeft + (lblPlotTitle.Left + lblPlotTitle.Width * 0.5) / picArea.ScaleWidth
  PS.PlotTitle.PosY = picArea.ScaleTop + lblPlotTitle.Top
  'XTitle - Center at constant relative horizontal position
  '         Top edge constant distance from plot area bottom
  PS.XTitle.PosX = picArea.ScaleLeft + (lblXaxis.Left + lblXaxis.Width * 0.5) / picArea.ScaleWidth
  PS.XTitle.PosY = picArea.ScaleTop + picArea.ScaleHeight - lblXaxis.Top
  'YTitle - Left edge constant distance from plot area left edge
  '         Center at constant relative vertical position
  PS.YTitle.PosX = picArea.ScaleLeft + lblYaxis.Left
  PS.YTitle.PosY = picArea.ScaleTop + (lblYaxis.Top + lblYaxis.Height * 0.5) / picArea.ScaleHeight
  'Legend - Right edge constant distance from plot area right edge
  '         Top edge constant distance from plot area top
  PS.Legend.PosX = picArea.ScaleLeft + picArea.ScaleWidth - picLegend.Left - picLegend.Width
  PS.Legend.PosY = picArea.ScaleTop + picLegend.Top
End Sub

Private Sub SetLabelPositions()
'Reposition the plot labels
  'RunTitle - Center at constant relative horizontal position
  '           Top edge constant distance from plot area top
  lblRunTitle.Left = picArea.ScaleLeft + PS.RunTitle.PosX * picArea.ScaleWidth - (lblRunTitle.Width * 0.5)
  lblRunTitle.Top = picArea.ScaleTop + PS.RunTitle.PosY
  'PlotTitle - Center at constant relative horizontal position
  '            Top edge constant distance from plot area top
  lblPlotTitle.Left = picArea.ScaleLeft + PS.PlotTitle.PosX * picArea.ScaleWidth - (lblPlotTitle.Width * 0.5)
  lblPlotTitle.Top = picArea.ScaleTop + PS.PlotTitle.PosY
  'XTitle - Center at constant relative horizontal position
  '         Top edge constant distance from plot area bottom
  lblXaxis.Left = picArea.ScaleLeft + PS.XTitle.PosX * picArea.ScaleWidth - (lblXaxis.Width * 0.5)
  lblXaxis.Top = picArea.ScaleTop + picArea.ScaleHeight - PS.XTitle.PosY
  'YTitle - Left edge constant distance from plot area left edge
  '         Center at constant relative vertical position
  lblYaxis.Left = picArea.ScaleLeft + PS.YTitle.PosX
  lblYaxis.Top = picArea.ScaleTop + PS.YTitle.PosY * picArea.ScaleHeight - (lblYaxis.Height * 0.5)
End Sub

Private Sub SetLegendPosition()
'Reposition the plot labels
  'Legend - Right edge constant distance from plot area right edge
  '         Top edge constant distance from plot area top
  picLegend.Left = picArea.ScaleLeft + picArea.ScaleWidth - PS.Legend.PosX - picLegend.Width
  picLegend.Top = picArea.ScaleTop + PS.Legend.PosY
End Sub

Private Sub Replot()
  PlotGraph picArea
  PlotLabel picArea, lblRunTitle
  PlotLabel picArea, lblPlotTitle
  PlotLabel picArea, lblXaxis
  PlotLabel picArea, lblYaxis
  PlotLegend picArea
End Sub

Public Sub PlotLegend(dest As Control)
'Plot a legend
'
' dest is the plot destination
'
  On Error GoTo PlotLegendErrHand

  Const nudge = 50       'a little extra room
  Const linlen = 500     'length of the legend line
  Const boxedge = 200    'size of a legend color box
'
  Dim legTop As Single       'top of the Legend Box
  Dim legLeft As Single      'left side of Legend Box
  Dim legWidth As Single
  Dim legHeight As Single
  Dim legX As Single
  Dim legY As Single

  Dim i As Integer
  Dim s As String
  Dim tmp As Single
  Dim flag As Boolean

  '
  'First, see if we need to plot a legend at all
  '
  If (Not PS.Legend.On) Or (PS.PlotType = GQ_BAR) Then
    If Not dest Is Printer Then
      picLegend.Visible = False
    End If
    Exit Sub
  End If
  '
  'If there is no data/titles at all, there is no legend
  '
  flag = False
  For i = 0 To 4
    If (PD(i).n > 0) And (Trim$(PS.DataTitle(i)) <> "") Then
      flag = True
      Exit For
    End If
  Next
  If Not flag Then
    If Not dest Is Printer Then
      picLegend.Visible = False
    End If
    Exit Sub
  End If

  '
  'Turn on the legend picture box
  '
  If Not dest Is Printer Then
    picLegend.Visible = True
    picLegend.Cls
  End If

  '
  'Set the font
  '
  PlotSetFont PS.Legend.Font, picLegend

  '
  'Find the size of the legend
  '
  'loop through the data to figure out sizes
  legWidth = 0
  legHeight = nudge + nudge
'tbc  Select Case PS.PlotType
'tbc  Case GQ_XYPLOT
    For i = 0 To 4
      If (PD(i).n > 0) And (Trim$(PS.DataTitle(i)) <> "") Then
        s = PS.DataTitle(i)
        If dest Is Printer Then
          tmp = nudge + linlen + nudge + Printer.TextWidth(s) + nudge
          If tmp > legWidth Then legWidth = tmp
          legHeight = legHeight + Printer.TextHeight(s)
        Else
          tmp = nudge + linlen + nudge + picLegend.TextWidth(s) + nudge
          If tmp > legWidth Then legWidth = tmp
          legHeight = legHeight + picLegend.TextHeight(s)
        End If
      End If
    Next

  '
  'Size and position the Legend
  '
  If dest Is Printer Then
    legTop = picLegend.Top
    legLeft = picLegend.Left
    Printer.Line (legLeft, legTop)-Step(legWidth, legHeight), vbWhite, BF
    Printer.Line (legLeft, legTop)-Step(legWidth, legHeight), vbBlack, B
  Else
    legTop = 0
    legLeft = 0
    picLegend.Height = legHeight
    picLegend.Width = legWidth
    SetLegendPosition
  End If
  
  '
  'Place data in the legend
  '
'tbc  Select Case PS.PlotType
'tbc  Case GQ_XYPLOT            'sample lines and data selections
    legX = legLeft + nudge
    legY = legTop + nudge
    If dest Is Printer Then
      For i = 0 To 4
        If (PD(i).n > 0) And (Trim$(PS.DataTitle(i)) <> "") Then
          s = PS.DataTitle(i)
          Printer.ForeColor = PS.DataColor(i)
          Printer.DrawStyle = PS.DataStyle(i)
          Printer.CurrentX = legX
          Printer.CurrentY = legY + (0.5 * Printer.TextHeight(s))
          Printer.Line -Step(linlen, 0)

          Printer.ForeColor = 0
          Printer.DrawStyle = 0
          Printer.CurrentX = Printer.CurrentX + nudge
          Printer.CurrentY = legY
          Printer.Print s;

          legY = legY + (Printer.TextHeight(s))
        End If
      Next
    Else
      For i = 0 To 4
        If (PD(i).n > 0) And (Trim$(PS.DataTitle(i)) <> "") Then
          s = PS.DataTitle(i)
          picLegend.ForeColor = PS.DataColor(i)
          picLegend.DrawStyle = PS.DataStyle(i)
          picLegend.CurrentX = legX
          picLegend.CurrentY = legY + (0.5 * picLegend.TextHeight(s))
          picLegend.Line -Step(linlen, 0)

          picLegend.ForeColor = 0
          picLegend.DrawStyle = 0
          picLegend.CurrentX = picLegend.CurrentX + nudge
          picLegend.CurrentY = legY
          picLegend.Print s;

          legY = legY + (picLegend.TextHeight(s))
        End If
      Next
    End If
  
  If Not dest Is Printer Then
    With picLegend
      picArea.PaintPicture .Image, .Left, .Top
      picArea.Line (.Left, .Top)-(.Left + .Width, .Top + .Height), , B
    End With
  End If
  Exit Sub

PlotLegendErrHand:
  Select Case Err
    'place specific trapped errors here
  Case Else
    Select Case UnexpectedError("Plot,PlotLegend")
    Case vbAbort  'Abort - Stop the whole program
      End
    Case vbRetry  'Retry - Resume at the same line
      Resume
    Case vbIgnore 'Ignore - Resume at the next line
      Resume Next
    End Select
  End Select
End Sub

