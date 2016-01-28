VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export"
   ClientHeight    =   7065
   ClientLeft      =   3225
   ClientTop       =   3690
   ClientWidth     =   7875
   ForeColor       =   &H80000008&
   HelpContextID   =   1135
   Icon            =   "EXPORT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7065
   ScaleWidth      =   7875
   Begin VB.Frame Frame1 
      Caption         =   "Notes"
      Height          =   1695
      Left            =   120
      TabIndex        =   46
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtNotes 
         Height          =   1455
         HelpContextID   =   1504
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Text            =   "EXPORT.frx":030A
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame fraOpt 
      Caption         =   "Options"
      Height          =   855
      Left            =   5640
      TabIndex        =   43
      Top             =   1680
      Width           =   2175
      Begin VB.CheckBox cbxIncludeHeaders 
         Caption         =   "Include &Headers"
         Height          =   255
         HelpContextID   =   1135
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame fraExport 
      Caption         =   "Results for Export"
      Height          =   5295
      Left            =   120
      TabIndex        =   41
      Top             =   1680
      Width           =   5415
      Begin VB.CheckBox cbxExport 
         Caption         =   "Carrier"
         Height          =   195
         HelpContextID   =   1135
         Index           =   29
         Left            =   1440
         TabIndex        =   17
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Nonvolatiles"
         Height          =   195
         HelpContextID   =   1135
         Index           =   19
         Left            =   2640
         TabIndex        =   18
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Cumulative"
         Height          =   195
         HelpContextID   =   1135
         Index           =   28
         Left            =   4200
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Cumulative"
         Height          =   195
         HelpContextID   =   1135
         Index           =   27
         Left            =   4200
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Incremental"
         Height          =   195
         HelpContextID   =   1135
         Index           =   26
         Left            =   2760
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Incremental"
         Height          =   195
         HelpContextID   =   1135
         Index           =   25
         Left            =   2760
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Area Coverage"
         Height          =   195
         HelpContextID   =   1135
         Index           =   24
         Left            =   2640
         TabIndex        =   28
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Incremental"
         Height          =   195
         HelpContextID   =   1135
         Index           =   22
         Left            =   2760
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Cumulative"
         Height          =   195
         HelpContextID   =   1135
         Index           =   23
         Left            =   4200
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Distance Accountancy"
         Height          =   195
         HelpContextID   =   1135
         Index           =   21
         Left            =   120
         TabIndex        =   31
         Top             =   4800
         Width           =   3855
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Application Layout"
         Enabled         =   0   'False
         Height          =   195
         HelpContextID   =   1135
         Index           =   20
         Left            =   120
         TabIndex        =   23
         Top             =   3120
         Width           =   3855
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Cumulative"
         Height          =   195
         HelpContextID   =   1135
         Index           =   14
         Left            =   4200
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Incremental"
         Height          =   195
         HelpContextID   =   1135
         Index           =   13
         Left            =   2760
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Height Accountancy"
         Height          =   195
         HelpContextID   =   1135
         Index           =   18
         Left            =   120
         TabIndex        =   32
         Top             =   5040
         Width           =   3855
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Time Accountancy"
         Height          =   195
         HelpContextID   =   1135
         Index           =   17
         Left            =   120
         TabIndex        =   30
         Top             =   4560
         Width           =   3855
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Canopy Deposition"
         Height          =   195
         HelpContextID   =   1135
         Index           =   16
         Left            =   120
         TabIndex        =   29
         Top             =   4320
         Width           =   3855
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Deposition"
         Height          =   195
         HelpContextID   =   1135
         Index           =   15
         Left            =   1440
         TabIndex        =   27
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Cumulative"
         Height          =   195
         HelpContextID   =   1135
         Index           =   12
         Left            =   4200
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Incremental"
         Height          =   195
         HelpContextID   =   1135
         Index           =   11
         Left            =   2760
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Cumulative"
         Height          =   195
         HelpContextID   =   1135
         Index           =   10
         Left            =   4200
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Incremental"
         Height          =   195
         HelpContextID   =   1135
         Index           =   9
         Left            =   2760
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Mean Deposition"
         Height          =   195
         HelpContextID   =   1135
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   3495
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "&Fraction Aloft"
         Height          =   195
         HelpContextID   =   1135
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   3840
         Width           =   3495
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "&1 Hour Average Concentration"
         Height          =   195
         HelpContextID   =   1135
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   2880
         Width           =   3495
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Coefficient of Va&riation"
         Height          =   195
         HelpContextID   =   1135
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   3495
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "&Vertical Profile"
         Height          =   195
         HelpContextID   =   1135
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   3495
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "&Pond-Integrated Deposition (Std. EPA Pond)"
         Height          =   195
         HelpContextID   =   1135
         Index           =   3
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "&Deposition"
         Height          =   195
         HelpContextID   =   1135
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   3495
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Cumulative"
         Height          =   195
         HelpContextID   =   1135
         Index           =   1
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox cbxExport 
         Caption         =   "Incremental"
         Height          =   195
         HelpContextID   =   1135
         Index           =   0
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblExport 
         AutoSize        =   -1  'True
         Caption         =   "Settling Velocity"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   56
         Top             =   1920
         Width           =   1125
      End
      Begin VB.Label lblDSD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Canopy"
         Height          =   195
         Index           =   6
         Left            =   1680
         TabIndex        =   55
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label lblDSD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Spray Block"
         Height          =   195
         Index           =   5
         Left            =   1680
         TabIndex        =   54
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label lblDSD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vertical Profile"
         Height          =   195
         Index           =   4
         Left            =   1680
         TabIndex        =   53
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label lblDSD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Downwind"
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   52
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label lblDSD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Initial DSD 3"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   51
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lblDSD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Initial DSD 2"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   50
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label lblExport 
         AutoSize        =   -1  'True
         Caption         =   "Spray Block"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lblExport 
         AutoSize        =   -1  'True
         Caption         =   "Drop Size Distribution"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lblDSD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "lblDSD"
         Height          =   195
         Index           =   0
         Left            =   2190
         TabIndex        =   47
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraDelim 
      Caption         =   "Delimiter"
      Height          =   2655
      Left            =   5640
      TabIndex        =   42
      Top             =   2520
      Width           =   2175
      Begin VB.TextBox txtColumns 
         Height          =   285
         HelpContextID   =   1135
         Left            =   1200
         TabIndex        =   40
         Top             =   2280
         Width           =   855
      End
      Begin VB.OptionButton optDelim 
         Caption         =   "&Fixed Width"
         Height          =   255
         HelpContextID   =   1135
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtDelim 
         Height          =   285
         HelpContextID   =   1135
         Left            =   1320
         TabIndex        =   38
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton optDelim 
         Caption         =   "&Other"
         Height          =   255
         HelpContextID   =   1135
         Index           =   3
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optDelim 
         Caption         =   "Co&mma"
         Height          =   255
         HelpContextID   =   1135
         Index           =   2
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optDelim 
         Caption         =   "&Space"
         Height          =   255
         HelpContextID   =   1135
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optDelim 
         Caption         =   "&Tab"
         Height          =   255
         HelpContextID   =   1135
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Character:"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblColumns 
         Caption         =   "Columns:"
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1135
      Left            =   5880
      TabIndex        =   0
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1135
      Left            =   6720
      TabIndex        =   1
      Top             =   6600
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: export.frm,v 1.12 2001/08/13 17:40:01 tom Exp $

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOk_Click()
  If ExportData() Then Me.Hide
End Sub

Private Sub DataToForm()
'update form controls according to user data
  Dim c As Control
  Dim i As Integer
  Dim PlotVar As Long
  
  'Special treatment for Initial DSD 1 label
  Select Case UD.Tier
  Case TIER_1, TIER_2
    lblDSD(0).Caption = "Initial DSD"
  Case TIER_3
    lblDSD(0).Caption = "Initial DSD 1"
  End Select
  
  'Store the plot variable in the checkbox's tag
  cbxExport(0).Tag = PV_VFINC0
  cbxExport(25).Tag = PV_VFINC1
  cbxExport(26).Tag = PV_VFINC2
  cbxExport(1).Tag = PV_VFCUM0
  cbxExport(27).Tag = PV_VFCUM1
  cbxExport(28).Tag = PV_VFCUM2
  cbxExport(9).Tag = PV_DWDSDINC
  cbxExport(10).Tag = PV_DWDSDCUM
  cbxExport(11).Tag = PV_FXDSDINC
  cbxExport(12).Tag = PV_FXDSDCUM
  cbxExport(13).Tag = PV_SBDSDINC
  cbxExport(14).Tag = PV_SBDSDCUM
  cbxExport(22).Tag = PV_CNDSDINC
  cbxExport(23).Tag = PV_CNDSDCUM
  cbxExport(29).Tag = PV_SVTANK
  cbxExport(19).Tag = PV_SVNONV
  cbxExport(2).Tag = PV_DEP
  cbxExport(3).Tag = PV_PID
  cbxExport(4).Tag = PV_VERT
  cbxExport(5).Tag = PV_CONC
  cbxExport(20).Tag = PV_LAYOUT
  cbxExport(6).Tag = PV_COV
  cbxExport(7).Tag = PV_MEAN
  cbxExport(8).Tag = PV_FA
  cbxExport(15).Tag = PV_SBDEP
  cbxExport(24).Tag = PV_SBCOVER
  cbxExport(16).Tag = PV_CANDEP
  cbxExport(17).Tag = PV_TA
  cbxExport(21).Tag = PV_DA
  cbxExport(18).Tag = PV_HA
  'some labels too
  lblDSD(0).Tag = PV_VFINC0
  lblDSD(1).Tag = PV_VFINC1
  lblDSD(2).Tag = PV_VFINC2
  lblDSD(3).Tag = PV_DWDSDINC
  lblDSD(4).Tag = PV_FXDSDINC
  lblDSD(5).Tag = PV_SBDSDINC
  lblDSD(6).Tag = PV_CNDSDINC

  'Initialize the checkboxes
  For Each c In cbxExport()
    PlotVar = CLng(c.Tag) 'Tag holds PV_
    c.Value = 0 'unchecked
    c.Visible = PlotIsAvailable(PlotVar)
    c.Enabled = PlotIsAvailableExtended(PlotVar)
  Next
  For Each c In lblDSD()
    PlotVar = CLng(c.Tag) 'Tag holds PV_
    c.Visible = PlotIsAvailable(PlotVar)
    c.Enabled = PlotIsAvailableExtended(PlotVar)
  Next
  'Settling Velocity label
  lblExport(2).Visible = cbxExport(29).Visible Or cbxExport(19).Visible
  lblExport(2).Enabled = cbxExport(29).Enabled Or cbxExport(19).Enabled
  'Spray Block label
  lblExport(1).Visible = cbxExport(15).Visible Or cbxExport(24).Visible
  lblExport(1).Enabled = cbxExport(15).Enabled Or cbxExport(24).Enabled
End Sub

Private Function ExportData() As Integer
'generate and export plot data
'This function uses the same functions as the plotting form

  Dim i As Integer
  Dim s As String
  ReDim buffer(1) As String
  Dim cbx As CheckBox
  Dim opt As OptionButton
  Dim fn As String
  Dim np As Integer
  Dim idelim As Integer
  Dim icbox As Integer
  Dim ivar As Long
  Dim icol As Integer
  Dim dlm As Integer
  Dim stat As Integer
  Dim cols As Integer
  Dim start As Integer
  Dim warnheaders As Integer
  Dim warnvalues As Integer
  Dim DoHeaders As Integer
  ReDim X(MAX_CALCDATA) As Single
  ReDim Y(MAX_CALCDATA) As Single
  ReDim Lbl(MAX_CALCDATA) As String

  'make sure at least one category was selected
  stat = False
  For Each cbx In cbxExport()
    If cbx.Value = 1 Then
      stat = True
      Exit For
    End If
  Next
  If Not stat Then
    MsgBox "No data selected for export.", vbExclamation
    ExportData = False
    Exit Function
  End If

  'recover form settings
  'Delimiter
  For Each opt In optDelim()
    If opt.Value Then
      idelim = opt.Index
      Exit For
    End If
  Next
  Select Case idelim
    Case 0  'tab
     dlm = 9
    Case 1  'space
      dlm = 32
    Case 2  'comma
      dlm = 44
    Case 3  'user-defined
      'the text is either a single delimiter character
      'or a two-char sequence representing a nonprintable
      Select Case Len(txtDelim.Text)
        Case 1
          dlm = Asc(txtDelim.Text)
        Case 2
          dlm = Asc(Right$(txtDelim.Text, 1)) - 64
        Case Else 'space
          dlm = 32
      End Select
    Case 4  'fixed columns
      cols = Val(txtColumns.Text)
      If cols > 0 Then
        dlm = -1  'negative dlm signals fixed columns
      Else
        MsgBox "Bad column width.", vbCritical + vbOKOnly
        ExportData = False
        Exit Function
      End If
  End Select
  'include header info?
  DoHeaders = False
  If cbxIncludeHeaders.Value = 1 Then DoHeaders = True

  'get an output file
  If Not FileDialog(FD_SAVEAS, FD_TYPE_TEXT, fn) Then
    ExportData = False
    Exit Function
  End If
  On Error GoTo ErrHandlerED
  Open fn For Output As #1

  'include the RunID, regardless of the header switch setting
  Print #1, "# "; GetRunID()
  
  'output the notes
  If txtNotes.Text <> "" Then
    start = 1
    Do
      Print #1, "# "; LineFromString(txtNotes.Text, start)
    Loop While start > 0
  End If
  
  'generate plot data for export
  warnheaders = False
  warnvalues = False
  For Each cbx In cbxExport()
    If cbx.Value = 1 Then
      ivar = CLng(cbx.Tag)
      If GenPlotTitles(ivar, False) And _
         GenPlotUnits(ivar) And _
         GenPlotDataUDUC(UD, UC, ivar, True, np, X(), Y(), Lbl()) Then
        'headers
        If DoHeaders Then
          Print #1, "# "; PS.PlotTitle.Text
          buffer(0) = PS.XTitle.Text + PS.Xunits
          buffer(1) = PS.YTitle.Text + PS.Yunits
          If dlm < 0 Then  'fixed width columns
            s = ""
            For i = 0 To 1
              If Len(buffer(i)) > cols Then
                buffer(i) = Left$(buffer(i), cols) 'truncate header
                warnheaders = True  'set flag for warning display
              ElseIf Len(buffer(i)) < cols Then
                buffer(i) = Space$(cols - Len(buffer(i))) + buffer(i) 'pad header
              End If
              s = s + buffer(i)
            Next
          ElseIf dlm = 32 Then 'for a space delimiter, quote the headers
            s = Chr$(34) + buffer(0) + Chr$(34) + Chr$(dlm) + Chr$(34) + buffer(1) + Chr$(34)
          Else  'other single-character delimiter
            s = buffer(0) + Chr$(dlm) + buffer(1)
          End If
          Print #1, "# "; s
        End If
        'data
        For i = 0 To np - 1
          If dlm >= 0 Then
            s = Format$(X(i)) + Chr$(dlm) + Format$(Y(i))
          Else 'fixed-width columns
            buffer(0) = Format$(X(i))
            buffer(1) = Format$(Y(i))
            s = ""
            For icol = 0 To 1
              If Len(buffer(icol)) > cols Then
                buffer(icol) = String$(cols, "#") 'display overflow
                warnvalues = True  'set flag for warning display
              ElseIf Len(buffer(icol)) < cols Then
                buffer(icol) = Space$(cols - Len(buffer(icol))) + buffer(icol) 'pad number
              End If
              s = s + buffer(icol)
            Next
          End If
          Print #1, s
        Next
      End If
    End If
  Next
  
  'check for warnings
  If warnheaders Then
    s = "Warning: Some headers were too wide for the number of columns specified."
    MsgBox s, vbExclamation + vbOKOnly
  End If
  If warnvalues Then
    s = "Warning: Some values were too wide for the number of columns specified."
    MsgBox s, vbExclamation + vbOKOnly
  End If

  'success!
  Close #1
  ExportData = True
  Exit Function

ErrHandlerED:
  s = "Error writing file: " + fn + Chr$(13) + Error$(Err)
  t% = vbCritical + vbOKOnly
  MsgBox s, t%
  ExportData = False
  Exit Function
End Function

Private Sub Form_Activate()
  DataToForm
End Sub

Private Sub Form_Load()
  InitForm
End Sub

Private Sub InitForm()
  CenterForm Me  'center the form
  txtNotes.Text = ""
  txtColumns.Text = "10"
End Sub

Private Sub txtDelim_Click()
  optDelim(3).Value = True
End Sub

Private Sub txtDelim_KeyPress(KeyAscii As Integer)
'Trap all ascii characters and place them in the text
'make the non-printing one printable
  Select Case KeyAscii
    Case 0 To 31  'nonprintables
      txtDelim.Text = "^" + Chr$(KeyAscii + 64)
      KeyAscii = 0
    Case 32 To 126
      txtDelim.Text = Chr$(KeyAscii)
      KeyAscii = 0
  End Select
End Sub

