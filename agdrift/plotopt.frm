VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlotOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plot Options"
   ClientHeight    =   5835
   ClientLeft      =   1905
   ClientTop       =   1920
   ClientWidth     =   7065
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "PLOTOPT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5835
   ScaleWidth      =   7065
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "&Defaults"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1220
      Left            =   5280
      TabIndex        =   0
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1220
      Left            =   6240
      TabIndex        =   1
      Top             =   5400
      Width           =   735
   End
   Begin VB.Frame fraData 
      Caption         =   "Data"
      ClipControls    =   0   'False
      Height          =   2775
      Left            =   120
      TabIndex        =   51
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtDataTitle 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txtDataTitle 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtDataTitle 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtDataTitle 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtDataTitle 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox cboDataSource 
         Height          =   315
         HelpContextID   =   1220
         Index           =   0
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   375
         HelpContextID   =   1220
         Left            =   120
         TabIndex        =   24
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox cboDataSource 
         Height          =   315
         HelpContextID   =   1220
         Index           =   4
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1920
         Width           =   2175
      End
      Begin VB.ComboBox cboDataSource 
         Height          =   315
         HelpContextID   =   1220
         Index           =   3
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1560
         Width           =   2175
      End
      Begin VB.ComboBox cboDataSource 
         Height          =   315
         HelpContextID   =   1220
         Index           =   2
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cboDataSource 
         Height          =   315
         HelpContextID   =   1220
         Index           =   1
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   2175
      End
      Begin VB.PictureBox picColor 
         Height          =   300
         Index           =   4
         Left            =   4920
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   300
      End
      Begin VB.PictureBox picColor 
         Height          =   300
         Index           =   3
         Left            =   4920
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   300
      End
      Begin VB.PictureBox picColor 
         Height          =   300
         Index           =   2
         Left            =   4920
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   300
      End
      Begin VB.PictureBox picColor 
         Height          =   300
         Index           =   1
         Left            =   4920
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         Top             =   840
         Width           =   300
      End
      Begin VB.PictureBox picColor 
         Height          =   300
         Index           =   0
         Left            =   4920
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   480
         Width           =   300
      End
      Begin VB.ComboBox cboStyle 
         Height          =   315
         HelpContextID   =   1220
         Index           =   4
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1920
         Width           =   1455
      End
      Begin VB.ComboBox cboStyle 
         Height          =   315
         HelpContextID   =   1220
         Index           =   3
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cboStyle 
         Height          =   315
         HelpContextID   =   1220
         Index           =   2
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cboStyle 
         Height          =   315
         HelpContextID   =   1220
         Index           =   1
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cboStyle 
         Height          =   315
         HelpContextID   =   1220
         Index           =   0
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Title"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         Height          =   195
         Left            =   4800
         TabIndex        =   58
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblStyle 
         AutoSize        =   -1  'True
         Caption         =   "Style"
         Height          =   195
         Left            =   5760
         TabIndex        =   56
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblDS 
         AutoSize        =   -1  'True
         Caption         =   "Data Source"
         Height          =   195
         Left            =   2640
         TabIndex        =   57
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Frame fraYaxis 
      Caption         =   "Y axis"
      ClipControls    =   0   'False
      Height          =   3015
      Left            =   2160
      TabIndex        =   46
      Top             =   2760
      Width           =   1935
      Begin VB.CheckBox cbxYlog 
         Caption         =   "Log"
         Height          =   255
         HelpContextID   =   1220
         Left            =   1080
         TabIndex        =   41
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton cmdYscaleFont 
         Caption         =   "Scale &Font"
         Height          =   375
         HelpContextID   =   1220
         Left            =   480
         TabIndex        =   42
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CheckBox cbxYgrid 
         Caption         =   "Grid"
         Height          =   255
         HelpContextID   =   1220
         Left            =   240
         TabIndex        =   40
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtYminorTics 
         Height          =   285
         HelpContextID   =   1220
         Left            =   1080
         TabIndex        =   39
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtYinc 
         Height          =   285
         HelpContextID   =   1220
         Left            =   1080
         TabIndex        =   38
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtYmax 
         Height          =   285
         HelpContextID   =   1220
         Left            =   1080
         TabIndex        =   37
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtYmin 
         Height          =   285
         HelpContextID   =   1220
         Left            =   1080
         TabIndex        =   36
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optYAutoSc 
         Caption         =   "Fixed"
         Height          =   195
         HelpContextID   =   1220
         Index           =   1
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optYAutoSc 
         Caption         =   "Auto"
         Height          =   195
         HelpContextID   =   1220
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label lblYminorTics 
         AutoSize        =   -1  'True
         Caption         =   "Minor Tics:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label lblYinc 
         AutoSize        =   -1  'True
         Caption         =   "Incr:"
         Height          =   195
         Left            =   555
         TabIndex        =   49
         Top             =   1485
         Width           =   405
      End
      Begin VB.Label lblYmax 
         AutoSize        =   -1  'True
         Caption         =   "Max:"
         Height          =   195
         Left            =   540
         TabIndex        =   48
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label lblYmin 
         AutoSize        =   -1  'True
         Caption         =   "Min:"
         Height          =   195
         Left            =   600
         TabIndex        =   47
         Top             =   780
         Width           =   375
      End
   End
   Begin VB.Frame fraXaxis 
      Caption         =   "X axis"
      ClipControls    =   0   'False
      Height          =   3015
      Left            =   120
      TabIndex        =   45
      Top             =   2760
      Width           =   1935
      Begin VB.CheckBox cbxXlog 
         Caption         =   "Log"
         Height          =   255
         HelpContextID   =   1220
         Left            =   1080
         TabIndex        =   32
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtXminorTics 
         Height          =   285
         HelpContextID   =   1220
         Left            =   1080
         TabIndex        =   30
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtXinc 
         Height          =   285
         HelpContextID   =   1220
         Left            =   1080
         TabIndex        =   29
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtXmax 
         Height          =   285
         HelpContextID   =   1220
         Left            =   1080
         TabIndex        =   28
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtXmin 
         Height          =   285
         HelpContextID   =   1220
         Left            =   1080
         TabIndex        =   27
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdXscaleFont 
         Caption         =   "Scale &Font"
         Height          =   375
         HelpContextID   =   1220
         Left            =   480
         TabIndex        =   33
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CheckBox cbxXgrid 
         Caption         =   "Grid"
         Height          =   255
         HelpContextID   =   1220
         Left            =   240
         TabIndex        =   31
         Top             =   2160
         Width           =   735
      End
      Begin VB.OptionButton optXAutoSc 
         Caption         =   "Fixed"
         Height          =   195
         HelpContextID   =   1220
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optXAutoSc 
         Caption         =   "Auto"
         Height          =   195
         HelpContextID   =   1220
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label lblXminorTics 
         AutoSize        =   -1  'True
         Caption         =   "Minor Tics:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label lblXinc 
         AutoSize        =   -1  'True
         Caption         =   "Incr:"
         Height          =   195
         Left            =   540
         TabIndex        =   53
         Top             =   1485
         Width           =   405
      End
      Begin VB.Label lblXmax 
         AutoSize        =   -1  'True
         Caption         =   "Max:"
         Height          =   195
         Left            =   540
         TabIndex        =   54
         Top             =   1110
         Width           =   420
      End
      Begin VB.Label lblXmin 
         AutoSize        =   -1  'True
         Caption         =   "Min:"
         Height          =   195
         Left            =   600
         TabIndex        =   55
         Top             =   795
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Legend"
      Height          =   1335
      Left            =   4200
      TabIndex        =   60
      Top             =   2760
      Width           =   2775
      Begin VB.CheckBox cbxShowLegend 
         Caption         =   "Show Legend"
         Height          =   255
         Left            =   720
         TabIndex        =   43
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdLegendFont 
         Caption         =   "Legend &Font"
         Height          =   375
         HelpContextID   =   1220
         Left            =   720
         TabIndex        =   44
         Top             =   840
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPlotOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: plotopt.frm,v 1.8 2001/04/26 16:21:58 tom Exp $
'Plot Options form
'
'PURPOSE
'Alter axis, color, linetype, data title and source settings for
'current and future plotting with the frmPlot form.
'
'INPUT
'
'OUTPUT
'The Tag property contains status information.
' "regen" = OK was pressed, data needs to be regenerated
' "replot" = OK was pressed, just replot existing data
' (blank) = user cancelled, do nothing
'
'USAGE
'show the form modally, then query the Tag property.
'=========================================================

'form-wide vars

'this flag is used to tell some controls not to
'take action on their new values. This is required
'to differentiate between programatic state changes
'and user actions
Dim PropTakeAction As Integer 'if true, execute automatic change-related code
                              'for Property text boxes
Dim RegenData As Integer 'true if plot data needs to be regenerated
Dim tmpPS As PlotSettingData 'local copy of plot settings

Private Sub cboDataSource_Click(Index As Integer)
  
' SourceID   -1=No Data                    ("")
'             0=Current Data               ("Current Data:")
'             1=Saved Results              ("File: filename")
'             2=Tier I Library             ("Tier1Lib: key")
'             3=Dropsize Library           ("Lib: key")
'             4=Field Trial Prediction     ("TrialPred: key")
'             5=Field Trial Measurement    ("TrialMeas: key")
'             6=Toolbox Plot Data          ("ToolboxData:")
  Dim SourceID As Integer
  Dim s As String
  Dim DataSource As String
  Dim DataTitle As String
   
  If PropTakeAction Then
    SourceID = cboDataSource(Index).ItemData(cboDataSource(Index).ListIndex)
    Select Case SourceID
    Case SID_NONE
      tmpPS.DataSource(Index) = ""
      txtDataTitle(Index) = ""
    Case SID_CURRENT
      tmpPS.DataSource(Index) = "Current Data:"
      txtDataTitle(Index) = "Current Data"
    Case SID_FILE
      If PlotFileDialog(s) Then
        tmpPS.DataSource(Index) = "File: " + LCase$(s)
        txtDataTitle(Index) = s
      End If
    Case SID_T1LIB
      If Tier1LibDialog(DataSource, DataTitle) Then
        tmpPS.DataSource(Index) = DataSource
        txtDataTitle(Index) = DataTitle
      End If
    Case SID_LIB
      If DropLibDialog(s) Then 'returns a library key
        tmpPS.DataSource(Index) = "Lib: " & s
        txtDataTitle(Index) = s
      End If
    Case SID_FTPRED
      If TrialLibDialog(s) Then
        tmpPS.DataSource(Index) = "TrialPred: " & s
        txtDataTitle(Index) = "(Pred.) " & s
      End If
    Case SID_FTMEAS
      If TrialLibDialog(s) Then
        tmpPS.DataSource(Index) = "TrialMeas: " & s
        txtDataTitle(Index) = "(Meas.) " & s
      End If
    Case SID_TPD
      tmpPS.DataSource(Index) = "ToolboxData:"
      txtDataTitle(Index) = tmpPS.DataTitle(Index)
    End Select
    RegenData = True 'must regen plot data
  End If
End Sub

Private Function CheckFormData(xPS As PlotSettingData) As Integer
'Sanity check for Proposed plot options
'
' returns true if plot settings are ok, false if not
' sets RegenData if necessary
'
  Dim problem As Integer

  CheckFormData = True 'default value
  
  'check for same min/max on manual scales
  problem = False
  s = "Scale minimum and maximum must be different."
  If Not xPS.Xauto Then
    If Abs(xPS.Xmax - xPS.Xmin) < 1E-36 Then problem = True
  End If
  If Not xPS.Yauto Then
    If Abs(xPS.Ymax - xPS.Ymin) < 1E-36 Then problem = True
  End If
  If problem Then MsgBox s, vbCritical: CheckFormData = False
  
  'log/linear toggle must cause data regen in case data
  'was/needs-to-be removed for log scales
  If xPS.Xlog <> PS.Xlog Then RegenData = True
  If xPS.Ylog <> PS.Ylog Then RegenData = True

  'log scale limit checks for manual scales
  problem = False
  s = "Zero or negative axis limits not allowed for log scales."
  If xPS.Xlog And Not xPS.Xauto Then
    If xPS.Xmin < 1E-36 Or xPS.Xmax < 1E-36 Then problem = True
  End If
  If xPS.Ylog And Not xPS.Yauto Then
    If xPS.Ymin < 1E-36 Or xPS.Ymax < 1E-36 Then problem = True
  End If
  If problem Then MsgBox s, vbCritical: CheckFormData = False

  'log scale increment checks for manual scales
  problem = False
  s = "Log scale increment must be an integer greater than 0."
  If xPS.Xlog And Not xPS.Xauto Then
    If xPS.Xinc < 1 Then problem = True
  End If
  If xPS.Ylog And Not xPS.Yauto Then
    If xPS.Yinc < 1 Then problem = True
  End If
  If problem Then MsgBox s, vbCritical: CheckFormData = False
End Function

Private Sub cmdCancel_Click()
  Me.Tag = ""
  Me.Hide
End Sub

Private Sub cmdClear_Click()
  cboDataSource(0).ListIndex = 0 'Select Current data
  For i = 1 To 4
    cboDataSource(0).ListIndex = 0
  Next
End Sub

Private Sub cmdDefaults_Click()
'Revert local Plot Settings to default values.
  Dim savRunTitleText As String
  Dim savPlotTitleText As String
  Dim savCaption As String
  Dim savHelpID As Integer
  Dim savXTitleText As String
  Dim savYTitleText As String
  
  'save some values that shouldn't be defaulted
  savRunTitleText = tmpPS.RunTitle.Text
  savPlotTitleText = tmpPS.PlotTitle.Text
  savCaption = tmpPS.Caption
  savHelpID = tmpPS.HelpID
  savXTitleText = tmpPS.XTitle.Text
  savYTitleText = tmpPS.YTitle.Text
  
  'default the local setting data
  PlotSettingsInit tmpPS
  
  'restore the saved values that shouldn't have been defaulted
  tmpPS.RunTitle.Text = savRunTitleText
  tmpPS.PlotTitle.Text = savPlotTitleText
  tmpPS.Caption = savCaption
  tmpPS.HelpID = savHelpID
  tmpPS.XTitle.Text = savXTitleText
  tmpPS.YTitle.Text = savYTitleText
 
  'If this copy of the settings data gets transferred to
  'permanent storage, it will be changed from what the
  'user had before. Therefore mark this data as changed.
  tmpPS.Changed = True
  
  'transfer newly defaulted data to form controls
  DataToForm tmpPS

  'The plot data must be regenerated in case there was
  'more than one data set selected.
  RegenData = True 'must regen plot data
End Sub

Private Sub cmdLegendFont_Click()
  FontDialog tmpPS.Legend.Font
End Sub

Private Sub cmdOk_Click()
'Make the new plot settings permanent
  FormToData tmpPS  'transfer form data to local storage
  'Do a sanity check on the data
  If CheckFormData(tmpPS) Then
    tmpPS.Changed = True
    PS = tmpPS 'Copy to permanent storage
    If RegenData Then
      Me.Tag = "regen"
    Else
      Me.Tag = "replot"
    End If
    Me.Hide
  End If
End Sub

Private Sub cmdSave_Click()
'Save current plot settings in the prefs file
  FormToData tmpPS  'transfer form data to local storage
  'Do a sanity check on the data
  If CheckFormData(tmpPS) Then
    tmpPS.Changed = False
    PS = tmpPS 'Copy to permanent storage
    PlotPrefsWrite PS 'Save the settings in the prefs file
    If RegenData Then
      Me.Tag = "regen"
    Else
      Me.Tag = "replot"
    End If
    Me.Hide
  End If
End Sub

Private Sub cmdXscaleFont_Click()
  FontDialog tmpPS.XScaleFont
End Sub

Private Sub cmdYscaleFont_Click()
  FontDialog tmpPS.YScaleFont
End Sub

Private Sub DataToForm(xPS As PlotSettingData)
'Update the status of the plot option controls based
'values of the given xPS structure
'
  Dim s As String
  Dim Param1 As String
  Dim Param2 As String
  Dim Param3 As String
  Dim Param4 As String
  Dim i As Integer
  Dim j As Integer
  Dim SourceID As Integer
  
  PropTakeAction = False 'turn off control changes

  'Data
  For i = 0 To 4
    txtDataTitle(i).Text = xPS.DataTitle(i)
    SourceID = SourceToSourceID(xPS.DataSource(i), Param1, Param2, Param3, Param4)
    For j = 0 To cboDataSource(i).ListCount - 1  'Match SourceID's
      If cboDataSource(i).ItemData(j) = SourceID Then
        cboDataSource(i).ListIndex = j
        Exit For
      End If
    Next
    picColor(i).BackColor = xPS.DataColor(i)
    cboStyle(i).ListIndex = xPS.DataStyle(i)
  Next
  'filter data sources
  For i = 0 To 4
    SourceID = SourceToSourceID(xPS.DataSource(i), Param1, Param2, Param3, Param4)
    If Not SourceIsValid(UI.PlotVar, SourceID) Then
      cboDataSource(i).ListIndex = 0
    End If
  Next
  'legend
  If xPS.Legend.On Then
    cbxShowLegend.Value = 1
  Else
    cbxShowLegend.Value = 0
  End If
  'X axis
  If xPS.Xauto Then
    optXAutoSc(0).Value = True
    optXAutoSc(1).Value = False
    XDisableScaleControls   'Dim the scale controls
  Else
    optXAutoSc(0).Value = False
    optXAutoSc(1).Value = True
    XEnableScaleControls    'Undim the scale controls
  End If
  txtXmin.Text = xPS.Xmin
  txtXmax.Text = xPS.Xmax
  txtXinc.Text = xPS.Xinc
  txtXminorTics.Text = xPS.XminorTics
  cbxXgrid.Value = xPS.Xgrid
  cbxXlog.Value = xPS.Xlog
  'Y axis
  If xPS.Yauto Then
    optYAutoSc(0).Value = True
    optYAutoSc(1).Value = False
    YDisableScaleControls   'Dim the scale controls
  Else
    optYAutoSc(0).Value = False
    optYAutoSc(1).Value = True
    YEnableScaleControls    'Undim the scale controls
  End If
  txtYmin.Text = xPS.Ymin
  txtYmax.Text = xPS.Ymax
  txtYinc.Text = xPS.Yinc
  txtYminorTics.Text = xPS.YminorTics
  cbxYgrid.Value = xPS.Ygrid
  cbxYlog.Value = xPS.Ylog

  PropTakeAction = True 'Reactivate form controls
End Sub

Private Function DropLibDialog(key)
  Load frmDropLib
'tbc support both SDTF/FS libraries
  If UD.Smokey = AUD_SDTF Then
    frmDropLib.SelectTable ""
  Else
    frmDropLib.SelectTable "FS"
  End If
  frmDropLib.Show vbModal  'get the mass fractions from lib
  key = frmDropLib.Tag 'retrieve library key
  Unload frmDropLib
  If key <> "" Then
    DropLibDialog = True
  Else
    DropLibDialog = False
  End If
End Function

Private Sub FontDialog(FD As FontData)
  'Load up Dialog Selections from Data label on form
  CMDialog1.FontName = FD.Name
  CMDialog1.FontSize = FD.Size
  CMDialog1.FontBold = FD.Bold
  CMDialog1.FontItalic = FD.Italic
  CMDialog1.FontUnderline = FD.Underline
  CMDialog1.FontStrikethru = FD.Strikethru
  CMDialog1.Color = FD.Color

  'Set cancel to true
  CMDialog1.CancelError = True
  On Error GoTo ErrHandler
  'Set the cdlCFBoth and cdlCFEffects flags
  CMDialog1.Flags = cdlCFBoth Or cdlCFEffects
  'Display the Font dialog box
  CMDialog1.Action = 4
  'Set text properties according to user's selections
  FD.Name = CMDialog1.FontName
  FD.Size = CMDialog1.FontSize
  FD.Bold = CMDialog1.FontBold
  FD.Italic = CMDialog1.FontItalic
  FD.Underline = CMDialog1.FontUnderline
  FD.Strikethru = CMDialog1.FontStrikethru
  FD.Color = CMDialog1.Color

ErrHandler:
  'User pressed Cancel
  Exit Sub
End Sub

Private Sub Form_Load()
  InitForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'prevent the form from being unloaded by the close box.
'Act as though the Cancel button was hit.
  If UnloadMode = 0 Then 'User selected CLose from control box
    Cancel = True 'stop the unloading
    cmdCancel_Click 'simulate cancel button
  End If
End Sub

Private Sub FormToData(xPS As PlotSettingData)
'Transfer form values to PS structure
'
'Data
  For i = 0 To 4
    xPS.DataTitle(i) = txtDataTitle(i).Text
    xPS.DataSource(i) = tmpPS.DataSource(i)
    xPS.DataColor(i) = picColor(i).BackColor
    xPS.DataStyle(i) = cboStyle(i).ListIndex
  Next
  If cbxShowLegend.Value = 1 Then
    xPS.Legend.On = True
  Else
    xPS.Legend.On = False
  End If
'X axis
  If optXAutoSc(0).Value Then
    xPS.Xauto = True
  Else
    xPS.Xauto = False
  End If
  xPS.Xmin = Val(txtXmin.Text)
  xPS.Xmax = Val(txtXmax.Text)
  xPS.Xinc = Val(txtXinc.Text)
  xPS.XminorTics = Val(txtXminorTics.Text)
  xPS.Xgrid = cbxXgrid.Value
  xPS.Xlog = cbxXlog.Value
  'For manual Log scales make sure the increment is an integer
  If xPS.Xlog And Not xPS.Xauto Then
    xPS.Xinc = Int(xPS.Xinc)
  End If
  'Y axis
  If optYAutoSc(0).Value Then
    xPS.Yauto = True
  Else
    xPS.Yauto = False
  End If
  xPS.Ymin = Val(txtYmin.Text)
  xPS.Ymax = Val(txtYmax.Text)
  xPS.Yinc = Val(txtYinc.Text)
  xPS.YminorTics = Val(txtYminorTics.Text)
  xPS.Ygrid = cbxYgrid.Value
  xPS.Ylog = cbxYlog.Value
  'For manual Log scales make sure the increment is an integer
  If xPS.Ylog And Not xPS.Yauto Then
    xPS.Yinc = Int(xPS.Yinc)
  End If
End Sub

Private Sub InitForm()
'initialize this form
  
  'Center the form on the screen
  Me.Left = (Screen.Width / 2) - (Me.Width / 2)
  Me.Top = (Screen.Height / 2) - (Me.Height / 2)
  
  'data source combo boxes
  For i = 0 To 4
    cboDataSource(i).AddItem "<none>"
    cboDataSource(i).ItemData(cboDataSource(i).NewIndex) = SID_NONE
    If SourceIsValid(UI.PlotVar, SID_CURRENT) Then
      cboDataSource(i).AddItem "Current Data"
      cboDataSource(i).ItemData(cboDataSource(i).NewIndex) = SID_CURRENT
    End If
    If SourceIsValid(UI.PlotVar, SID_FILE) Then
      cboDataSource(i).AddItem "Saved Results"
      cboDataSource(i).ItemData(cboDataSource(i).NewIndex) = SID_FILE
    End If
    If SourceIsValid(UI.PlotVar, SID_T1LIB) Then
      cboDataSource(i).AddItem "Tier I Library"
      cboDataSource(i).ItemData(cboDataSource(i).NewIndex) = SID_T1LIB
    End If
    If SourceIsValid(UI.PlotVar, SID_LIB) Then
      cboDataSource(i).AddItem "Dropsize Library Entry"
      cboDataSource(i).ItemData(cboDataSource(i).NewIndex) = SID_LIB
    End If
    If SourceIsValid(UI.PlotVar, SID_FTPRED) Then
      cboDataSource(i).AddItem "Field Trial Prediction"
      cboDataSource(i).ItemData(cboDataSource(i).NewIndex) = SID_FTPRED
    End If
    If SourceIsValid(UI.PlotVar, SID_FTMEAS) Then
      cboDataSource(i).AddItem "Field Trial Measurement"
      cboDataSource(i).ItemData(cboDataSource(i).NewIndex) = SID_FTMEAS
    End If
    If SourceIsValid(UI.PlotVar, SID_TPD) Then
      For j = 0 To TPD.NC - 1
        cboDataSource(i).AddItem "Toolbox Plot Data Column " + Format$(j)
        cboDataSource(i).ItemData(cboDataSource(i).NewIndex) = SID_TPD
      Next
    End If
  Next
  
  'Style combo boxes
  For i = 0 To 4
    cboStyle(i).AddItem "Solid"
    cboStyle(i).AddItem "Dash"
    cboStyle(i).AddItem "Dot"
    cboStyle(i).AddItem "Dash-Dot"
    cboStyle(i).AddItem "Dash-Dot-Dot"
    cboStyle(i).AddItem "Transparent"
    cboStyle(i).AddItem "Circles"
  Next

  'Init Form-wide flags
  PropTakeAction = True
  RegenData = False 'true if plot data needs to be regenerated

  'copy plot settings to local storage
  tmpPS = PS
  'Transfer data to form ontrols
  DataToForm tmpPS
End Sub

Private Sub optXAutoSc_Click(Index As Integer)
  If PropTakeAction Then
    If Index = 0 Then            'Auto Scale
      XDisableScaleControls
    Else                         'Fixed Scale
      XEnableScaleControls
    End If
  End If
End Sub

Private Sub optYAutoSc_Click(Index As Integer)
  If PropTakeAction Then
    If Index = 0 Then            'Auto Scale
      YDisableScaleControls
    Else                         'Fixed Scale
      YEnableScaleControls
    End If
  End If
End Sub

Private Sub picColor_Click(Index As Integer)
  
  'CancelError is True.
  On Error GoTo picColorHandler
    
  CMDialog1.Color = tmpPS.DataColor(Index) ' Set initial color selection for dialog.
  CMDialog1.Flags = cdlCCRGBInit
  CMDialog1.Action = 3    ' Display color dialog.
  picColor(Index).BackColor = CMDialog1.Color
  Exit Sub

picColorHandler:
'User pressed cancel button
  Exit Sub

End Sub

Private Function PlotFileDialog(s As String) As Integer
'Dialog box for obtaining a file name
' -returns true on OK, false on Cancel
'
  Dim d As Control

  Set d = CMDialog1
  '
  On Error GoTo ErrHandlerPFD
  'Turn on CancelError
  d.CancelError = True
  'Set Default Extension
  'added if a file name is entered without an extension
  d.DefaultExt = "agd"
  'Set filter list
  d.Filter = "All Files (*.*)|*.*|Data Files (*.agd)|*.agd"
  'Specify current filter
  d.FilterIndex = 2
  'Set the default file name
  d.FileName = s
  'Set dialog flags
  d.Flags = cdlOFNHideReadOnly
  'Display the dialog box
  d.ShowOpen
  'full file path is CMDialog1.FileName
  'file name only is CMDialog1.FileTitle
  s = d.FileName
  PlotFileDialog = True
  Exit Function

ErrHandlerPFD:
'User pressed cancel button
  s = ""
  PlotFileDialog = False
  Exit Function
End Function

Private Sub XDisableScaleControls()
  lblXmin.Enabled = False
  txtXmin.Enabled = False
  lblXmax.Enabled = False
  txtXmax.Enabled = False
  lblXinc.Enabled = False
  txtXinc.Enabled = False
  lblXminorTics.Enabled = False
  txtXminorTics.Enabled = False
End Sub

Private Sub XEnableScaleControls()
  lblXmin.Enabled = True
  txtXmin.Enabled = True
  lblXmax.Enabled = True
  txtXmax.Enabled = True
  lblXinc.Enabled = True
  txtXinc.Enabled = True
  lblXminorTics.Enabled = True
  txtXminorTics.Enabled = True
End Sub

Private Sub YDisableScaleControls()
  lblYmin.Enabled = False
  txtYmin.Enabled = False
  lblYmax.Enabled = False
  txtYmax.Enabled = False
  lblYinc.Enabled = False
  txtYinc.Enabled = False
  lblYminorTics.Enabled = False
  txtYminorTics.Enabled = False
End Sub

Private Sub YEnableScaleControls()
  lblYmin.Enabled = True
  txtYmin.Enabled = True
  lblYmax.Enabled = True
  txtYmax.Enabled = True
  lblYinc.Enabled = True
  txtYinc.Enabled = True
  lblYminorTics.Enabled = True
  txtYminorTics.Enabled = True
End Sub

Private Function TrialLibDialog(key)
  frmTrialLib.Show vbModal
  key = frmTrialLib.Tag 'retrieve library key
  Unload frmTrialLib
  If key <> "" Then
    TrialLibDialog = True
  Else
    TrialLibDialog = False
  End If
End Function

Private Function Tier1LibDialog(DataSource As String, DataTitle As String)
  Load frmTier1Lib  'load the form, but don't show it.
  If (UI.PlotVar = PV_VFCUM) Or _
     (UI.PlotVar = PV_VFINC) Then
    frmTier1Lib.AerialOnly = True 'restrict input
  Else
    frmTier1Lib.AerialOnly = False 'allow all models
  End If
  frmTier1Lib.Show vbModal
  'harvest the results
  DataSource = frmTier1Lib.DataSource
  DataTitle = frmTier1Lib.DataTitle
  'done with the form, unload it.
  Unload frmTier1Lib
  'return true only if we found strings
  If DataSource <> "" And DataTitle <> "" Then
    Tier1LibDialog = True
  Else
    Tier1LibDialog = False
  End If
End Function

