VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "AgDRIFT"
   ClientHeight    =   6510
   ClientLeft      =   1350
   ClientTop       =   1905
   ClientWidth     =   9480
   Icon            =   "MAIN.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   9420
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9480
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      HelpContextID   =   1140
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuSepFile0 
         Caption         =   "-"
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuLoadField 
         Caption         =   "&Load Field Trial Data..."
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuSepFile1 
         Caption         =   "-"
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export..."
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuSepFile2 
         Caption         =   "-"
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuPrintPreview 
         Caption         =   "Print Pre&view"
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuPrintSetup 
         Caption         =   "Print Set&up..."
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print..."
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuSepFile3 
         Caption         =   "-"
         HelpContextID   =   1140
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         HelpContextID   =   1140
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      HelpContextID   =   1130
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         HelpContextID   =   1130
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         HelpContextID   =   1130
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         HelpContextID   =   1130
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
         HelpContextID   =   1130
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSepEdit1 
         Caption         =   "-"
         HelpContextID   =   1130
      End
      Begin VB.Menu mnuPrefs 
         Caption         =   "&Preferences..."
         HelpContextID   =   1130
      End
   End
   Begin VB.Menu mnuTier 
      Caption         =   "&Tier"
      HelpContextID   =   1454
      Begin VB.Menu mnuTierModeAg 
         Caption         =   "Tier I Aerial (Agricultural)"
         HelpContextID   =   1454
         Index           =   0
      End
      Begin VB.Menu mnuTierModeAg 
         Caption         =   "Tier I Ground (Agricultural)"
         HelpContextID   =   1454
         Index           =   1
      End
      Begin VB.Menu mnuTierModeAg 
         Caption         =   "Tier I Orchard/Airblast (Agricultural)"
         HelpContextID   =   1454
         Index           =   2
      End
      Begin VB.Menu mnuTierModeAg 
         Caption         =   "Tier II Aerial (Agricultural)"
         HelpContextID   =   1454
         Index           =   3
      End
      Begin VB.Menu mnuTierModeAg 
         Caption         =   "Tier III Aerial (Agricultural)"
         HelpContextID   =   1454
         Index           =   4
      End
      Begin VB.Menu mnuTierSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTierModeFS 
         Caption         =   "Tier II Aerial (Forestry)"
         HelpContextID   =   1454
         Index           =   0
      End
      Begin VB.Menu mnuTierModeFS 
         Caption         =   "Tier III Aerial (Forestry)"
         HelpContextID   =   1454
         Index           =   1
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      HelpContextID   =   1320
      Begin VB.Menu nmuViewNotes 
         Caption         =   "&Notes"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuInputSummary 
         Caption         =   "&Input Summary"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewNumerics 
         Caption         =   "Numerical &Values"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewCalcLog 
         Caption         =   "Calculation &Log"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuSepView1 
         Caption         =   "-"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewDropDist 
         Caption         =   "Drop &Size Distribution"
         HelpContextID   =   1320
         Begin VB.Menu mnuViewINDSD 
            Caption         =   "&Initial"
            HelpContextID   =   1320
            Begin VB.Menu mnuDSDinc 
               Caption         =   "&Incremental"
               HelpContextID   =   1320
               Index           =   0
            End
            Begin VB.Menu mnuDSDcumul 
               Caption         =   "&Cumulative"
               HelpContextID   =   1320
               Index           =   0
            End
         End
         Begin VB.Menu mnuViewDWDSD 
            Caption         =   "&Downwind"
            HelpContextID   =   1320
            Begin VB.Menu mnuDWDSDinc 
               Caption         =   "&Incremental"
               HelpContextID   =   1320
            End
            Begin VB.Menu mnuDWDSDcumul 
               Caption         =   "&Cumulative"
               HelpContextID   =   1320
            End
         End
         Begin VB.Menu mnuViewFXDSD 
            Caption         =   "&Vertical Profile"
            HelpContextID   =   1320
            Begin VB.Menu mnuFXDSDinc 
               Caption         =   "&Incremental"
               HelpContextID   =   1320
            End
            Begin VB.Menu mnuFXDSDcumul 
               Caption         =   "&Cumulative"
               HelpContextID   =   1320
            End
         End
         Begin VB.Menu mnuViewSBDSD 
            Caption         =   "&Spray Block"
            HelpContextID   =   1320
            Begin VB.Menu mnuSBDSDinc 
               Caption         =   "&Incremental"
               HelpContextID   =   1320
            End
            Begin VB.Menu mnuSBDSDcumul 
               Caption         =   "&Cumulative"
               HelpContextID   =   1320
            End
         End
         Begin VB.Menu mnuViewCNDSD 
            Caption         =   "&Canopy"
            HelpContextID   =   1320
            Begin VB.Menu mnuCNDSDinc 
               Caption         =   "&Incremental"
               HelpContextID   =   1320
            End
            Begin VB.Menu mnuCNDSDcumul 
               Caption         =   "&Cumulative"
               HelpContextID   =   1320
            End
         End
      End
      Begin VB.Menu mnuViewSetVel 
         Caption         =   "Settling Velocit&y"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewDepos 
         Caption         =   "&Deposition"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewAvgDepos 
         Caption         =   "&Pond-Integrated Deposition"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewFlux 
         Caption         =   "&Vertical Profile"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewConc 
         Caption         =   "&1 Hour Average Concentration"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewLay 
         Caption         =   "Application &Layout"
      End
      Begin VB.Menu mnuViewCOV 
         Caption         =   "&Coefficient of Variation"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewMeanDep 
         Caption         =   "&Mean Deposition"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewFAloft 
         Caption         =   "&Fraction Aloft"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewSB 
         Caption         =   "Spray &Block"
         HelpContextID   =   1320
         Begin VB.Menu mnuViewSBDep 
            Caption         =   "&Deposition"
            HelpContextID   =   1320
         End
         Begin VB.Menu mnuViewSBArea 
            Caption         =   "&Area Coverage"
            HelpContextID   =   1320
         End
      End
      Begin VB.Menu mnuViewCANDep 
         Caption         =   "C&anopy Deposition"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewTimeACC 
         Caption         =   "&Time Accountancy"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewDistACC 
         Caption         =   "&Distance Accountancy"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewHgtACC 
         Caption         =   "&Height Accountancy"
         HelpContextID   =   1320
      End
      Begin VB.Menu mnuViewTotACC 
         Caption         =   "T&otal Accountancy"
         HelpContextID   =   1320
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      HelpContextID   =   1240
      Begin VB.Menu mnuRunRun 
         Caption         =   "&Run Calculations"
         HelpContextID   =   1240
      End
      Begin VB.Menu mnuRunRevert 
         Caption         =   "Revert to &Last Calculations"
         HelpContextID   =   1240
      End
      Begin VB.Menu mnuRunSep1 
         Caption         =   "-"
         HelpContextID   =   1240
      End
      Begin VB.Menu mnuBatch 
         Caption         =   "&Batch Operations..."
         HelpContextID   =   1240
      End
   End
   Begin VB.Menu mnuToolbox 
      Caption         =   "&Toolbox"
      HelpContextID   =   1305
      Begin VB.Menu mnuTBAquatic 
         Caption         =   "&Aquatic Assessment"
         HelpContextID   =   1305
      End
      Begin VB.Menu mnuTBTerrestrial 
         Caption         =   "&Terrestrial Assessment"
         HelpContextID   =   1305
      End
      Begin VB.Menu mnuTBDropDist 
         Caption         =   "D&rop Distance Calculator"
         HelpContextID   =   1305
      End
      Begin VB.Menu mnuTBSBStats 
         Caption         =   "Spray Block Statistics"
         HelpContextID   =   1305
      End
      Begin VB.Menu mnuTBSprayBlock 
         Caption         =   "Spray &Block Assessment"
         HelpContextID   =   1305
      End
      Begin VB.Menu mnuTBStream 
         Caption         =   "S&tream Assessment"
         HelpContextID   =   1305
      End
      Begin VB.Menu mnuTBMultiApp 
         Caption         =   "&Multiple Application Assessment"
         HelpContextID   =   1305
      End
      Begin VB.Menu mnuTBTrajDetails 
         Caption         =   "&Trajectory Details"
         HelpContextID   =   1305
      End
      Begin VB.Menu mnuTBSprayBlockDetails 
         Caption         =   "Spray &Block Details"
         HelpContextID   =   1305
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   1170
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         HelpContextID   =   1170
      End
      Begin VB.Menu mnuHelpConversion 
         Caption         =   "&Metric-English Units Conversion Table"
         HelpContextID   =   1446
      End
      Begin VB.Menu mnuHelpUsing 
         Caption         =   "&Using Help"
         HelpContextID   =   1170
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
         HelpContextID   =   1170
      End
      Begin VB.Menu mnuHelpLicense 
         Caption         =   "End User &License Agreement"
         HelpContextID   =   1170
      End
      Begin VB.Menu mnuHelpSep2 
         Caption         =   "-"
         HelpContextID   =   1170
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About AgDRIFT"
         HelpContextID   =   1170
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: main.frm,v 1.16 2008/10/22 17:26:06 tom Exp $
Option Explicit

Private Sub ExitApp()
'Exit the whole program, performing final cleanup stuff
  Dim Msg As String
  Dim MBType As Integer
  Dim fn As String
  Dim s As String
  Dim dum As Integer

  'give the user a chance to save plot settings
  If PS.Changed Then
    Msg = "The Plot Options have changed." + Chr$(13)
    Msg = Msg + "Do you want to save the changes?"
    MBType = vbQuestion + vbYesNoCancel
    Select Case MsgBox(Msg, MBType) '6="Yes" 7="No" 2="Cancel"
    Case vbYes    'Yes - go to the save routine
      PlotPrefsWrite PS
    Case vbNo     'No - just continue
    Case vbCancel 'Cancel - exit this routine
      Exit Sub
    End Select
  End If
  
  'give the user a chance to save data if it needs it
  If UI.DataChanged Then
    If Not QuerySaveChanged() Then
      Exit Sub
    End If
  End If
  
  'try to delete the "revert file"
  fn = App.Path & Chr$(92) & App.EXEName & ".rvt"
  On Error Resume Next  'skip it if there is an error
  Kill fn
  On Error GoTo 0 'turn off error handling
  
  'send a quit request to help
  dum = WinHelp(Me.hwnd, s, cdlHelpQuit, 0)
  
  'End closes all forms and does a lot of cleanup, but
  'does not trigger a QueryUnload or an Unload event.
  End
End Sub

Private Sub T1GndOrchKluge()
  'The following code is a kluge for Tier 1 Ground and Orchard.
  'Because these forms do range checking on some of the advanced
  'input values, the controls rely on the LostFocus event to tell
  'them when to do it. Clicking on the menu does not cause the control
  'to lose focus, so we must do it here. Other forms that do range checking
  'rely on an OK button to trigger checking, but these main form
  'children are different. This routine is called from all the top
  'level menus
  If Screen.ActiveForm Is frmTier1Gnd Then
    With frmTier1Gnd
      If .ActiveControl Is .txtSwaths Then .txtSwaths_LostFocus
    End With
  End If
  If Screen.ActiveForm Is frmTier1orc Then
    With frmTier1orc
      If .ActiveControl Is .txtStartSwath Then .txtStartSwath_LostFocus
      If .ActiveControl Is .txtEndSwath Then .txtEndSwath_LostFocus
    End With
  End If
End Sub

Private Sub FixMenuBar()
'Set up the menu bar according to the form and Tier
'
  'defaults for items accessed later
  mnuRun.Visible = True
  'do menu changes that depend on Tier
  Select Case UD.Tier
    Case TIER_1
      mnuRun.Visible = False
    Case TIER_2
    Case TIER_3
  End Select
  'Fix Sub Menus
  FixMenuFile
  FixMenuTier
  FixMenuView  'Note that the Initial DSD submenu is fixed elsewhere
  FixMenuToolBox
End Sub

Private Sub FixMenuFile()
'Set up the file menu
  'start with a known state
  mnuNew.Visible = True
  mnuOpen.Visible = True
  mnuSave.Visible = True
  mnuSaveAs.Visible = True
  mnuSepFile0.Visible = True
  mnuLoadField.Visible = True
  mnuExport.Visible = True
  mnuPrintPreview.Visible = True
  mnuPrintSetup.Visible = True
  mnuPrint.Visible = True
  mnuExit.Visible = True
  Select Case UD.Tier
    Case TIER_1  'Tier 1
      mnuSave.Visible = False
      mnuSaveAs.Visible = False
      mnuSepFile0.Visible = False
      mnuLoadField.Visible = False
    Case TIER_2  'Tier 2
      mnuSepFile0.Visible = False
      mnuLoadField.Visible = False
    Case TIER_3  'Tier 3
      mnuSepFile0.Visible = False
      mnuLoadField.Visible = FieldTrialDataExists And (UD.Smokey = AUD_SDTF)
  End Select
End Sub

Private Sub FixMenuTier()
  'allow selection of Tier 1,2 Aerial only in
  'regulatory version
  mnuTierModeAg(0).Visible = AGDRIFTREGULATORY    'T1A
  mnuTierModeAg(1).Visible = True                 'T1G
  mnuTierModeAg(2).Visible = True                 'T1O
  mnuTierModeAg(3).Visible = AGDRIFTREGULATORY    'T2A
  mnuTierModeAg(4).Visible = True                 'T3A

  mnuTierModeFS(0).Visible = AGDRIFTREGULATORY    'T2A
  mnuTierModeFS(1).Visible = True                 'T3A
End Sub

Private Sub FixMenuView()
'set up the view menu according to the current data state
  'Numerics: not available for Gnd/Orch
  mnuViewNumerics.Visible = UD.ApplMethod <> AM_GROUND And UD.ApplMethod <> AM_ORCHARD
  mnuViewCalcLog.Visible = UD.Tier <> TIER_1
  
  mnuViewDropDist.Visible = FixMenuViewDSD
  mnuViewDepos.Visible = PlotIsAvailable(PV_DEP)
  mnuViewAvgDepos.Visible = PlotIsAvailable(PV_PID)
  mnuViewConc.Visible = PlotIsAvailable(PV_CONC)
  mnuViewLay.Visible = PlotIsAvailable(PV_LAYOUT)
  mnuViewCOV.Visible = PlotIsAvailable(PV_COV)
  mnuViewFlux.Visible = PlotIsAvailable(PV_VERT)
  mnuViewMeanDep.Visible = PlotIsAvailable(PV_MEAN)
  mnuViewFAloft.Visible = PlotIsAvailable(PV_FA)
  mnuViewSB.Visible = FixMenuViewSB
  mnuViewCANDep.Visible = PlotIsAvailable(PV_CANDEP)
  mnuViewTimeACC.Visible = PlotIsAvailable(PV_TA)
  mnuViewDistACC.Visible = PlotIsAvailable(PV_DA)
  mnuViewHgtACC.Visible = PlotIsAvailable(PV_HA)
  mnuViewTotACC.Visible = PlotIsAvailable(PV_TAB)
  mnuViewSetVel.Visible = PlotIsAvailable(PV_SV)
'tbc  mnuViewSetVel.Enabled = False 'tbc
End Sub

Private Function FixMenuViewDSD() As Boolean
'Set up the DSD submenu under View
'If there are any visible items, return True
  FixMenuViewDSD = False
  FixMenuViewDSD = FixMenuViewDSD Or PlotIsAvailable(PV_VFINC)
  FixMenuViewDSD = FixMenuViewDSD Or PlotIsAvailable(PV_VFINC0)
  FixMenuViewDSD = FixMenuViewDSD Or PlotIsAvailable(PV_VFINC1)
  FixMenuViewDSD = FixMenuViewDSD Or PlotIsAvailable(PV_VFINC2)
  FixMenuViewDSD = FixMenuViewDSD Or PlotIsAvailable(PV_DWDSDINC)
  FixMenuViewDSD = FixMenuViewDSD Or PlotIsAvailable(PV_FXDSDINC)
  FixMenuViewDSD = FixMenuViewDSD Or PlotIsAvailable(PV_SBDSDINC)
  FixMenuViewDSD = FixMenuViewDSD Or PlotIsAvailable(PV_CNDSDINC)
  
  If FixMenuViewDSD Then
    mnuViewINDSD.Visible = PlotIsAvailable(PV_VFINC) Or _
                           PlotIsAvailable(PV_VFINC0) Or _
                           PlotIsAvailable(PV_VFINC1) Or _
                           PlotIsAvailable(PV_VFINC2)
    mnuViewDWDSD.Visible = PlotIsAvailable(PV_DWDSDINC)
    mnuViewFXDSD.Visible = PlotIsAvailable(PV_FXDSDINC)
    mnuViewSBDSD.Visible = PlotIsAvailable(PV_SBDSDINC)
    mnuViewCNDSD.Visible = PlotIsAvailable(PV_CNDSDINC)
  End If
  
  'Set up the Initial DSD submenu
  If PlotIsAvailable(PV_VFINC) Then FixMenuViewDSDInitial
End Function

Private Function FixMenuViewDSDInitial() As Boolean
'Fixup the Initial DSD menu to match the available options
'In Tier I and Tier II there can be only one DSD, but in Tier III,
'one, two, or three DSD's are acceptable. For all tiers, if there
'is only one DSD, display the default "Incremental" and "Cumulative"
'menu entries. For Tier III, if there is more than one DSD, display
'"Incremental - All, Incremental - DSD 1, etc.". Same for Cumulative.
'
'The Tag property of each menu control contains the PV_* plot
'variable that is to be plotted. So, for Tier I and II this is
'always PV_VFINC0 or PV_VFCUM0.
'
  Dim c As Control
  Dim nDSD As Integer
  Dim Index As Integer
  Dim i As Integer
  Dim inc_flags(2) As Long
  Dim cum_flags(2) As Long
  
  'unload any residual menus
  For Each c In mnuDSDinc()
    If c.Index > 0 Then Unload c
  Next
  For Each c In mnuDSDcumul()
    If c.Index > 0 Then Unload c
  Next
  
  'Set up the default menu items.
  'for tier 1 and tier 2 this is all that needs to be done
  mnuDSDinc(0).Caption = "&Incremental"
  mnuDSDinc(0).Tag = PV_VFINC0 'View DSD 1 only
  mnuDSDcumul(0).Caption = "&Cumulative"
  mnuDSDcumul(0).Tag = PV_VFCUM0 'View DSD 1 only
  
  'for tier 3 count the dsd's and adjust things if necessary
  If UD.Tier = TIER_3 Then
    'copy the PV_* flags into arrays for easy access
    inc_flags(0) = PV_VFINC0
    inc_flags(1) = PV_VFINC1
    inc_flags(2) = PV_VFINC2
    cum_flags(0) = PV_VFCUM0
    cum_flags(1) = PV_VFCUM1
    cum_flags(2) = PV_VFCUM2
    'How many DSD's are attached to nozzles? (always at least 1)
    nDSD = NumberOfDSDsUsed(UD)
    Select Case nDSD
    Case 1 'One DSD, but which one?
      'the menus are all set up, we just need to rewire the
      'PV_* flag
      For i = 0 To MAX_DSD - 1
        If DSDIsUsed(UD, i) Then
          mnuDSDinc(0).Tag = inc_flags(i) 'menu tag contains PV_*
          mnuDSDcumul(0).Tag = cum_flags(i) 'menu tag contains PV_*
        End If
      Next
    Case Is > 1 'more than one DSD
      mnuDSDinc(0).Caption = "&Incremental - All"
      mnuDSDinc(0).Tag = PV_VFINC  'Multi-curve plot
      mnuDSDcumul(0).Caption = "&Cumulative - All"
      mnuDSDcumul(0).Tag = PV_VFCUM 'Multi-curve plot
      Index = 1
      For i = 0 To MAX_DSD - 1
        If DSDIsUsed(UD, i) Then
          Load mnuDSDinc(Index)
          Set c = mnuDSDinc(Index)
          c.Caption = "Incremental - DSD &" + CStr(i + 1) + " only"
          c.Tag = inc_flags(i) 'menu tag contains PV_*
          c.Visible = True
          
          Load mnuDSDcumul(Index)
          Set c = mnuDSDcumul(Index)
          c.Caption = "Cumulative - DSD &" + CStr(i + 1) + " only"
          c.Tag = cum_flags(i) 'menu tag contains PV_*
          c.Visible = True
          Index = Index + 1
        End If
      Next
    End Select
  End If
End Function

Private Function FixMenuViewSB() As Boolean
'Set up the Spray Block submenu under View
'If there are any visible items, return True
  FixMenuViewSB = False
  FixMenuViewSB = FixMenuViewSB Or PlotIsAvailable(PV_SBDEP)
  FixMenuViewSB = FixMenuViewSB Or PlotIsAvailable(PV_SBCOVER)

  If FixMenuViewSB Then
    mnuViewSBDep.Visible = PlotIsAvailable(PV_SBDEP)
    mnuViewSBArea.Visible = PlotIsAvailable(PV_SBCOVER)
  End If
End Function

Private Sub FixMenuToolBox()
'Set up Toolbox menu
  'Start with known state
  mnuTBAquatic.Visible = True
  mnuTBTerrestrial.Visible = True
  mnuTBDropDist.Visible = True
  mnuTBSBStats.Visible = True
  mnuTBSprayBlock.Visible = True
  mnuTBStream.Visible = True
  mnuTBMultiApp.Visible = True
  mnuTBTrajDetails.Visible = True
  mnuTBSprayBlockDetails.Visible = True
  'turn off drop dist for Tier 1
  If UD.Tier = 1 Then
    mnuTBDropDist.Visible = False
  End If
  If UD.Tier < 3 Then
    mnuTBSBStats.Visible = False
  End If
  If UD.Tier < 3 Or UD.Smokey = 0 Then
    mnuTBTrajDetails.Visible = False
    mnuTBSprayBlockDetails.Visible = False
  End If
End Sub

Private Sub MDIForm_Load()
'initialize this form
  Dim Warn As Integer

  'Size and Center the form on the screen
  Me.Width = 9600 '640 pixels
  Me.Height = 7200 '480 pixels
  CenterForm Me

  'set up for the initial tier
  Warn = False 'Don't warn on tier change
  SelectNewTier UP.InitialTier, UP.InitialAM, UP.InitialAUD, Warn
  UI.DataChanged = False 'SelectNewTier has set this flag
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ExitApp 'if we return, the user cancelled
  Cancel = True
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal
End Sub

Private Sub mnuBatch_Click()
'load the calcuations form in batch mode
  'If Pause is true, the user must press start to continue
  UI.StartCalcsOnLoad = False
  UI.CalcsBatchMode = True 'set form mode flag
  frmCalc.Show vbModal
  Unload frmCalc
  SwitchToForm
End Sub

Private Sub mnuClear_Click()
  'clear the contents of the active control
  If TypeOf Screen.ActiveControl Is TextBox Then
    Screen.ActiveControl.SelText = ""
  End If
End Sub

Private Sub mnuCNDSDcumul_Click()
  ShowPlot PV_CNDSDCUM
End Sub

Private Sub mnuCNDSDinc_Click()
  ShowPlot PV_CNDSDINC
End Sub

Private Sub mnuCopy_Click()
  Clipboard.Clear
  If TypeOf Screen.ActiveControl Is TextBox Then
    Clipboard.SetText Screen.ActiveControl.SelText
  End If
End Sub

Private Sub mnuCut_Click()
  'first do the same as a copy
  mnuCopy_Click
  'now clear the contents of the active control
  mnuClear_Click
End Sub

Private Sub mnuDSDcumul_Click(Index As Integer)
  'The menu's Tag property holds the PV_* flag
  ShowPlot CLng(mnuDSDcumul(Index).Tag)
End Sub

Private Sub mnuDSDinc_Click(Index As Integer)
  'The menu's Tag property holds the PV_* flag
  ShowPlot CLng(mnuDSDinc(Index).Tag)
End Sub

Private Sub mnuDWDSDcumul_Click()
  ShowPlot PV_DWDSDCUM
End Sub

Private Sub mnuDWDSDinc_Click()
  ShowPlot PV_DWDSDINC
End Sub

Private Sub mnuEdit_Click()
'determine which edit controls are available
  
  'special kluge for Tier I ground/orchard
  T1GndOrchKluge
  
  'if there is an error, it will be because there is no
  'Screen.ActiveControl, as for plot forms. In this case
  'no cutting or pasting is possible
  On Error GoTo ErrHandlerMEC

  'start with a known state
  mnuCut.Enabled = True
  mnuCopy.Enabled = True
  mnuPaste.Enabled = True
  mnuClear.Enabled = True
  mnuPrefs.Enabled = True

  'If the clipboard is empty, we can't paste
  If Clipboard.GetText() = "" Then mnuPaste.Enabled = False

  If TypeOf Screen.ActiveControl Is TextBox Then
    'can't paste text if the clipboard doesn't have text
    If Not Clipboard.GetFormat(vbCFText) Then mnuPaste.Enabled = False
  Else
    'unsupported control, no editing
    mnuCut.Enabled = False
    mnuCopy.Enabled = False
    mnuPaste.Enabled = False
    mnuClear.Enabled = False
  End If
  Exit Sub

ErrHandlerMEC:
  'turn off all clipboard edit choices
  mnuCut.Enabled = False
  mnuCopy.Enabled = False
  mnuPaste.Enabled = False
  mnuClear.Enabled = False
  Exit Sub
End Sub

Private Sub mnuExit_Click()
  ExitApp
End Sub

Private Sub mnuExport_Click()
  frmExport.Show vbModal
End Sub

Private Sub mnuFile_Click()
  T1GndOrchKluge
End Sub

Private Sub mnuFXDSDcumul_Click()
  ShowPlot PV_FXDSDCUM
End Sub

Private Sub mnuFXDSDinc_Click()
  ShowPlot PV_FXDSDINC
End Sub

Private Sub mnuHelp_Click()
  'special kluge for Tier I ground/orchard
  T1GndOrchKluge
End Sub

Private Sub mnuHelpContents_Click()
'display the contents section of the  help file
  Dim hf As String
  Dim dum As Integer
  hf = App.HelpFile + Chr$(0)
  WinHelp Me.hwnd, hf, cdlHelpIndex, 0
End Sub

Private Sub mnuHelpConversion_Click()
  Dim hf As String
  hf = App.HelpFile + Chr$(0)
  WinHelp Me.hwnd, hf, cdlHelpContext, ByVal CLng(1446)
End Sub

Private Sub mnuHelpLicense_Click()
  frmLicense.Show vbModal
End Sub

Private Sub mnuHelpUsing_Click()
' Display the Standard Windows Help Topic for "Using Help".
  Dim hf As String
  hf = App.HelpFile + Chr$(0)
  WinHelp Me.hwnd, hf, cdlHelpHelpOnHelp, 0
End Sub

Private Sub mnuInputSummary_Click()
'bring up the input summary screen
  Me.MousePointer = vbHourglass
  frmInputSummary.Show vbModal
  Me.MousePointer = vbDefault
End Sub

Private Sub mnuLoadField_Click()
  Me.MousePointer = vbHourglass
  If LoadFieldTrial() Then SwitchToForm
  Me.MousePointer = vbDefault
End Sub

Private Sub mnuNew_Click()
'Reset data to default state
  If NewFile() Then SwitchToForm
End Sub

Private Sub mnuOpen_Click()
  If OpenFile() Then SwitchToForm
End Sub

Private Sub mnuPaste_Click()
  If TypeOf Screen.ActiveControl Is TextBox Then
    Screen.ActiveControl.SelText = Clipboard.GetText()
  End If
End Sub

Private Sub mnuPrefs_Click()
  Me.MousePointer = vbHourglass
  frmPrefs.Show vbModal
  'The tag on this for indcates whether to redisplay the form
  If frmPrefs.Tag = "True" And UD.Tier > 1 Then SwitchToForm
  Unload frmPrefs
  Me.MousePointer = vbDefault
End Sub

Private Sub mnuPrint_Click()
'print the current UserData
  Dim BeginPage As Integer
  Dim EndPage As Integer
  Dim NumCopies As Integer
  Dim ReportText As String
  Dim i As Integer
  Dim pages As Integer
  Dim Mag As Variant

  If PrinterExists() Then
    Me.MousePointer = vbHourglass
    If PrintDialog(BeginPage, EndPage, NumCopies) Then
      ReportText = GenReportText()
      For i = 1 To NumCopies
        PrintData ReportText, False, pages, Mag
      Next
    End If
    Me.MousePointer = vbDefault
  End If
End Sub

Private Sub mnuPrintPreview_Click()
'Preview formatted User Data
  If PrinterExists() Then
    Me.MousePointer = vbHourglass
    frmPrintPreview.Tag = CStr(GenReportText()) 'pass the text to the form
    frmPrintPreview.Show vbModal
    Me.MousePointer = vbDefault
  End If
End Sub

Private Sub mnuPrintSetup_Click()
  If PrinterExists() Then PrintSetupDialog
End Sub

Private Sub mnuRun_Click()
'Set up run menu
  'special kluge for Tier I ground/orchard
  T1GndOrchKluge
  
  mnuRunRun.Enabled = True
  mnuRunRevert.Enabled = False
  mnuBatch.Enabled = True
  'check for running calcs
  If UI.CalcsInProgress Then
    mnuRunRun.Enabled = False
    mnuBatch.Enabled = False
  End If
  'revert
  If UI.RevertCalcsAvailable Then
    mnuRunRevert.Enabled = True
  End If
End Sub

Private Sub mnuRunRevert_Click()
  Dim fn As String
  Dim Msg As String
  Dim MBType As Integer
  Dim dum As Integer
  Msg = "Replace current data and calculation"
  Msg = Msg & " results with those of the last run?"
  MBType = vbQuestion + vbOKCancel
  If MsgBox(Msg, MBType) <> vbOK Then Exit Sub
  
  fn = App.Path & Chr$(92) & App.EXEName & ".rvt"
  dum = UserDataRead(fn, UD, UC, False)
  SwitchToForm 'refresh data form
End Sub

Private Sub mnuRunRun_Click()
  Dim Msg As String
  Dim MBType As Integer
  Dim dum As Integer
  If UC.Valid Then
    Msg = "The calculations have already been performed "
    Msg = Msg + "for these data. Do you want to do them again?"
    MBType = vbQuestion + vbOKCancel
    If MsgBox(Msg, MBType) <> vbOK Then Exit Sub
  End If
  dum = PerformCalcs(UP.PauseBeforeCalc)
End Sub

Private Sub mnuSave_Click()
  If SaveFile() Then ActiveForm.Caption = FormCaption
End Sub

Private Sub mnuSaveAs_Click()
  If SaveAsFile() Then ActiveForm.Caption = FormCaption
End Sub

Private Sub mnuSBDSDcumul_Click()
  ShowPlot PV_SBDSDCUM
End Sub

Private Sub mnuSBDSDinc_Click()
  ShowPlot PV_SBDSDINC
End Sub

Private Sub mnuTBAquatic_Click()
  TBAquatic
End Sub

Private Sub mnuTBTerrestrial_Click()
  TBTerrestrial
End Sub

Private Sub mnuTBDropDist_Click()
  TBDropDist
End Sub

Private Sub mnuTBSBStats_Click()
  TBSBStats
End Sub

Private Sub mnuTBMultiApp_Click()
  TBMultiApp
End Sub

Private Sub mnuTBSprayBlock_Click()
  TBSprayBlock
End Sub

Private Sub mnuTBSprayBlockDetails_Click()
  TBSprayBlockDetails
End Sub

Private Sub mnuTBStream_Click()
  TBStream
End Sub

Private Sub mnuTBTrajDetails_Click()
  TBTrajDetails
End Sub

Private Sub mnuTier_Click()
  Dim c As Control
  'special kluge for Tier I ground/orchard
  T1GndOrchKluge
  With UD
    'The following is a kluge to make checkmarks work under NT/2K/XP. If you
    'don't do this, the *old* checkmark state appears when the menu pulls down
    'the *first* time. If you slide the mouse off the menu so that it goes
    'away, then cause it to reappear, all is well. Menus worked fine without
    'this "fix" under W9x.
    For Each c In mnuTierModeAg(): c.Enabled = False: Next
    For Each c In mnuTierModeFS(): c.Enabled = False: Next
    'check the appropriate menu item
    mnuTierModeAg(0).Checked = .Tier = TIER_1 And .ApplMethod = AM_AERIAL And .Smokey = AUD_SDTF
    mnuTierModeAg(1).Checked = .Tier = TIER_1 And .ApplMethod = AM_GROUND And .Smokey = AUD_SDTF
    mnuTierModeAg(2).Checked = .Tier = TIER_1 And .ApplMethod = AM_ORCHARD And .Smokey = AUD_SDTF
    mnuTierModeAg(3).Checked = .Tier = TIER_2 And .ApplMethod = AM_AERIAL And .Smokey = AUD_SDTF
    mnuTierModeAg(4).Checked = .Tier = TIER_3 And .ApplMethod = AM_AERIAL And .Smokey = AUD_SDTF
    'check the appropriate menu item
    mnuTierModeFS(0).Checked = .Tier = TIER_2 And .ApplMethod = AM_AERIAL And .Smokey = AUD_FS
    mnuTierModeFS(1).Checked = .Tier = TIER_3 And .ApplMethod = AM_AERIAL And .Smokey = AUD_FS
    'This is the other half of the above kluge
    For Each c In mnuTierModeAg(): c.Enabled = True: Next
    For Each c In mnuTierModeFS(): c.Enabled = True: Next
  End With
End Sub

Private Sub mnuTierModeAg_Click(Index As Integer)
  Dim NewTier As Integer
  Dim NewApplMethod As Integer
  Dim NewAudience As Integer

  NewAudience = AUD_SDTF
  
  Select Case Index
  Case 0 'Tier 1 Aerial
    NewTier = TIER_1
    NewApplMethod = AM_AERIAL
  Case 1 'Tier 1 Ground
    NewTier = TIER_1
    NewApplMethod = AM_GROUND
  Case 2 'Tier 1 Orchard
    NewTier = TIER_1
    NewApplMethod = AM_ORCHARD
  Case 3 'Tier 2 Aerial
    NewTier = TIER_2
    NewApplMethod = AM_AERIAL
  Case 4 'Tier 3 Aerial
    NewTier = TIER_3
    NewApplMethod = AM_AERIAL
  End Select
  If NewTier <> UD.Tier Or _
     NewApplMethod <> UD.ApplMethod Or _
     NewAudience <> UD.Smokey Then
    SelectNewTier NewTier, NewApplMethod, NewAudience, UP.WarnOnTierChange
  End If
End Sub

Private Sub mnuTierModeFS_Click(Index As Integer)
  Dim NewTier As Integer
  Dim NewApplMethod As Integer
  Dim NewAudience As Integer

  NewAudience = AUD_FS
  
  Select Case Index
  Case 0 'Tier 2 Aerial/FS
    NewTier = TIER_2
    NewApplMethod = AM_AERIAL
  Case 1 'Tier 3 Aerial/FS
    NewTier = TIER_3
    NewApplMethod = AM_AERIAL
  End Select
  If NewTier <> UD.Tier Or _
     NewApplMethod <> UD.ApplMethod Or _
     NewAudience <> UD.Smokey Then
    SelectNewTier NewTier, NewApplMethod, NewAudience, UP.WarnOnTierChange
  End If
End Sub

Private Sub mnuToolbox_Click()
  'special kluge for Tier I ground/orchard
  T1GndOrchKluge
End Sub

Private Sub mnuView_Click()
'make adjustments not based only on Tier
'This code is here and not in FixMenuView because it needs to run
'each time the menu is pulled down.
'(FixMenuView is called from FixMenuBar, which is called when the
' tier changes)
  
  'special kluge for Tier I ground/orchard
  T1GndOrchKluge
  
  '*** Swath Displacement Adjustments ***
  'start with a known configuration
  mnuViewConc.Enabled = True
  mnuViewFlux.Enabled = True
  'no vert depos or conc if Swath Displacement Type is "frac app rate"
  If UD.CTL.SwathDispType = 1 Then
    mnuViewFlux.Enabled = False
    mnuViewConc.Enabled = False
  End If
  
  '*** DSD Adjustments ***
  FixMenuViewDSDInitial
End Sub

Private Sub mnuViewAvgDepos_Click()
  ShowPlot PV_PID
End Sub

Private Sub mnuViewCalcLog_Click()
'bring up the Numerics screen
  Me.MousePointer = vbHourglass
  frmCalcLog.Show vbModal
  Me.MousePointer = vbDefault
End Sub

Private Sub mnuViewCANDep_Click()
  ShowPlot PV_CANDEP
End Sub

Private Sub mnuViewConc_Click()
  ShowPlot PV_CONC
End Sub

Private Sub mnuViewCOV_Click()
  ShowPlot PV_COV
End Sub

Private Sub mnuViewDepos_Click()
  ShowPlot PV_DEP
End Sub

Private Sub mnuViewDistACC_Click()
  ShowPlot PV_DA
End Sub

Private Sub mnuViewFAloft_Click()
  ShowPlot PV_FA
End Sub

Private Sub mnuViewFlux_Click()
  ShowPlot PV_VERT
End Sub

Private Sub mnuViewHgtACC_Click()
  ShowPlot PV_HA
End Sub

Private Sub mnuViewLay_Click()
  ShowPlot PV_LAYOUT
End Sub

Private Sub mnuViewMeanDep_Click()
  ShowPlot PV_MEAN
End Sub

Private Sub mnuViewNumerics_Click()
'bring up the Numerics screen
  Me.MousePointer = vbHourglass
  frmNumerics.Show vbModal
  Me.MousePointer = vbDefault
End Sub

Private Sub mnuViewSBArea_Click()
  ShowPlot PV_SBCOVER
End Sub

Private Sub mnuViewSBDep_Click()
  ShowPlot PV_SBDEP
End Sub

Private Sub mnuViewSetVel_Click()
  ShowPlot PV_SV
End Sub

Private Sub mnuViewTimeACC_Click()
  ShowPlot PV_TA
End Sub

Private Sub mnuViewTotACC_Click()
  ShowPlot PV_TAB
End Sub

Private Sub nmuViewNotes_Click()
'view or edit the notes
  Me.MousePointer = vbHourglass
  frmNotes.Show vbModal
  Me.MousePointer = vbDefault
End Sub

Private Sub PrintSetupDialog()
'bring up the print setup dialog box
  Dim c As Control
  Set c = CMDialog1       'Point to the dialog control
  c.CancelError = False   'Do not gen an Error on Cancel
  c.Flags = cdlPDPrintSetup 'Set dialog box flags
  c.ShowPrinter           'Display the dialog box
End Sub

Private Sub SelectNewTier(NewTier As Integer, _
                          NewAM As Integer, _
                          NewFS As Integer, _
                          Warn As Integer)
'Select a new tier
'
'NewTier I The new Tier level; 1, 2, or 3
'NewAM   I The new Application Method; 0, 1, or 2
'NewFS   I The new Audience (Smokey=0 or 1)
'Warn    I   if true display warning messages about changing tiers
'            if false, skip the messages
'
  Dim CurrentTier As Integer
  Dim CurrentFS As Integer
  Dim Msg As String
  Dim MBType As Integer
  Dim xUD As UserData
  Dim i As Integer

  'Save the current tier and audience
  CurrentTier = UD.Tier
  CurrentFS = UD.Smokey
  
  'if the tier is changing, warn the user and allow cancellation
  If Warn Then
    If NewTier > CurrentTier Then
      Msg = "You are about to increase the tier level." + Chr$(13)
      Msg = Msg + "New input parameters and plot options" + Chr$(13)
      Msg = Msg + "will be available." + Chr$(13)
      Msg = Msg + "Do you want to continue?"
    ElseIf NewTier < CurrentTier Then
      Msg = "You are about to decrease the tier level." + Chr$(13)
      Msg = Msg + "Some input parameters and plot options" + Chr$(13)
      Msg = Msg + "will be removed and reset to default" + Chr$(13)
      Msg = Msg + "values." + Chr$(13)
      Msg = Msg + "Do you want to continue?"
    Else
      Msg = "You have selected a new Tier that is " + Chr$(13)
      Msg = Msg + "the same as the current Tier." + Chr$(13)
      Msg = Msg + "Do you want to continue?"
    End If
    MBType = vbQuestion + vbOKCancel
    If MsgBox(Msg, MBType) = vbCancel Then Exit Sub
  End If
  
  'change the tier and application method
  UD.Tier = NewTier
  UD.ApplMethod = NewAM
  UD.Smokey = NewFS
  
  'manipulate data according to the new tier
  Select Case NewTier
  Case TIER_1  'Changing to Tier I
    'remove any existing file name
    UI.FileName = ""
    'For tier 1, save current data that is visible to
    'the user and default all the rest
    'save data that shouldn't be defaulted
    xUD.Tier = UD.Tier
    xUD.ApplMethod = UD.ApplMethod
    xUD.Smokey = UD.Smokey
    xUD.Title = UD.Title
    xUD.DSD(0).BasicType = UD.DSD(0).BasicType
    xUD.GA.BasicType = UD.GA.BasicType
    xUD.GA.NumSwaths = UD.GA.NumSwaths
    xUD.OA.BasicType = UD.OA.BasicType
    xUD.OA.BegTrow = UD.OA.BegTrow
    xUD.OA.EndTrow = UD.OA.EndTrow
    'set data to defaults
    UserDataDefault UD
    ClearUserCalc UC
    'restore saved values
    UD.Tier = xUD.Tier
    UD.ApplMethod = xUD.ApplMethod
    UD.Smokey = xUD.Smokey
    UD.Title = xUD.Title
    'DSD.BasicType can be 0-17 for Tier 2,3 Aerial but is
    'more restricted for Tier 1 SDTF/FS
    'Copy the saved value only if it is legal, otherwise
    'allow the default value to take over.
    Select Case UD.Smokey 'treat two audiences differently
    Case AUD_SDTF
      Select Case xUD.DSD(0).BasicType
      Case 2, 4, 6, 8 'these values are okay, copy saved value
        UD.DSD(0).BasicType = xUD.DSD(0).BasicType
      End Select
    Case AUD_FS
      Select Case xUD.DSD(0).BasicType
      Case 0, 2, 4, 6, 8, 10 'these values are okay, copy saved value
        UD.DSD(0).BasicType = xUD.DSD(0).BasicType
      End Select
    End Select
    UD.GA.BasicType = xUD.GA.BasicType
    UD.GA.NumSwaths = xUD.GA.NumSwaths
    UD.OA.BasicType = xUD.OA.BasicType
    UD.OA.BegTrow = xUD.OA.BegTrow
    UD.OA.EndTrow = xUD.OA.EndTrow
    'recover the drop distribution
    LoadTier1Data UD, UC
  Case TIER_2  'Changing to Tier II
    'Most defaults are okay for Tier II or III, but if
    'any need modification, do it here.
    '
    'Force Active Frac = Nonvol Frac for Tier II
    If UD.SM.ACFrac <> UD.SM.NVFrac Then
      UD.SM.ACFrac = UD.SM.NVFrac 'Active=nonvol
      If CurrentTier > 0 Then 'set flags only if not starting up
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone (for other Tiers)
      End If
    End If
    'Other modifications depend on where you're coming from...
    Select Case CurrentTier
    Case 1 'from Tier I
      UD.ApplMethod = 0 'reset to "aerial"
      UpdateDataChangedFlag True   'data has changed
      UC.Valid = False 'Calcs need to be redone (for other Tiers)
    Case 3 'from Tier III
      'Generate a set of default data for comparison
      UserDataDefault xUD
      'Calcs must be always be redone
      UC.Valid = False 'Calcs need to be redone
      'Check DropKick
      For i = 0 To 2
        If UD.DK(i).SprayType = 1 Then
          UD.DK(i).SprayType = 0 'No DSD output in Tier 2
        End If
      Next
      'Check Spray Material
      If UD.SM.Type <> 0 Then 'must be basic
        UD.SM.Type = 0 'switch to Basic
        GetBasicDataSM UD.SM.BasicType, UD.SM 'recover Basic data
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      If UD.SM.ACFrac <> UD.SM.NVFrac Then 'Force Active Frac = Nonvol Frac for Tier II
        UD.SM.ACFrac = UD.SM.NVFrac 'Active=nonvol
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      'Check Aircraft
      If UD.AC.Type <> 0 Then 'must be basic
        UD.AC.Type = 0 'switch to Basic
        GetBasicDataAC UD.AC.BasicType, UD.AC 'recover Basic data
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      If UD.AC.PropEff <> xUD.AC.PropEff Then
        UD.AC.PropEff = xUD.AC.PropEff 'set to default
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      If UD.AC.DragCoeff <> xUD.AC.DragCoeff Then
        UD.AC.DragCoeff = xUD.AC.DragCoeff 'set to default
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      'Check Nozzles
      For i = 0 To UD.NZ.NumNoz - 1 'Nozzles must be DSD 0
        If UD.NZ.NozType(i) <> 0 Then
          UD.NZ.NozType(i) = 0 'Select DSD 0
          UpdateDataChangedFlag True   'data has changed
          UC.Valid = False 'Calcs need to be redone
        End If
      Next
      If UD.NZ.Type <> 0 Then 'distribution must be basic
        UD.NZ.Type = 0 'switch to Basic distribution
        GetBasicDataNZ UD.AC.BasicType, UD.NZ 'recover Basic data (match AC)
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      'Check Meteorology
      If UD.MET.SurfRough <> xUD.MET.SurfRough Then
        UD.MET.SurfRough = xUD.MET.SurfRough 'set to default
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      If UD.MET.WD <> xUD.MET.WD Then
        UD.MET.WD = xUD.MET.WD 'set to default
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      If UD.MET.WindHeight <> xUD.MET.WindHeight Then
        UD.MET.WindHeight = xUD.MET.WindHeight 'set to default
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      If UD.MET.VortexDecay <> xUD.MET.VortexDecay Then
        UD.MET.VortexDecay = xUD.MET.VortexDecay 'set to default
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      If UD.MET.Pressure <> xUD.MET.Pressure Then
        UD.MET.Pressure = xUD.MET.Pressure 'set to default
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      'Canopy
      If UD.CAN.Height > 0 And UD.CAN.Type <> 3 Then 'non-basic canopy
        UD.CAN.Type = 3 'switch to basic
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      'Control
      If UD.CTL.SwathWidthType > 1 Then
        UD.CTL.SwathWidthType = 1    '1.2 * Wingspan
        UD.CTL.SwathWidth = xUD.CTL.SwathWidth
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
      If UD.CTL.MaxComputeTime <> xUD.CTL.MaxComputeTime Then
        UD.CTL.MaxComputeTime = xUD.CTL.MaxComputeTime 'set to default
        UpdateDataChangedFlag True   'data has changed
        UC.Valid = False 'Calcs need to be redone
      End If
    End Select
  Case TIER_3  'Changing to Tier III
    Select Case CurrentTier
    Case 1 'from Tier I
      UD.ApplMethod = 0 'reset to "aerial"
      UpdateDataChangedFlag True   'data has changed
      UC.Valid = False 'Calcs need to be redone
    Case 2 'from Tier II
      UC.Valid = False 'Force a recalc anyway
    End Select
  End Select

  SwitchToForm    'load the new input form
End Sub

Private Sub SwitchToForm()
'load a new child form, unloading any previous form
'isel - the new form type
'       0 = input form
'
  Dim f As Form

  On Error Resume Next 'In case there is no active form
  Unload ActiveForm    'Unload current form
  On Error GoTo 0      'turn off error trapping
  FixMenuBar 'adjust the main menu bar
  Select Case UD.Tier
  Case TIER_1
    Select Case UD.ApplMethod
    Case AM_AERIAL 'aerial
      Set f = frmTier1air
    Case AM_GROUND 'ground
      Set f = frmTier1Gnd
    Case AM_ORCHARD 'orchard/airblast
      Set f = frmTier1orc
    End Select
  Case TIER_2
    Select Case UD.Smokey
    Case AUD_SDTF 'SDTF
      Set f = frmTier2air
    Case AUD_FS 'FS
      Set f = frmTier2afs
    End Select
  Case TIER_3
    Select Case UD.Smokey
    Case AUD_SDTF 'SDTF
      Set f = frmTier3air
    Case AUD_FS 'FS
      Set f = frmTier3afs
    End Select
  End Select
  'To compensate for an apparent bug in compiled Visual
  'Basic code, size the child form to be smaller than the
  'MDI form, then maximize it. This was required to get
  'the child caption integrated with the MDI caption
  'properly. It seems that when a child form is larger
  'than the MDI parent, it's caption may not be integrated
  'into the parent caption as documented.
  Load f
  f.Width = Me.Width / 2
  f.Height = Me.Height / 2
  f.WindowState = 2         'maximized
  f.Show
End Sub

Private Sub TBAquatic()
'bring up the Aquatic calculator
  If UD.Tier > 1 Then
    'for tier 2, see about calculations
    If Not UC.Valid Then
      If Not QueryPerformCalcs() Then Exit Sub
    End If
  End If
  'Load the form, then its data, then show it.
  Load frmTBAquatic
  frmTBAquatic.LoadDeposition UC.NumDep, UC.DepDist(), UC.DepVal(), _
                              UC.NumPID, UC.PIDDist(), UC.PIDVal()
  frmTBAquatic.Show vbModal
End Sub

Private Sub TBTerrestrial()
'bring up the Terrestrial Assessment Toolbox
  If UD.Tier > 1 Then
    'for tier 2, see about calculations
    If Not UC.Valid Then
      If Not QueryPerformCalcs() Then Exit Sub
    End If
  End If
  'Load the form, then its data, then show it.
  Load frmTBTerrestrial
  frmTBTerrestrial.LoadDeposition UC.NumDep, UC.DepDist(), UC.DepVal(), _
                                  UC.NumPID, UC.PIDDist(), UC.PIDVal()
  frmTBTerrestrial.Show vbModal
End Sub

Private Sub TBDropDist()
'bring up the Drop Distance Toolbox item
  If UD.Tier > 1 Then
    'for tier 2, see about calculations
    If Not UC.Valid Then
      If Not QueryPerformCalcs() Then Exit Sub
    End If
  End If
  frmTBDropDist.Show vbModal
End Sub

Private Sub TBSBStats()
'bring up the Spray Block Statistics Toolbox item
  If UD.Tier > 1 Then
    'for tier 2 or 3, see about calculations
    'Since Field Trial calcs don't include COV,
    'test for positive number of data points as well
    If (Not UC.Valid) Or (UC.NumCOV = 0) Then
      If Not QueryPerformCalcs() Then Exit Sub
    End If
  End If
  frmTBSBStats.Show vbModal
End Sub

Private Sub TBSprayBlock()
'bring up the Aquatic calculator
  If UD.Tier > 1 Then
    'for tier 2, see about calculations
    'Since Field Trial calcs don't include Single Swath depos,
    'test for positive number of data points as well
    If (Not UC.Valid) Or (UC.NumSgl = 0) Then
      If Not QueryPerformCalcs() Then Exit Sub
    End If
  End If
  frmTBSprayBlock.Show vbModal
End Sub

Private Sub TBSprayBlockDetails()
'bring up the Stream calculator
  If UD.Tier > 1 Then
    'for tier 2, see about calculations
    If (Not UC.Valid Or UC.NumDep = 0) Then
      If Not QueryPerformCalcs() Then Exit Sub
    End If
  End If
  frmTBSprayBlockDetails.Show vbModal
End Sub

Private Sub TBStream()
'bring up the Stream calculator
  If UD.Tier > 1 Then
    'for tier 2, see about calculations
    If (Not UC.Valid Or UC.NumDep = 0) Then
      If Not QueryPerformCalcs() Then Exit Sub
    End If
  End If
  frmTBStream.Show vbModal
End Sub

Private Sub TBTrajDetails()
'bring up the Stream calculator
  If UD.Tier > 1 Then
    'for tier 2, see about calculations
    If (Not UC.Valid Or UC.NumDep = 0) Then
      If Not QueryPerformCalcs() Then Exit Sub
    End If
  End If
  frmTBTrajDetails.Show vbModal
End Sub

Private Sub TBMultiApp()
'bring up the Multple Application calculator
  Dim DB As Database
  'make sure the library is there
  If LibOpenMAADB(DB) Then
    DB.Close 'Close the library
    frmTBMultiApp.Show 'can't show form as modal
  End If
End Sub

