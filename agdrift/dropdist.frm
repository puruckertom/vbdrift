VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDropDist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Drop Size Distribution"
   ClientHeight    =   6795
   ClientLeft      =   1320
   ClientTop       =   2205
   ClientWidth     =   7530
   ForeColor       =   &H80000008&
   Icon            =   "DROPDIST.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6795
   ScaleWidth      =   7530
   Begin VB.Frame fraName 
      Caption         =   "Drop Distribution Name"
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtName 
         Height          =   285
         HelpContextID   =   1100
         Left            =   120
         MaxLength       =   40
         TabIndex        =   2
         Top             =   240
         Width           =   7095
      End
   End
   Begin VB.Frame fraDropDist 
      Caption         =   "Drop Distribution"
      Height          =   5415
      Left            =   3840
      TabIndex        =   19
      Top             =   840
      Width           =   3615
      Begin VB.TextBox txtEdit 
         BorderStyle     =   0  'None
         Height          =   255
         HelpContextID   =   1100
         Left            =   720
         TabIndex        =   32
         Text            =   "grid edit text box"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clea&r"
         Height          =   375
         HelpContextID   =   1100
         Left            =   2280
         TabIndex        =   17
         Top             =   4440
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         HelpContextID   =   1100
         Left            =   1320
         TabIndex        =   16
         Top             =   4440
         Width           =   855
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "&Insert"
         Height          =   375
         HelpContextID   =   1100
         Left            =   360
         TabIndex        =   15
         Top             =   4440
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid grdDrop 
         Height          =   3975
         Left            =   0
         TabIndex        =   31
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   7011
         _Version        =   393216
         Cols            =   4
         WordWrap        =   -1  'True
         Appearance      =   0
      End
      Begin VB.Label lblStats 
         AutoSize        =   -1  'True
         Caption         =   "V0.5"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   26
         Top             =   5115
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "µm"
         Height          =   195
         Left            =   1680
         TabIndex        =   25
         Top             =   5040
         Width           =   210
      End
      Begin VB.Label lblRelSpan 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label lblVMD 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label lblStats 
         Alignment       =   2  'Center
         Caption         =   "Relative Span:"
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   22
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label lblStats 
         AutoSize        =   -1  'True
         Caption         =   "D          :"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   5040
         Width           =   615
      End
   End
   Begin VB.Frame fraDistType 
      Caption         =   "Drop Distribution Type"
      Height          =   5415
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   3615
      Begin VB.OptionButton optDistType 
         Caption         =   "FS Rotary &Atomizer Models"
         Height          =   255
         HelpContextID   =   1100
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Frame fraUserLib 
         Caption         =   "User Library"
         Height          =   1215
         Left            =   1560
         TabIndex        =   28
         Top             =   720
         Width           =   1935
         Begin VB.CommandButton cmdUserLibAdd 
            Caption         =   "Add Current"
            Height          =   375
            HelpContextID   =   1100
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdUserLibSelect 
            Caption         =   "Select From/Modify"
            Height          =   375
            HelpContextID   =   1100
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkSwathDispAdjust 
         Caption         =   "Adjust Swath Displacement"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   3000
         Width           =   2295
      End
      Begin VB.OptionButton optDistType 
         Caption         =   "USDA &ARS Nozzle Models"
         Height          =   255
         HelpContextID   =   1100
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   3840
         Width           =   2295
      End
      Begin VB.OptionButton optDistType 
         Caption         =   "&Library (FS)"
         Height          =   255
         HelpContextID   =   1100
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   4920
         Width           =   2295
      End
      Begin VB.CommandButton cmdParametric 
         Caption         =   "Parametric"
         Height          =   375
         HelpContextID   =   1100
         Left            =   360
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   375
         HelpContextID   =   1100
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton cmdInterpolate 
         Caption         =   "Interpolate"
         Height          =   375
         HelpContextID   =   1100
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optDistType 
         Caption         =   "&Library (SDTF)"
         Height          =   255
         HelpContextID   =   1100
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   4560
         Width           =   2295
      End
      Begin VB.OptionButton optDistType 
         Caption         =   "Drop&Kick"
         Height          =   255
         HelpContextID   =   1100
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   2295
      End
      Begin VB.OptionButton optDistType 
         Caption         =   "&Basic"
         Height          =   255
         HelpContextID   =   1100
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   2295
      End
      Begin VB.OptionButton optDistType 
         Caption         =   "&User-defined"
         Height          =   255
         HelpContextID   =   1100
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cboASAEtype 
         Height          =   315
         HelpContextID   =   1100
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2640
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1100
      Left            =   6600
      TabIndex        =   1
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1100
      Left            =   5640
      TabIndex        =   0
      Top             =   6360
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDSDselection 
      Caption         =   "Change the caption of this label to force a DataToForm"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   6240
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmDropDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: dropdist.frm,v 1.12 2011/12/27 17:47:01 tom Exp $
'this flag is used to tell some controls not to
'take action on their new values. This is required
'to differentiate between programatic state changes
'and user actions
Dim PropTakeAction As Integer 'if true, take action

Dim SaveDistType As Integer  'place to save distrib type

'Swath displacement adjustment
'Basic distributions and some DropKick and ARS distributions
'can adjust the swath displacement to specific values
'These vars help determine when and how to do it
Public AdjustSwathDispFlag As Boolean  'true=adjust
Public AdjustSwathDispValue As Single  'Swath Disp value

'grid editing vars
Dim gRow As Integer
Dim gCol As Integer

'Flag indicating which DSD is being edited
'Its value comes from the form's Tag
Dim DSDsel As Integer

Public Sub BasicDistToGrid(iDSD As Integer)
'recover a ASAE distribution from the FORTRAN
'and stuff it in the grid control
  Dim SaveOTA As Integer
  Dim xDSD As DropSizeDistData
  
  'get the distribution and name
  GetBasicDataDSD iDSD, xDSD
  
  'recover the name
  SaveOTA = PropTakeAction
  PropTakeAction = False    'disable Control reactions
  txtName.Text = GetBasicNameDSD(iDSD)
  PropTakeAction = SaveOTA

  'place the new DSD in the grid
  ArrayToGrid xDSD.NumDrop, xDSD.Diam(), xDSD.MassFrac()
End Sub

Private Sub cboASAEtype_Click()
  'only do this if the "Basic (ASAE)" option button is set
  If optDistType(0).Value = True Then
    BasicDistToGrid cboASAEtype.ListIndex
  End If
End Sub

Private Sub ChangeDistType(NewType As Integer)
'Select a new Drop Dist Type and do what is necessary to
'get new data
  Dim i As Integer
  Dim key As String
  Dim nv As Integer
  ReDim dv(MAX_DROPS - 1) As Single
  ReDim xv(MAX_DROPS - 1) As Single

  Me.MousePointer = vbHourglass 'change pointer to hourglass
  
  'by and large, we will not adjust the swath displacement on exit,
  'but there are a few exceptions which will override this below
  AdjustSwathDispFlag = False 'do not adjust swath disp on exit
  AdjustSwathDispValue = 0
  
  Select Case NewType
    Case 0 'Basic
      BasicDistToGrid cboASAEtype.ListIndex 'uses ArrayToGrid to transfer new distribution
      If chkSwathDispAdjust.Value = 1 Then 'checked
        AdjustSwathDispFlag = True 'do adjust swath disp on exit
        GetBasicDataDSDSwathDisp cboASAEtype.ListIndex, AdjustSwathDispValue
      End If
    Case 1 'DropKick
      'The DropKick form allows the user to create a new DSD.
      'If DropKick is successful (.Tag=True), it loads the
      'new DSD into this form with ArrayToGrid and
      'sets the values of AdjustSwathDispFlag and Value
      frmDropKick.Show vbModal
      If frmDropKick.Tag = "False" Then 'Tag holds status info
        'reset original dist type
        temp = PropTakeAction  'save flag state
        PropTakeAction = False 'disable actions on change
        optDistType(SaveDistType).Value = True  'reset option button
        PropTakeAction = temp  'restore flag value
      End If
      Unload frmDropKick
    Case 5 'DropKirk
      'The DropKirk form allows the user to create a new DSD.
      'If DropKirk is successful (.Tag=True), it loads the
      'new DSD into this form with ArrayToGrid and
      'sets the values of AdjustSwathDispFlag and Value
      frmDropKirk.Show vbModal
      If frmDropKirk.Tag = "False" Then 'Tag holds status info
        'reset original dist type
        temp = PropTakeAction  'save flag state
        PropTakeAction = False 'disable actions on change
        optDistType(SaveDistType).Value = True  'reset option button
        PropTakeAction = temp  'restore flag value
      End If
      Unload frmDropKirk
    Case 6 'Rotary Atomizer
      'The Rotary Atomizer form allows the user to create a new DSD.
      'If successful it loads the
      'new DSD into this form with ArrayToGrid and
      'sets the values of AdjustSwathDispFlag and Value
      Dim RotaryAtomizerDialog As frmRotaryAtomizer
      Set RotaryAtomizerDialog = New frmRotaryAtomizer
      With RotaryAtomizerDialog
        .AtomizerIndex = UD.HK(0).RotType
        .SprayMaterialIndex = UD.HK(0).MatType
        .AirSpeed = UD.AC.TypSpeed 'UD.HK(0).Speed
        .BladeAngle = UD.HK(0).BladeAngle
        .BladeRPM = UD.HK(0).BladeRPM
        If UD.SM.FlowRateUnits = 0 Then 'L/ha
          .FlowRate = UD.SM.FlowRate * UD.CTL.SwathWidth * UD.AC.TypSpeed * 0.006 / UD.NZ.NumNoz 'UD.HK(0).Flowrate
        Else 'L/min
          .FlowRate = UD.SM.FlowRate / UD.NZ.NumNoz
        End If
        .DropDistributionType = UD.HK(0).SprayType
        .Show vbModal
        If Not .Cancelled Then
          'Harvest return values here
          UD.HK(0).RotType = .AtomizerIndex
          UD.HK(0).MatType = .SprayMaterialIndex
          UD.HK(0).Speed = .AirSpeed
          UD.HK(0).BladeAngle = .BladeAngle
          UD.HK(0).BladeRPM = .BladeRPM
          UD.HK(0).FlowRate = .FlowRate
          UD.HK(0).SprayType = .DropDistributionType
          If .DropDistributionClassification >= 0 Then 'Classification (Quality)
            BasicDistToGrid .DropDistributionClassification
          Else
            nv = .DropDistributionNumber
            For i = 0 To nv - 1
              dv(i) = .DropDistributionDiameter(i)
              xv(i) = .DropDistributionMassFraction(i)
            Next i
            ArrayToGrid nv, dv(), xv()
          End If
        Else
          'reset original dist type
          temp = PropTakeAction  'save flag state
          PropTakeAction = False 'disable actions on change
          optDistType(SaveDistType).Value = True  'reset option button
          PropTakeAction = temp  'restore flag value
        End If
      End With
      Unload RotaryAtomizerDialog
      Set RotaryAtomizerDialog = Nothing
    Case 2 'User-defined
      'nothing to do here, but do a final cleanup on form exit
    Case 3, 4 'Library
      'The library form allows the user to select a "key" that
      'GetLibDataDSD can use. The Library form does not return
      'the DSD directly.
      Load frmDropLib
      If NewType = 3 Then
        frmDropLib.SelectTable ""
      Else
        frmDropLib.SelectTable "FS"
      End If
      frmDropLib.Show vbModal  'get the mass fractions from lib
      key = frmDropLib.Tag 'retrieve library key
      Unload frmDropLib
      If key = "" Then
        'reset original dist type
        temp = PropTakeAction  'save flag state
        PropTakeAction = False 'disable actions on change
        optDistType(SaveDistType).Value = True  'reset option button
        PropTakeAction = temp  'restore flag value
      Else
        temp = PropTakeAction  'save flag state
        PropTakeAction = False 'disable actions on change
        txtName.Text = key
        GetLibDataDSD key, nv, dv(), xv()
        ArrayToGrid nv, dv(), xv()
        PropTakeAction = temp  'restore flag value
      End If
  End Select
  UpdateTypeControls
  Me.MousePointer = vbDefault 'change pointer back to default
End Sub

Private Sub ClearGrid()
'blank out all the grid rows
  For i = grdDrop.Rows - 1 To 1 Step -1
    grdDrop.Row = i
    grdDrop.Col = 1
    grdDrop.Text = ""
    grdDrop.Col = 2
    grdDrop.Text = ""
  Next
End Sub

Private Sub ClearSelectedCells()
'clear the selected cells in a grid
  With grdDrop
    'Ensure .Row is before .RowSel
    If .RowSel >= .Row Then
      saverow = .Row
    Else
      saverow = .RowSel
      .RowSel = .Row
      .Row = saverow
    End If
    'Ensure .Col is before .ColSel
    If .ColSel >= .Col Then
      savecol = .Col
    Else
      savecol = .ColSel
      .ColSel = .Col
      .Col = savecol
    End If
    'Clear the cell contents in the selected area
    ir1 = .Row: ir2 = .RowSel
    ic1 = .Col: ic2 = .ColSel
    For ic = ic1 To ic2
      .Col = ic
      For ir = ir1 To ir2
        .Row = ir
        .Text = ""
      Next
    Next
    .Row = saverow
    .Col = savecol
  End With
  UpdateCMF 'update the total mass fraction
  UpdateDSDStatsFromGrid  'update the DSD stats
  'set dist type to "user-defined"
  If optDistType(2).Value = False Then optDistType(2).Value = True
End Sub

Private Sub chkSwathDispAdjust_Click()
  If chkSwathDispAdjust.Value = 1 Then 'checked
    AdjustSwathDispFlag = True 'do adjust swath disp on exit
    GetBasicDataDSDSwathDisp cboASAEtype.ListIndex, AdjustSwathDispValue
  Else
    AdjustSwathDispFlag = False
    AdjustSwathDispValue = 0
  End If
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdClear_Click()
  ClearSelectedCells
End Sub

Private Sub cmdDelete_Click()
  DeleteCellRow
End Sub

Private Sub cmdImport_Click()
  ImportDSD
End Sub

Private Sub cmdInsert_Click()
  InsertCellRow
End Sub

Private Sub cmdInterpolate_Click()
'Using the current DSD, Interpolate a new DSD
  InterpolateDSD
End Sub

Private Sub cmdOk_Click()
  FormToData
  Unload Me
End Sub

Private Sub DataToForm()
'transfer user data to form controls for editing
  Dim g As Control
  Set g = grdDrop
  
  Me.Caption = "Drop Size Distribution " + Format$(DSDsel + 1)
  
  'Set the type controls
  cboASAEtype.ListIndex = UD.DSD(DSDsel).BasicType   'ASAE combo box
  
  temp = PropTakeAction                 'save flag value
  PropTakeAction = False                'disable actions for the following
  optDistType(UD.DSD(DSDsel).Type) = True  'dist type radio buttons
  UpdateTypeControls
  'Set the Name
  txtName.Text = Left$(UD.DSD(DSDsel).Name, UD.DSD(DSDsel).LName)

  PropTakeAction = temp                 'restore flag value
  
  'set up the grid
  ArrayToGrid UD.DSD(DSDsel).NumDrop, _
              UD.DSD(DSDsel).Diam(), UD.DSD(DSDsel).MassFrac()
  
  'Copy the user's DropKick and DropKirk Settings to working storage
  DK2 = UD.DK(DSDsel) 'copy from UD to DK2
  BK2 = UD.BK(DSDsel) 'copy from UD to BK1
End Sub

Private Sub DeleteCellRow()
'Delete the selected row(s) from the grid
'and add a new blank one to the end
  
  'Save the beginning of the selection and ensure that
  'the Row property is at the beginning of the selection
  If grdDrop.Row > grdDrop.RowSel Then
    saverow = grdDrop.RowSel
    grdDrop.RowSel = grdDrop.Row
    grdDrop.Row = saverow
  Else
    saverow = grdDrop.Row
  End If
  
  saverows = grdDrop.Rows          'save the original num of rows
  n = grdDrop.RowSel - grdDrop.Row + 1
  R = grdDrop.Row
  For i = 1 To n
    grdDrop.RemoveItem R           'delete the current row
    grdDrop.Rows = saverows        'add blank rows to the end
  Next
  grdDrop.Row = saverow
  grdDrop.RowSel = saverow
  RenumberGrid                'renumber the grid
  UpdateCMF 'update the total mass fraction
  UpdateDSDStatsFromGrid  'update the DSD stats
  'set dist type to "user-defined"
  If optDistType(2).Value = False Then optDistType(2).Value = True
End Sub

Private Sub EditGridCell(KeyAscii As Integer)
'Start editing the current grid cell
  ' Move the text box to the current grid cell:
  PositionTextBox

  ' Save the position of the grids Row and Col for later:
  gRow = grdDrop.Row
  gCol = grdDrop.Col

  ' Make text box same size as current grid cell:
  txtEdit.Width = grdDrop.ColWidth(grdDrop.Col) - 2 * Screen.TwipsPerPixelX
  txtEdit.Height = grdDrop.RowHeight(grdDrop.Row) - 2 * Screen.TwipsPerPixelY

  ' Transfer the grid cell text:
  txtEdit.Text = grdDrop.Text
  txtEdit.SelStart = Len(grdDrop.Text)
  
  ' Show the text box:
  txtEdit.Visible = True
  txtEdit.ZOrder 0
  txtEdit.SetFocus

  'Set the Drop Type ype to User-defined
  If optDistType(2).Value = False Then optDistType(2).Value = True
  
  ' Redirect this KeyPress event to the text box:
  If KeyAscii <> 13 Then 'Enter
     SendKeys Chr$(KeyAscii)
  End If
End Sub

Private Sub FormToData()
'Place the form data in user data storage
  
  Dim nlong As Long
  Dim c As Control
  Dim SwathDisp As Single

  'get the name
  UD.DSD(DSDsel).Name = RTrim$(txtName.Text)
  UD.DSD(DSDsel).LName = Len(RTrim$(txtName.Text))

  'get drop dist ASAE selection, even if the type isn't ASAE
  UD.DSD(DSDsel).BasicType = cboASAEtype.ListIndex  'ASAE selection

  'find the current type selection
  For Each c In optDistType()
    If c.Value = True Then
      UD.DSD(DSDsel).Type = c.Index
      Exit For
    End If
  Next
  
  'get drop distribution from the grid control
  GridToArray UD.DSD(DSDsel).NumDrop, _
              UD.DSD(DSDsel).Diam(), UD.DSD(DSDsel).MassFrac()
  
  'zero out the rest of the arrays
  For i = UD.DSD(DSDsel).NumDrop To MAX_DROPS - 1
    UD.DSD(DSDsel).Diam(i) = 0
    UD.DSD(DSDsel).MassFrac(i) = 0
  Next
  
  'Adjust the swath displacement
  If AdjustSwathDispFlag Then
    UD.CTL.SwathDispType = 0
    UD.CTL.SwathDisp = AdjustSwathDispValue
  End If
  
  'Copy the working DropKick Settings to UD
  UD.DK(DSDsel) = DK2 'copy from DK2 to UD.DK
  UD.BK(DSDsel) = BK2 'copy from DK2 to UD.DK
  
  UpdateDataChangedFlag True 'Data was changed
  UC.Valid = False 'Calcs need to be redone
End Sub

Private Sub cmdParametric_Click()
  Dim VMD As Single
  Dim RelSpan As Single
  Dim SprayType As Long
  Dim SDAdj As Boolean
  Dim SprayQual As Long
  Dim SpectrumSource As Long
  Dim c As Control
  Dim dum As Single
  Dim nd As Long
  Dim xDSD As DropSizeDistData
  
  frmDropParam.Show vbModal
  If Not frmDropParam.Canceled Then
    'gather form data
    VMD = Val(frmDropParam!txtVMD.Text)
    RelSpan = Val(frmDropParam!txtRelSpan.Text)
    For Each c In frmDropParam!optSprayType()
      If c.Value Then
        SprayType = c.Index
        Exit For
      End If
    Next
    SpectrumSource = frmDropParam!chkConvert.Value
    SDAdj = frmDropParam!chkSwathDispAdjust.Value
    
    'agparm computes the Distribution from the VMD and Relative Span
    agparm SprayType, SprayQual, SpectrumSource, VMD, RelSpan, nd, xDSD.Diam(0), xDSD.MassFrac(0)
    xDSD.NumDrop = CInt(nd)
    
    'Get the distribution for Spray Quality
    AdjustSwathDispFlag = False
    AdjustSwathDispValue = 0
    If SprayQual >= 0 Then
      GetBasicDataDSD CInt(SprayQual), xDSD
      'Swath width adjustment
      If SDAdj Then
        AdjustSwathDispFlag = True 'do adjust swath disp on exit
        GetBasicDataDSDSwathDisp CInt(SprayQual), AdjustSwathDispValue
      End If
    End If
    
    'stuff the new data into the grid control
    ArrayToGrid xDSD.NumDrop, xDSD.Diam(), xDSD.MassFrac()
  End If
  Unload frmDropParam
End Sub

Private Sub cmdUserLibAdd_Click()
  Dim DSDName As String
  Dim NumDrop As Integer
  Dim Diam(MAX_DROPS - 1) As Single
  Dim mfrac(MAX_DROPS - 1) As Single
  
  DSDName = txtName.Text
  GridToArray NumDrop, Diam(), mfrac()
  
  If UserLibAddDropsizeRecord(DSDName, NumDrop, Diam(), mfrac()) Then
    If UserLibSelectDropsize(DSDName, NumDrop, Diam(), mfrac()) Then
      txtName.Text = DSDName
      ArrayToGrid NumDrop, Diam(), mfrac()
    End If
  End If
End Sub

Private Sub cmdUserLibSelect_Click()
  Dim DB As Database
  Dim DSDName As String
  Dim NumDrop As Integer
  Dim Diam(MAX_DROPS - 1) As Single
  Dim mfrac(MAX_DROPS - 1) As Single
  
  DSDName = txtName.Text
  If UserLibOpen(DB) Then 'see if it exists
    DB.Close
    If UserLibSelectDropsize(DSDName, NumDrop, Diam(), mfrac()) Then
      txtName.Text = DSDName
      ArrayToGrid NumDrop, Diam(), mfrac()
    End If
  End If
End Sub

Private Sub Form_Load()
'Initialize the controls on this form
  Dim g As Control

  'center the form
  CenterForm Me

  'init Swath Displacement adjustment stuff
  AdjustSwathDispFlag = False 'do not adjust
  AdjustSwathDispValue = 0
  
  'init the ASAE combo box
  cboASAEtype.Clear
  For i = 0 To 17
    cboASAEtype.AddItem GetBasicNameDSD(i)
  Next

  'Init the Library Options
  If UD.Smokey = 0 Then optDistType(4).Enabled = False
  
  'allow option button changes to take action
  '(see declarations section)
  PropTakeAction = True

  'init the grid
  Set g = grdDrop

  'set the number of rows
  g.Rows = MAX_DROPS + 1
  
  'set column headings and alignments
  g.Row = 0
  g.Col = 1
  g.Text = "Average Diameter (µm)"
  g.Col = 2
  g.Text = "Incremental Volume Fraction"
  g.Col = 3
  g.Text = "Cumulative Volume Fraction"
  g.FixedAlignment(0) = flexAlignCenterCenter
  g.ColAlignment(0) = flexAlignCenterCenter
  g.FixedAlignment(1) = flexAlignCenterCenter
  g.ColAlignment(1) = flexAlignCenterCenter
  g.FixedAlignment(2) = flexAlignCenterCenter
  g.ColAlignment(2) = flexAlignCenterCenter
  g.FixedAlignment(3) = flexAlignCenterCenter
  g.ColAlignment(3) = flexAlignCenterCenter
  
  'set Column widths for 3 columns
  g.RowHeight(0) = 650  'set height of first row
  g.ColWidth(0) = 500   'set width of first column
  wid = CSng(g.Width - g.ColWidth(0) - 325) / 3!
  For i = 1 To g.cols - 1
    g.ColWidth(i) = wid
  Next

  'number the rows
  RenumberGrid
End Sub

Private Sub grdDrop_DblClick()
  EditGridCell 0
End Sub

Private Sub grdDrop_KeyDown(KeyCode As Integer, Shift As Integer)
  'PgUp and PgDn mess up the grid control
  If KeyCode = 33 Or KeyCode = 34 Then
    KeyCode = 0
  End If
End Sub

Private Sub grdDrop_KeyPress(KeyAscii As Integer)
  EditGridCell KeyAscii
End Sub

Private Sub ImportDSD()
'Import a new Dropsize Distribution, either incremental
'or cumulative, fractions or percentages, from a two-column
'text file and load it into the grid. This routine discriminates
'between incremental and cumulative by summing all the mass
'fractions. Distributions that total more than 1.1 are considered
'to be cumulative. Cumulative mass fractions must be converted
'to incremental.
'
'The file may start with an arbitrary number of comment lines. These
'lines begin with a pound sign (#) in column one. No comment lines
'may appear after the first data line.
'
'If the input file describes a distribution with the upper diameter
'of each class specified, the first data line of the file must contain
'a single number which depicts the lower limit of the first diameter
'range.

  Dim fn As String
  Dim SavePTA As Integer
  Dim xDSD As DropSizeDistData
  Dim totalMF As Single
  Dim maxMF As Single
  Dim dmin As Single
  Dim buf As String
  Dim UpperDiam As Boolean
  
  'define a collection of separators (space, comma, tab)
  seps = " ," + vbTab
  
  If FileDialog(FD_OPEN, FD_TYPE_TEXT, fn) Then  'get a filename
    'Open the file
    On Error GoTo ErrHandlerIMF
    OpenFileAndSkipComments fn, 1
    
    'Sniff the first data line for the number of values. If there is
    'one number, the file is in upper diameter format. Otherwise,
    'it is an average diameter file.
    Line Input #1, buf
    buf = Trim(buf)
    UpperDiam = True 'default value
    If InStr(buf, " ") Then UpperDiam = False
    If InStr(buf, ",") Then UpperDiam = False
    If InStr(buf, vbTab) Then UpperDiam = False
    
    'Reopen the file to rewind it to the beginning of the data
    Close #1
    OpenFileAndSkipComments fn, 1
    
    xDSD.NumDrop = 0
    totalMF = 0
    maxMF = 0
    
    'For upper diameter import, the first line contains the
    'lower bound for the first range
    If UpperDiam Then Input #1, dmin
    
    'the rest of the file is the same for both types
    While Not EOF(1)
      Input #1, xDSD.Diam(xDSD.NumDrop), xDSD.MassFrac(xDSD.NumDrop)
      totalMF = totalMF + xDSD.MassFrac(xDSD.NumDrop)
      If xDSD.MassFrac(xDSD.NumDrop) > maxMF Then maxMF = xDSD.MassFrac(xDSD.NumDrop)
      xDSD.NumDrop = xDSD.NumDrop + 1
    Wend
    
    Close #1
    
    'If any MassFrac is greater than 1, the file must contain percentages
    'rather than fractions. Convert to fractions
    If maxMF > 1 Then
      For i = 0 To xDSD.NumDrop - 1
        xDSD.MassFrac(i) = xDSD.MassFrac(i) / 100#
      Next
      totalMF = totalMF / 100#
      maxMF = maxMF / 100#
    End If
    
    'If we just read in a cumulative distribution, we
    'must convert it to incremental.
    If totalMF >= 1.1 Then
      For i = xDSD.NumDrop - 1 To 1 Step -1 'backwards, skip first
        xDSD.MassFrac(i) = xDSD.MassFrac(i) - xDSD.MassFrac(i - 1)
      Next
    End If
    
    'If we read an upper bound distributions, we must convert
    'it to average diameters
    If UpperDiam Then
      Call agaver(CLng(xDSD.NumDrop), xDSD.Diam(0), dmin, xDSD.Diam(0))
    End If
    
    'stuff the new data into the grid control
    ArrayToGrid xDSD.NumDrop, xDSD.Diam(), xDSD.MassFrac()
  End If
  Exit Sub

ErrHandlerIMF:
  Close #1
  s = "Error importing file: " + fn + Chr$(13) + Error$(Err)
  t% = vbCritical + vbOKOnly
  MsgBox s, t%
  Exit Sub
End Sub

Private Sub InsertCellRow()
'Insert a blank row of cells in a grid above the current row
  saverow = grdDrop.Row
  grdDrop.Row = grdDrop.Rows - 1   'move to the end of the table
  grdDrop.Col = 1
  s1$ = grdDrop.Text
  grdDrop.Col = 2
  s2$ = grdDrop.Text
  If s1$ = "" And s2$ = "" Then
    grdDrop.RemoveItem grdDrop.Rows - 1  'remove the last row
    grdDrop.AddItem "", saverow          'add a new row
  End If
  grdDrop.Row = saverow
  RenumberGrid
  'set dist type to "user-defined"
  If optDistType(2).Value = False Then optDistType(2).Value = True
End Sub

Private Sub grdDrop_Scroll()
  If txtEdit.Visible Then PositionTextBox
End Sub

Private Sub lblDSDselection_Change()
  DSDsel = Int(lblDSDselection.Caption)
  DataToForm
End Sub

Private Sub optDistType_Click(Index As Integer)
  If PropTakeAction Then ChangeDistType Index
End Sub

Private Sub optDistType_DblClick(Index As Integer)
  If PropTakeAction Then ChangeDistType Index
End Sub

Private Sub PositionTextBox()
'Move the txtEdit Text to cover the current grid cell
  With grdDrop
    If .RowIsVisible(.Row) And .ColIsVisible(.Col) Then
    txtEdit.Left = .Left + .CellLeft
    txtEdit.Top = .Top + .CellTop
    txtEdit.Height = .CellHeight
    txtEdit.Width = .CellWidth
    Else
      .SetFocus
    End If
  End With
End Sub

Private Sub RenumberGrid()
'Redo the row numbering on the grid
  Dim g As Control
  Set g = grdDrop
  saverow = g.Row
  savecol = g.Col

  g.Col = 0
  For i = 1 To g.Rows - 1
    g.Row = i
    g.Text = AGFormat$(i)
  Next

  g.Row = saverow
  g.Col = savecol
End Sub

Public Sub ArrayToGrid(nd As Integer, Diam() As Single, mfrac() As Single)
'Transfer the given Drop Distribution to the grid control
  Dim g As Control
  
  Set g = grdDrop

  'transfer the distribution to the output control
  For i = 0 To nd - 1
    g.Row = i + 1
    g.Col = 1
    g.Text = Format$(Diam(i)) 'AGFormat$(Diam(i))
    g.Col = 2
    g.Text = Format$(mfrac(i)) 'AGFormat$(mfrac(i))
  Next

  'clean out the rest of the grid control
  For i = nd + 1 To g.Rows - 1
    g.Row = i
    g.Col = 1
    g.Text = ""
    g.Col = 2
    g.Text = ""
  Next
  'Fill in the Cumulative column
  UpdateCMF
  'Calculate and display statistics
  UpdateDSDStats nd, Diam(), mfrac()
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'Change the function of some special keys
  Select Case KeyCode
    Case vbKeyDown
      grdDrop.SetFocus ' Set focus back to grid, see txtEdit_LostFocus.
      SendKeys "{DOWN}" 'send a downarrow to the grid
    Case vbKeyUp
      grdDrop.SetFocus ' Set focus back to grid, see txtEdit_LostFocus.
      SendKeys "{UP}"  'send an uparrow to the grid
  End Select
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then  'Enter
    grdDrop.SetFocus ' Set focus back to grid, see txtEdit_LostFocus.
    KeyAscii = 0     ' Ignore this KeyPress.
  End If
End Sub

Private Sub txtEdit_LostFocus()
  Dim tmpRow As Integer
  Dim tmpCol As Integer

  ' Save current settings of Grid Row and col. This is needed only if
  ' the focus is set somewhere else in the Grid.
  tmpRow = grdDrop.Row
  tmpCol = grdDrop.Col

  ' Set Row and Col back to what they were before Text1_LostFocus:
  grdDrop.Row = gRow
  grdDrop.Col = gCol

  grdDrop.Text = txtEdit.Text  ' Transfer text back to grid.
  txtEdit.SelStart = 0       ' Return caret to beginning.
  txtEdit.Visible = False    ' Disable text box.

  ' Return row and Col contents:
  grdDrop.Row = tmpRow
  grdDrop.Col = tmpCol

  'recalc Mass Fraction
  If gCol = 2 Then 'IMF column was changed
    UpdateCMF
  ElseIf gCol = 3 Then 'CMF column was changed
    UpdateIMF
  End If

  UpdateDSDStatsFromGrid  'update the DSD stats
End Sub

Private Sub txtName_Change()
'if this field is changed by the user, flip to user-defined
  If PropTakeAction Then
    'Change if not DropKick or User-defined
    If Not optDistType(1).Value And Not optDistType(2).Value Then optDistType(2).Value = True
  End If
End Sub

Private Sub UpdateCMF()
'redo the Cumulative Mass Fraction column based on the
'Incremental Mass Fraction Column
  Dim saverow As Integer
  Dim savecol As Integer
  Dim tot As Single
  Dim g As Control

  Set g = grdDrop
  'save current position
  saverow = g.Row
  savecol = g.Col
  
  'sum up the mass fractions and place the running total
  'in the CMF column
  tot = 0
  For i = 1 To g.Rows - 1
    g.Row = i
    g.Col = 2
    If g.Text <> "" Then
      tot = tot + Val(g.Text)
      g.Col = 3
      g.Text = AGFormat$(tot)
    Else
      g.Col = 3
      g.Text = ""
    End If
  Next

  'restore the original position
  g.Col = savecol
  g.Row = saverow

End Sub

Private Sub UpdateIMF()
'redo the Incremental Mass Fraction column based on the
'Cumulative Mass Fraction Column
  Dim saverow As Integer
  Dim savecol As Integer
  Dim newMF As Single
  Dim prev As Single
  Dim g As Control

  Set g = grdDrop
  'save current position
  saverow = g.Row
  savecol = g.Col
  
  'calculate the IMF from the CMF
  prev = 0
  For i = 1 To g.Rows - 1
    g.Row = i
    g.Col = 3 'CMF column
    If g.Text <> "" Then
      g.Col = 3 'CMF column
      newMF = Val(g.Text)
      g.Col = 2 'IMF column
      g.Text = AGFormat$(newMF - prev)
      prev = newMF 'save value for next time
    Else
      g.Col = 2 'IMF column
      g.Text = ""
    End If
  Next

  'restore the original position
  g.Col = savecol
  g.Row = saverow

End Sub

Private Sub UpdateTypeControls()
'based on the value of the option buttons, set the other
'dist type controls
  Dim c As Control
  Dim userctls As New Collection
  Dim basicctls As New Collection
  
  'fill collections with associated controls
  userctls.Add cmdInterpolate
  userctls.Add cmdImport
  userctls.Add cmdParametric
  userctls.Add fraUserLib
  userctls.Add cmdUserLibAdd
  userctls.Add cmdUserLibSelect
  
  basicctls.Add cboASAEtype
  basicctls.Add chkSwathDispAdjust

  'find the current type selection
  For Each c In optDistType()
    If c.Value Then
      SaveDistType = c.Index
      Exit For
    End If
  Next
  'set the minor controls
  For Each c In userctls
    c.Enabled = (SaveDistType = 2)
  Next
  For Each c In basicctls
    c.Enabled = (SaveDistType = 0)
  Next

End Sub

Private Sub GridToArray(n As Integer, Diam() As Single, mf() As Single)
  'get drop distribution from the grid control, skipping
  'blank entries and place in the given arrays
  Dim g As Control
  Set g = grdDrop
  saverow = g.Row
  savecol = g.Col
  n = 0
  For i = 1 To g.Rows - 1
    g.Row = i 'set current row
    g.Col = 1 'set current column
    If Trim$(g.Text) <> "" Then
      Diam(n) = Val(g.Text)
      g.Col = 2 'set current column
      mf(n) = Val(g.Text)
      n = n + 1
    End If
  Next
  g.Row = saverow
  g.Col = savecol
End Sub

Private Sub InterpolateDSD()
  'interpolate the current DSD
  Dim nv As Integer
  ReDim dv(100) As Single
  ReDim xv(100) As Single
  Dim f As Form
  Dim itype As Long
  Dim nlong As Long
  Dim ier As Long
  Dim cdat As String * 40
  Dim clen As Long

  'Get the current DSD from the grid
  GridToArray nv, dv(), xv()
  nlong = nv
  'Get options with the interp form
  Set f = frmInterpolateDSD
  f.Show vbModal
  itype = Val(f.Tag) 'get return value from form
  Unload f 'get rid of the form
  If itype >= 0 Then
    Call agfill(itype, nlong, dv(0), xv(0), nlong, dv(0), xv(0), ier, cdat, clen)
    nv = nlong
    Select Case ier
      Case 0: 'success!
        ArrayToGrid nv, dv(), xv()
      Case 4: 'error, don't use output
        MsgBox "Error! " & Left$(cdat, clen), vbCritical + vbOKOnly
      Case 5: 'warning, do use output
        MsgBox "Warning! " & Left$(cdat, clen), vbInformation + vbOKOnly
        ArrayToGrid nv, dv(), xv()
    End Select
  End If
End Sub

Private Sub UpdateDSDStatsFromGrid()
'Retrieve the current grid values and use them
'to update the DSD Stats controls
  Dim nd As Integer
  Dim Diam(MAX_DROPS - 1) As Single
  Dim mfrac(MAX_DROPS - 1) As Single
  
  GridToArray nd, Diam(), mfrac()
  UpdateDSDStats nd, Diam(), mfrac()
End Sub

Private Sub UpdateDSDStats(nd As Integer, Diam() As Single, mfrac() As Single)
'Calculate certain DSD stats and display them
  Dim VMD As Single
  Dim Span As Single
  Dim D10 As Single
  Dim D90 As Single
  Dim F141 As Single
  Dim DP As Single
  
  Call agdsrn(0, CLng(nd), Diam(0), mfrac(0), _
              VMD, Span, D10, D90, F141, DP)
  
  'Update the controls
  If VMD >= 0 Then
    lblVMD.Caption = AGFormat$(VMD)
  Else
    lblVMD.Caption = ""
  End If
  If Span >= 0 Then
    lblRelSpan.Caption = AGFormat$(Span)
  Else
    lblRelSpan.Caption = ""
  End If
End Sub
