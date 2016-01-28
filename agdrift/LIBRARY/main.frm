VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "AgDRIFT Database Utility"
   ClientHeight    =   7095
   ClientLeft      =   1185
   ClientTop       =   2100
   ClientWidth     =   10155
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "MAIN.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7095
   ScaleWidth      =   10155
   Begin VB.Frame fraTable 
      Caption         =   "Info"
      Height          =   615
      Index           =   10
      Left            =   3480
      TabIndex        =   47
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton cmdDeleteInfo 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdInfoRecords 
         Caption         =   "Create"
         Height          =   255
         Left            =   1080
         TabIndex        =   48
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraTable 
      Caption         =   "MAA - Wind Rose (src/Windrose/*)"
      Height          =   615
      Index           =   8
      Left            =   6840
      TabIndex        =   38
      Top             =   3600
      Width           =   3255
      Begin VB.CommandButton cmdWindroseRecords 
         Caption         =   "Create"
         Height          =   255
         Left            =   1080
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteWindrose 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraTable 
      Caption         =   "Dropsize (atomize.out,dropsize.out)"
      Height          =   615
      Index           =   0
      Left            =   3480
      TabIndex        =   12
      Top             =   600
      Width           =   3255
      Begin VB.CommandButton cmdDropsize 
         Caption         =   "Create"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteDropsize 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Library Output Control"
      Height          =   5295
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3375
      Begin VB.Frame Frame12 
         Caption         =   "Output Directory"
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3135
         Begin VB.TextBox txtDBpath 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Text            =   "database path"
            Top             =   240
            Width           =   2775
         End
         Begin VB.OptionButton optDB 
            Caption         =   "Test DB Directory"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   2655
         End
         Begin VB.OptionButton optDB 
            Caption         =   "Actual DB Directory"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   2655
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Library Type"
         Height          =   1455
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   3135
         Begin VB.OptionButton optDBOutputType 
            Caption         =   "MAA Library"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   44
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox txtDBfile 
            Height          =   285
            Left            =   120
            TabIndex        =   37
            Text            =   "database file"
            Top             =   240
            Width           =   2775
         End
         Begin VB.OptionButton optDBOutputType 
            Caption         =   "Standard Library"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton optDBOutputType 
            Caption         =   "SDTF Proprietary Library"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   840
            Width           =   2415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Library Version"
         Height          =   615
         Left            =   120
         TabIndex        =   45
         Top             =   2880
         Width           =   3135
         Begin VB.Label lblLibVersion 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Operations"
         Height          =   1695
         Left            =   120
         TabIndex        =   33
         Top             =   3480
         Width           =   3135
         Begin VB.CommandButton cmdRegenerateStdProp 
            Caption         =   "Completely Rebuild Std, Prop"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   960
            Width           =   2895
         End
         Begin VB.CommandButton cmdRegenerateAll 
            Caption         =   "Completely Rebuild All Libraries"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1320
            Width           =   2895
         End
         Begin VB.CommandButton cmdDeleteDB 
            Caption         =   "Delete Library "
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdCreateDB 
            Caption         =   "Create Library"
            Height          =   255
            Left            =   1560
            TabIndex        =   35
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdRegenerate 
            Caption         =   "Completely Rebuild Library"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   2895
         End
      End
   End
   Begin VB.Frame fraTable 
      Caption         =   "(SDTF) Field Trial (fieldrun.out)"
      Height          =   615
      Index           =   7
      Left            =   6840
      TabIndex        =   1
      Top             =   2760
      Width           =   3255
      Begin VB.CommandButton cmdDeleteTrial 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdTrialRecords 
         Caption         =   "Create"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ListBox lstLog 
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   10215
   End
   Begin VB.Frame fraTable 
      Caption         =   "Basic (bcpcdrp/dep,air/nozbasic)"
      Height          =   615
      Index           =   6
      Left            =   6840
      TabIndex        =   30
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton cmdDeleteBasic 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBasicRecords 
         Caption         =   "Create"
         Height          =   255
         Left            =   1080
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraTable 
      Caption         =   "Components (subst.out)"
      Height          =   615
      Index           =   1
      Left            =   3480
      TabIndex        =   15
      Top             =   1200
      Width           =   3255
      Begin VB.CommandButton cmdComponent 
         Caption         =   "Create"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteComponents 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraTable 
      Caption         =   "Canopy (canopy.out)"
      Height          =   615
      Index           =   9
      Left            =   6840
      TabIndex        =   41
      Top             =   600
      Width           =   3255
      Begin VB.CommandButton cmdDeleteCanopy 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCanopyRecords 
         Caption         =   "Create"
         Height          =   255
         Left            =   1080
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraTable 
      Caption         =   "Evaporation (evap.out)"
      Height          =   615
      Index           =   2
      Left            =   3480
      TabIndex        =   18
      Top             =   1800
      Width           =   3255
      Begin VB.CommandButton cmdDeleteEvaporation 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdEvaporationRecords 
         Caption         =   "Create"
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraTable 
      Caption         =   "DK Nozzle (nozzle.out,nozzletm.out)"
      Height          =   615
      Index           =   3
      Left            =   3480
      TabIndex        =   21
      Top             =   2400
      Width           =   3255
      Begin VB.CommandButton cmdNozzleRecords 
         Caption         =   "Create"
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteNozzle 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraTable 
      Caption         =   "DK Material (nd.out)"
      Height          =   615
      Index           =   4
      Left            =   3480
      TabIndex        =   24
      Top             =   3000
      Width           =   3255
      Begin VB.CommandButton cmdDeleteMaterial 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdMaterialRecords 
         Caption         =   "Create"
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraTable 
      Caption         =   "ARS Nozzle (ars.out)"
      Height          =   615
      Index           =   11
      Left            =   3480
      TabIndex        =   52
      Top             =   3600
      Width           =   3255
      Begin VB.CommandButton cmdARSRecords 
         Caption         =   "Create"
         Height          =   255
         Left            =   1080
         TabIndex        =   54
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteARS 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraTable 
      Caption         =   "Aircraft (aircraft.out)"
      Height          =   615
      Index           =   5
      Left            =   3480
      TabIndex        =   27
      Top             =   4200
      Width           =   3255
      Begin VB.CommandButton cmdAircraftRecords 
         Caption         =   "Create"
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteAircraft 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAircraftRecords_Click()
  Me.MousePointer = 11
  cmdAircraftRecords.Enabled = False
  AddToLog "Creating Aircraft"
  AddAircraftTables "Aircraft"
  AddAircraftRecords "Aircraft", "aircraft_ag.out"
  AddAircraftTables "AircraftFS"
  AddAircraftRecords "AircraftFS", "aircraft_fs.out"
  cmdAircraftRecords.Enabled = True
  Me.MousePointer = 0
End Sub

Private Sub cmdARSRecords_Click()
  Me.MousePointer = 11
  cmdARSRecords.Enabled = False
  AddToLog "Creating ARS Data"
  AddARSTables
  AddARSRecords
  cmdARSRecords.Enabled = True
  Me.MousePointer = 0
End Sub

Private Sub cmdBasicRecords_Click()
  Me.MousePointer = 11
  cmdBasicRecords.Enabled = False
  AddToLog "Creating Basic Data"
  AddBasicTables
  AddBasicRecords
  cmdBasicRecords.Enabled = True
  Me.MousePointer = 0
End Sub

Private Sub cmdCanopyRecords_Click()
  Me.MousePointer = 11
  cmdCanopyRecords.Enabled = False
  AddToLog "Creating Canopy"
  AddCanopyTables "Canopy"
  AddCanopyRecords "Canopy", "fs_can.out"
  cmdCanopyRecords.Enabled = True
  Me.MousePointer = 0
End Sub

Private Sub cmdComponent_Click()
  Me.MousePointer = 11
  cmdComponent.Enabled = False
  AddToLog "Creating Components"
  AddComponentTables
  AddComponentRecords
  cmdComponent.Enabled = True
  Me.MousePointer = 0
End Sub

Private Sub cmdCreateDB_Click()
  cmdCreateDB.Enabled = False
  AddToLog "Creating empty DB " + GD.DBDirPath & GD.DBFileName
  CreateAgdriftDatabase
  cmdCreateDB.Enabled = True
End Sub


Private Sub cmdDeleteAircraft_Click()
  cmdDeleteAircraft.Enabled = False
  AddToLog "Deleting Aircraft"
  DeleteAircraftTables "Aircraft"
  DeleteAircraftTables "AircraftFS"
  cmdDeleteAircraft.Enabled = True
End Sub

Private Sub cmdDeleteARS_Click()
  cmdDeleteARS.Enabled = False
  AddToLog "Deleting ARS"
  DeleteARSTables
  cmdDeleteARS.Enabled = True
End Sub

Private Sub cmdDeleteBasic_Click()
  cmdDeleteBasic.Enabled = False
  AddToLog "Deleting Basic Data"
  DeleteBasicTables
  cmdDeleteBasic.Enabled = True
End Sub

Private Sub cmdDeleteCanopy_Click()
  cmdDeleteCanopy.Enabled = False
  AddToLog "Deleting Canopy"
  DeleteCanopyTables
  cmdDeleteCanopy.Enabled = True
End Sub

Private Sub cmdDeleteComponents_Click()
  cmdDeleteComponents.Enabled = False
  AddToLog "Deleting Components"
  DeleteComponentsTables
  cmdDeleteComponents.Enabled = True
End Sub

Private Sub cmdDeleteDB_Click()
  On Error GoTo ErrHandlerDel
  cmdDeleteDB.Enabled = False
  AddToLog "Deleting " + GD.DBDirPath & GD.DBFileName
  Kill GD.DBDirPath & GD.DBFileName
  cmdDeleteDB.Enabled = True
  Exit Sub

ErrHandlerDel:
  AddToLog "Database did not exist"
  cmdDeleteDB.Enabled = True
  Exit Sub

End Sub

Private Sub cmdDeleteDropsize_Click()
  cmdDeleteDropsize.Enabled = False
  AddToLog "Deleting Dropsize"
  DeleteDropsizeTables 0
  DeleteDropsizeTables 1
  cmdDeleteDropsize.Enabled = True
End Sub

Private Sub cmdDeleteEvaporation_Click()
  cmdDeleteEvaporation.Enabled = False
  AddToLog "Deleting Evaporation"
  DeleteEvaporationTables
  cmdDeleteEvaporation.Enabled = True
End Sub


Private Sub cmdDeleteInfo_Click()
  cmdDeleteInfo.Enabled = False
  AddToLog "Deleting Info"
  DeleteInfoTables
  cmdDeleteInfo.Enabled = True
End Sub

Private Sub cmdDeleteMaterial_Click()
  cmdDeleteMaterial.Enabled = False
  AddToLog "Deleting Materials"
  DeleteMaterialTables
  cmdDeleteMaterial.Enabled = True
End Sub

Private Sub cmdDeleteNozzle_Click()
  cmdDeleteNozzle.Enabled = False
  AddToLog "Deleting Nozzles"
  DeleteNozzleTables
  AddToLog "Deleting Nozzle type/mfg's"
  DeleteNozzleTMTables
  cmdDeleteNozzle.Enabled = True
End Sub

Private Sub cmdDeleteTrial_Click()
  cmdDeleteTrial.Enabled = False
  AddToLog "Deleting Field Trial Data"
  DeleteTrialTables
  cmdDeleteTrial.Enabled = True
End Sub

Private Sub cmdDeleteWindrose_Click()
  cmdDeleteWindrose.Enabled = False
  AddToLog "Deleting Wind Rose Data"
  DeleteWindroseTables
  cmdDeleteWindrose.Enabled = True
End Sub


Private Sub cmdDropsize_Click()
  Me.MousePointer = 11
  cmdDropsize.Enabled = False
  AddToLog "Creating Dropsize"
  AddDropsizeTables 0
  AddDropsizeRecords 0
  AddDropsizeTables 1
  AddDropsizeRecords 1
  cmdDropsize.Enabled = True
  Me.MousePointer = 0
End Sub

Private Sub cmdEvaporationRecords_Click()
  Me.MousePointer = 11
  cmdEvaporationRecords.Enabled = False
  AddToLog "Creating Evaporation"
  AddEvaporationTables
  AddEvaporationRecords
  cmdEvaporationRecords.Enabled = True
  Me.MousePointer = 0
End Sub

Private Sub cmdInfoRecords_Click()
  Me.MousePointer = 11
  cmdInfoRecords.Enabled = False
  AddToLog "Creating Info"
  AddInfoTables
  AddInfoRecords
  cmdInfoRecords.Enabled = True
  Me.MousePointer = 0
End Sub

Private Sub cmdMaterialRecords_Click()
  Me.MousePointer = 11
  cmdMaterialRecords.Enabled = False
  AddToLog "Creating Materials"
  AddMaterialTables
  AddMaterialRecords
  cmdMaterialRecords.Enabled = True
  Me.MousePointer = 0
End Sub

Private Sub cmdNozzleRecords_Click()
  Me.MousePointer = 11
  cmdNozzleRecords.Enabled = False
  AddToLog "Creating Nozzles"
  AddNozzleTables
  AddNozzleRecords
  AddToLog "Creating Nozzle type/mfg's"
  AddNozzleTMTables
  AddNozzleTMRecords
  cmdNozzleRecords.Enabled = True
  Me.MousePointer = 0
End Sub

Private Sub cmdRegenerate_Click()
  cmdRegenerate.Enabled = False
  AddToLog "Rebuilding " + GD.DBDirPath & GD.DBFileName + " from scratch"
  cmdDeleteDB_Click
  cmdCreateDB_Click
  cmdInfoRecords_Click
  Select Case GD.LibType
  Case 0, 1  'World, SDTF libraries
    cmdDropsize_Click
    cmdComponent_Click
    cmdEvaporationRecords_Click
    cmdNozzleRecords_Click
    cmdMaterialRecords_Click
    cmdARSRecords_Click
    cmdAircraftRecords_Click
    cmdBasicRecords_Click
    cmdCanopyRecords_Click
    If GD.LibType = 1 Then cmdTrialRecords_Click
  Case 2   'MAA Library
    cmdWindroseRecords_Click
  End Select
  AddToLog "Rebuild complete."
  cmdRegenerate.Enabled = True
End Sub

Private Sub cmdRegenerateAll_Click()
  Dim c As Control
  cmdRegenerateAll.Enabled = False
  AddToLog "Rebuilding All Libraries in" + GD.DBDirPath
  For Each c In optDBOutputType
    c.Value = True
    cmdRegenerate_Click
  Next
  AddToLog "Rebuild All Libraries complete."
  cmdRegenerateAll.Enabled = True
End Sub

Private Sub cmdRegenerateStdProp_Click()
  cmdRegenerateStdProp.Enabled = False
  AddToLog "Rebuilding Std,Prop Libraries in" + GD.DBDirPath
  optDBOutputType(0).Value = True
  cmdRegenerate_Click
  optDBOutputType(1).Value = True
  cmdRegenerate_Click
  AddToLog "Rebuild Std,Prop Libraries complete."
  cmdRegenerateStdProp.Enabled = True
End Sub

Private Sub cmdTrialRecords_Click()
  Me.MousePointer = 11
  cmdTrialRecords.Enabled = False
  AddToLog "Creating Field Trial Data"
  AddTrialTables
  AddTrialRecords
  cmdTrialRecords.Enabled = True
  Me.MousePointer = 0
End Sub

Private Sub cmdWindroseRecords_Click()
  Me.MousePointer = 11
  cmdWindroseRecords.Enabled = False
  AddToLog "Creating Wind Rose Data"
  AddWindroseTables
  AddWindroseRecords
  cmdWindroseRecords.Enabled = True
  Me.MousePointer = 0
End Sub

Private Sub Form_Load()
  CenterForm Me
  Me.Show  'show the form so that the next command,
           'which uses SetFocus, will work
  
  optDB(0).Value = True  'default to test database
  optDBOutputType(1).Value = True 'default to SDTF Library
  GD.SrcPath = App.Path & "\src\" 'path to source files
  lblLibVersion = Format$(LIBRARYVERSION)
End Sub

Private Sub mnuExit_Click()
  End
End Sub

Private Sub optDB_Click(Index As Integer)
  Select Case Index
    Case 0
      txtDBpath.Text = LCase$(App.Path) & "\"
      AddToLog "Test database location " + txtDBpath.Text + " selected"
    Case 1
      txtDBpath.Text = "c:\my documents\work\agdrift\"
      AddToLog "Actual database location " + txtDBpath.Text + " selected"
  End Select
End Sub

Private Sub optDBOutputType_Click(Index As Integer)
  GD.LibType = Index
  Select Case GD.LibType
  Case 0 'Standard Library
    txtDBfile.Text = "agdrift.mdb"
    GD.Src.airbasic = "airbasic.out"
    GD.Src.aircraft = "aircraft.out"
    GD.Src.atomize = "atomize.wld"
    GD.Src.atomizeFS = "fs_dsd.out"
    GD.Src.asaedep = "asaedep.out"
    GD.Src.basicdrp = "basicdrp.out"
    GD.Src.dropsize = "dropsize.out"
    GD.Src.evap = "evap.wld"
    GD.Src.fieldrun = ""
    GD.Src.nd = "nd.wld"
    GD.Src.nozbasic = "nozbasic.out"
    GD.Src.nozzle = "nozzle.out"
    GD.Src.nozzletm = "nozzletm.out"
    GD.Src.arsnoz = "ars.out"
    GD.Src.sngldep = "sngldep.out"
    GD.Src.subst = "subst.wld"
    GD.Src.canopy = "fs_can.out"
    GD.Src.infbasic = "infbasic.out"
    AddToLog "Standard Library source files selected"
  Case 1 'SDTF Library
    txtDBfile.Text = "agdsdtf.mdb"
    GD.Src.airbasic = "airbasic.out"
    GD.Src.aircraft = "aircraft.out"
    GD.Src.atomize = "atomize.out"
    GD.Src.atomizeFS = "fs_dsd.out"
    GD.Src.asaedep = "asaedep.out"
    GD.Src.basicdrp = "basicdrp.out"
    GD.Src.dropsize = "dropsize.out"
    GD.Src.evap = "evap.out"
    GD.Src.fieldrun = "fieldrun.out"
    GD.Src.nd = "nd.out"
    GD.Src.nozbasic = "nozbasic.out"
    GD.Src.nozzle = "nozzle.out"
    GD.Src.nozzletm = "nozzletm.out"
    GD.Src.arsnoz = "ars.out"
    GD.Src.sngldep = "sngldep.out"
    GD.Src.subst = "subst.out"
    GD.Src.canopy = "fs_can.out"
    GD.Src.infbasic = "infbasic.out"
    AddToLog "SDTF Library source files selected"
  Case 2 'SDTF Library
    txtDBfile.Text = "agdmaa.mdb"
    AddToLog "MAA Library source files selected"
  End Select
End Sub

Private Sub txtDBfile_Change()
  'update DB file name
  GD.DBFileName = txtDBfile.Text
End Sub

Private Sub txtDBpath_Change()
  'update DB path
  GD.DBDirPath = txtDBpath.Text
End Sub


