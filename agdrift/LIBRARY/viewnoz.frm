VERSION 4.00
Begin VB.Form frmViewNozzles 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Nozzles"
   ClientHeight    =   3255
   ClientLeft      =   3210
   ClientTop       =   2295
   ClientWidth     =   3420
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   3660
   Left            =   3150
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   3420
   Top             =   1950
   Width           =   3540
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      DataField       =   "Diameter"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Text            =   "Diameter"
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      DataField       =   "VMD"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Text            =   "VMD"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "F220"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Text            =   "F220"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "F50"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Text            =   "F50"
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Close"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Nozzles"
      Connect         =   ""
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   270
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Nozzles"
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "Nozzle"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "Nozzle"
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "F141"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "F141"
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "frmViewNozzles"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  CenterForm Me
  Data1.DatabaseName = GD.DBDirPath & GD.DBFileName
End Sub

