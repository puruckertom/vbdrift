VERSION 4.00
Begin VB.Form frmViewEvaporation 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Evaporation"
   ClientHeight    =   2340
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
   Height          =   2745
   Left            =   3150
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3420
   Top             =   1950
   Width           =   3540
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "EvaporationRate"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Text            =   "Evaporation Rate"
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
      Top             =   1800
      Width           =   735
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Evaporation"
      Connect         =   ""
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   270
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Evaporation"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "Substance"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "Substance"
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "NonvolFraction"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "Nonvolatile Fraction"
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "frmViewEvaporation"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  CenterForm Me
  Data1.DatabaseName = GD.DBDirPath & GD.DBFileName
End Sub

