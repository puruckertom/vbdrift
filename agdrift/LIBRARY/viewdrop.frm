VERSION 4.00
Begin VB.Form frmViewDropsize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Dropsize"
   ClientHeight    =   5130
   ClientLeft      =   3870
   ClientTop       =   1770
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
   Height          =   5535
   Left            =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   3420
   Top             =   1425
   Width           =   3540
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "DSLflag"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   10
      Text            =   "Flag"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Close"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   4680
      Width           =   735
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Dropsize"
      Connect         =   ""
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   270
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Dropsize"
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "Substance"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "Substance"
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "Nozzle"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Text            =   "Nozzle"
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "NozzleAngle"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "Angle"
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      DataField       =   "WindSpeed"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "Speed"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label lblMF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "mass frac"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblMF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "mass frac"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblMF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "mass frac"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblMF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "mass frac"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblMF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "mass frac"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "frmViewDropsize"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Data1_Reposition()
  GetMFdata
End Sub

Private Sub Form_Load()
  CenterForm Me
  Data1.DatabaseName = GD.DBDirPath & GD.DBFileName
End Sub

Private Sub GetMFdata()
'extract the mass frac data from the current record
'and stuff it into the label controls
   Dim MFfield As Field
   Dim FieldStr As String
   ReDim mf(31) As Single

   Set MFfield = Data1.Recordset.Fields("MassFrac")
   FieldStr = MFfield.GetChunk(0, MFfield.FieldSize())

   StringToArray mf(), FieldStr

   For i = 0 To 4
     lblMF(i).Caption = Format$(mf(i))
   Next
End Sub

