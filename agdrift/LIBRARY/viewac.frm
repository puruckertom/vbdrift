VERSION 4.00
Begin VB.Form frmViewAircraft 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Aircraft"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   1350
   ClientWidth     =   5445
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
   Height          =   4680
   Left            =   60
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   5445
   Top             =   1005
   Width           =   5565
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      DataField       =   "EngHoriz"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      DataField       =   "EngFwd"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      DataField       =   "EngVert"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      DataField       =   "PropRad"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      DataField       =   "PropRPM"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4320
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      DataField       =   "PropEff"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4320
      TabIndex        =   6
      Text            =   "PropEff"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "PlanArea"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      DataField       =   "DragCoef"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Text            =   "DragCoef"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      DataField       =   "Weight"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      DataField       =   "BiplSep"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "TypSpeed"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "SemiSpan"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "Type"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Text            =   "Type"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Text            =   "Name"
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   3600
      Width           =   735
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Aircraft"
      Connect         =   ""
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   270
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aircraft"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Eng horiz (ft)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3000
      TabIndex        =   27
      Top             =   2280
      Width           =   1110
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Eng fwd (ft)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3000
      TabIndex        =   26
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Eng vert (ft)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3000
      TabIndex        =   25
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Prop Rad (ft)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3000
      TabIndex        =   24
      Top             =   1200
      Width           =   1110
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Prop RPM"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3000
      TabIndex        =   23
      Top             =   840
      Width           =   870
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "PropEff"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3000
      TabIndex        =   22
      Top             =   480
      Width           =   645
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Planform area (ft2)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   2640
      Width           =   1590
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "DragCoef"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   810
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Weight (lbs)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Biplane sep (ft)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   1305
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Typ speed (mph)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Semi-span (ft)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   435
   End
End
Attribute VB_Name = "frmViewAircraft"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  CenterForm Me
  Data1.DatabaseName = GD.DBDirPath & GD.DBFileName
End Sub

