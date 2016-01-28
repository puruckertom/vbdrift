VERSION 5.00
Begin VB.Form frmNozzleRD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Regular Nozzle Distribution"
   ClientHeight    =   1905
   ClientLeft      =   1530
   ClientTop       =   2535
   ClientWidth     =   6510
   HelpContextID   =   1480
   Icon            =   "NOZZLERD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1905
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtExtent 
      DataField       =   "TypSpeed"
      DataSource      =   "Data1"
      Height          =   285
      HelpContextID   =   1480
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtSpacing 
      DataField       =   "TypSpeed"
      DataSource      =   "Data1"
      Height          =   285
      HelpContextID   =   1480
      Left            =   5040
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtNozzles 
      DataField       =   "TypSpeed"
      DataSource      =   "Data1"
      Height          =   285
      HelpContextID   =   1480
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1480
      Left            =   5520
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1480
      Left            =   4560
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblRegDistExtent 
      Alignment       =   2  'Center
      Caption         =   "Extent"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblRegDistSpacing 
      Alignment       =   2  'Center
      Caption         =   "Spacing"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblRegDistNozzles 
      Alignment       =   2  'Center
      Caption         =   "Nozzles"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblSpacingUnits 
      AutoSize        =   -1  'True
      Caption         =   "units"
      Height          =   195
      Left            =   5880
      TabIndex        =   7
      Top             =   600
      Width           =   330
   End
   Begin VB.Label lblRegDistNozzlesUnits 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   195
      Left            =   3720
      TabIndex        =   6
      Top             =   600
      Width           =   120
   End
   Begin VB.Label lblRegDistNote 
      Alignment       =   2  'Center
      Caption         =   "Enter values in any two boxes to generate regularly distributed nozzles."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmNozzleRD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cancelled As Boolean

Private Sub cmdCancel_Click()
  Cancelled = True
  Hide
End Sub

Private Sub cmdOk_Click()
  Cancelled = False
  Hide
End Sub

Private Sub Form_Load()
  CenterForm Me
  lblSpacingUnits.Caption = UnitsName(UN_LENGTH)
  Cancelled = True 'default value
End Sub
