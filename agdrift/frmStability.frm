VERSION 5.00
Begin VB.Form frmStability 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Atmospheric Stability"
   ClientHeight    =   2970
   ClientLeft      =   4695
   ClientTop       =   1905
   ClientWidth     =   5130
   ForeColor       =   &H80000008&
   HelpContextID   =   1548
   Icon            =   "frmStability.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2970
   ScaleWidth      =   5130
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1548
      Left            =   3240
      TabIndex        =   0
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1548
      Left            =   4200
      TabIndex        =   1
      Top             =   2520
      Width           =   855
   End
   Begin VB.Frame fraDisp 
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cboInsolation 
         Height          =   315
         HelpContextID   =   1545
         Index           =   1
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   2535
      End
      Begin VB.ComboBox cboInsolation 
         Height          =   315
         HelpContextID   =   1545
         Index           =   0
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton optInsolation 
         Caption         =   "Night (sunset to 1 hr after sunrise)"
         ForeColor       =   &H80000008&
         Height          =   255
         HelpContextID   =   1548
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2775
      End
      Begin VB.OptionButton optInsolation 
         Caption         =   "Day (1 hr after sunrise to sunset)"
         ForeColor       =   &H80000008&
         Height          =   255
         HelpContextID   =   1548
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblInsolation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Solar Insolation:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label lblInsolation 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cloud Cover:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   975
         TabIndex        =   7
         Top             =   1680
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmStability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: frmStability.frm,v 1.2 2008/10/22 17:26:06 tom Exp $

Private Sub DataToForm()
  Dim c As Control
  Dim i As Integer
  For Each c In cboInsolation()
    c.ListIndex = 0
    For i = 0 To c.ListCount - 1
      If c.ItemData(i) = UD.MET.Insolation Then
        c.ListIndex = i
        Exit For
      End If
    Next
  Next
  Select Case UD.MET.Insolation
  Case 0 To 3: optInsolation(0).Value = True
  Case 4 To 6: optInsolation(1).Value = True
  End Select
End Sub

Private Sub FormToData()
  Dim c As Control
  Dim cbo As ComboBox
  For Each c In optInsolation()
    If c.Value Then
      Set cbo = cboInsolation(c.Index)
      UD.MET.Insolation = cbo.ItemData(cbo.ListIndex)
      Exit For
    End If
  Next
  
  UpdateDataChangedFlag True 'Data was changed
  UC.Valid = False 'Calcs need to be redone
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  FormToData
  Unload Me
End Sub

Private Sub Form_Load()
  Dim i As Integer
  
  CenterForm Me
  
  With cboInsolation(0)
    .Clear
    For i = 0 To 3
      .AddItem GetTypeNameStability(i)
      .ItemData(.NewIndex) = i
    Next
    .ListIndex = 3
  End With
  
  With cboInsolation(1)
    .Clear
    For i = 4 To 6
      .AddItem GetTypeNameStability(i)
      .ItemData(.NewIndex) = i
    Next
    .ListIndex = 0
  End With
    
  DataToForm
End Sub

Private Sub optInsolation_Click(Index As Integer)
  Dim i As Integer
  
  For i = 0 To 1
    lblInsolation(i).Enabled = (i = Index)
    cboInsolation(i).Enabled = (i = Index)
  Next
End Sub
