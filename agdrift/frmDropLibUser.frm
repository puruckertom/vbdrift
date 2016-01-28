VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDropLibUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Drop Size Distribution User Library"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   HelpContextID   =   1459
   Icon            =   "frmDropLibUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Entry"
      Height          =   375
      HelpContextID   =   1459
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1459
      Left            =   1920
      TabIndex        =   8
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1459
      Left            =   2880
      TabIndex        =   7
      Top             =   4920
      Width           =   855
   End
   Begin VB.Frame fraName 
      Caption         =   "Name"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   3615
      Begin VB.ComboBox cboName 
         Height          =   315
         HelpContextID   =   1459
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame fraDropDist 
      Caption         =   "Drop Distribution"
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3615
      Begin MSFlexGridLib.MSFlexGrid grdDrop 
         Height          =   3135
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   4
         WordWrap        =   -1  'True
         Appearance      =   0
      End
      Begin VB.Label lblStats 
         Alignment       =   2  'Center
         Caption         =   "Relative Span:"
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   5
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label lblVMD 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label lblRelSpan 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "µm"
         Height          =   195
         Left            =   1680
         TabIndex        =   2
         Top             =   3720
         Width           =   210
      End
      Begin VB.Label lblStats 
         AutoSize        =   -1  'True
         Caption         =   "V0.5"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   1
         Top             =   3795
         Width           =   405
      End
      Begin VB.Label lblStats 
         AutoSize        =   -1  'True
         Caption         =   "D          :"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmDropLibUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OK As Boolean  'return value
Public DSDName As String 'returned DSD name

Private mNumDrop As Integer
Private mDiam(MAX_DROPS - 1) As Single
Private mMfrac(MAX_DROPS - 1) As Single

Public Sub SelectEntry(EntryName As String)
'Try to select the supplied entry in the combo
  Dim i As Integer
  
  For i = 0 To cboName.ListCount - 1
    If Trim$(cboName.List(i)) = Trim$(EntryName) Then
      cboName.ListIndex = i
      Exit For
    End If
  Next
End Sub

Private Sub ArrayToGrid()
'Transfer the given Drop Distribution to the grid control
  Dim g As Control
  Dim i As Integer
  Dim tot As Single
  
  Set g = grdDrop

  'set the appropriate number of rows
  g.Rows = mNumDrop + g.FixedRows
  'always leave one blank row if the array is blank
  If mNumDrop = 0 Then g.Rows = g.Rows + 1
  
  'transfer the distribution to the output control
  tot = 0
  For i = 0 To mNumDrop - 1
    tot = tot + mMfrac(i)
    g.Row = i + 1
    g.Col = 0
    g.Text = AGFormat$(i)
    g.Col = 1
    g.Text = AGFormat$(mDiam(i))
    g.Col = 2
    g.Text = AGFormat$(mMfrac(i))
    g.Col = 3
    g.Text = AGFormat$(tot)
  Next

  'Calculate and display statistics
  UpdateDSDStats
End Sub

Private Sub UpdateDSDStats()
'Calculate certain DSD stats and display them
  Dim VMD As Single
  Dim Span As Single
  Dim D10 As Single
  Dim D90 As Single
  Dim F141 As Single
  Dim DP As Single
  
  Call agdsrn(0, CLng(mNumDrop), mDiam(0), mMfrac(0), _
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

Private Sub cboName_Click()
  UserLibGetDropsizeRecord cboName.Text, mNumDrop, mDiam(), mMfrac()
  ArrayToGrid
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdDelete_Click()
  Dim i As Integer
  If Trim$(cboName.Text) <> "" Then
    If UserLibDeleteDropsizeRecord(cboName.Text) Then
      i = cboName.ListIndex
      cboName.RemoveItem i
      'try to keep the same place in the list
      If cboName.ListCount - 1 >= i Then
        cboName.ListIndex = i
      ElseIf cboName.ListCount > 0 Then
        cboName.ListIndex = cboName.ListCount - 1
      Else
        mNumDrop = 0
        ArrayToGrid
      End If
    End If
  End If
End Sub

Private Sub cmdOK_Click()
  OK = True
  DSDName = cboName.Text 'return DSD name
  Me.Hide
End Sub

Private Sub Form_Load()
  Dim g As Control
  Dim i As Integer
  Dim wid As Single
  Dim MatchingIndex As Integer
  Dim DB As Database
  Dim RS As Recordset
  
  CenterForm Me

  'init the grid
  Set g = grdDrop
  
  'set column headings and alignments
  g.Rows = 2
  g.cols = 4
  g.Row = 0
  g.Col = 1
  g.Text = "Average Diameter (µm)"
  g.Col = 2
  g.Text = "Incremental Volume Fraction"
  g.Col = 3
  g.Text = "Cumulative Volume Fraction"
  g.FixedAlignment(0) = flexAlignCenterCenter
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
  
  'init return value
  OK = False

  'Load the name combo
  MatchingIndex = -1 'index for entry that matches current aircraft
  If UserLibOpen(DB, False) Then
    If UserLibOpenRS(DB, "Dropsize", RS) Then
      If Not (RS.BOF And RS.EOF) Then
        While Not RS.EOF
          cboName.AddItem RS("Name")
          RS.MoveNext
        Wend
      End If
      RS.Close
    End If
    DB.Close
  End If
  
  If cboName.ListCount > 0 Then cboName.ListIndex = 0
End Sub
