VERSION 4.00
Begin VB.Form frmTBSBDFlightLines 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flight Lines"
   ClientHeight    =   4395
   ClientLeft      =   2490
   ClientTop       =   2460
   ClientWidth     =   5490
   Height          =   4800
   Icon            =   "TBSBDFLT.frx":0000
   Left            =   2430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Top             =   2115
   Width           =   5610
   Begin VB.TextBox txtGridEdit 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Text            =   "txtGridEdit"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clea&r"
      Height          =   255
      HelpContextID   =   1100
      Left            =   1800
      TabIndex        =   4
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   255
      HelpContextID   =   1100
      Left            =   960
      TabIndex        =   3
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   255
      HelpContextID   =   1100
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1227
      Left            =   4560
      TabIndex        =   1
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1227
      Left            =   3600
      TabIndex        =   0
      Top             =   3960
      Width           =   855
   End
   Begin MSGrid.Grid grdTable 
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5295
      _version        =   65536
      _extentx        =   9340
      _extenty        =   6800
      _stockprops     =   77
      backcolor       =   16777215
      mouseicon       =   "TBSBDFLT.frx":030A
   End
End
Attribute VB_Name = "frmTBSBDFlightLines"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Public eg As New clsEditGrid
Public Cancelled As Boolean

Private Sub cmdCancel_Click()
  Cancelled = True
  Me.Hide
End Sub

Private Sub cmdOk_Click()
  Cancelled = False
  Me.Hide
End Sub

Private Sub Form_Load()
  Dim wid As Single
  Dim i As Integer
  
  CenterForm Me

  Const MAX_BOUND = 50
  
  'Leaf Area Index Grid
  With grdTable
    .FixedRows = 1
    .FixedCols = 1
    .Rows = MAX_BOUND + .FixedRows
    .cols = 4 + .FixedCols
    .Row = 0
    .Col = 1
    .Text = "Start X (" + UnitsName(UN_LENGTH) + ")"
    .Col = 2
    .Text = "Start Y (" + UnitsName(UN_LENGTH) + ")"
    .Col = 3
    .Text = "End X (" + UnitsName(UN_LENGTH) + ")"
    .Col = 4
    .Text = "End Y (" + UnitsName(UN_LENGTH) + ")"
    .FixedAlignment(0) = 2
'    .ColAlignment(0) = 2
    .FixedAlignment(1) = 2
    .ColAlignment(1) = 2
    .FixedAlignment(2) = 2
    .ColAlignment(2) = 2
    .FixedAlignment(3) = 2
    .ColAlignment(3) = 2
    .FixedAlignment(4) = 2
    .ColAlignment(4) = 2
  
    'set Column widths
    .RowHeight(0) = 430  'set height of first row
    .ColWidth(0) = 500   'set width of first column
    wid = CSng(.Width - .ColWidth(0) - 350) / CSng(.cols - .FixedCols)
    For i = 1 To .cols - 1
      .ColWidth(i) = wid
    Next
    .Col = 0
  End With
  eg.Setup grdTable, txtGridEdit, MAX_BOUND
End Sub

Private Sub grdTable_DblClick()
  eg.GridDblClick
End Sub

Private Sub grdTable_KeyDown(KeyCode As Integer, Shift As Integer)
  eg.GridKeyDown KeyCode, Shift
End Sub

Private Sub grdTable_KeyPress(KeyAscii As Integer)
  eg.GridKeyPress KeyAscii
End Sub

Private Sub txtGridEdit_KeyDown(KeyCode As Integer, Shift As Integer)
  eg.TextKeyDown KeyCode, Shift
End Sub
    
Private Sub txtGridEdit_KeyPress(KeyAscii As Integer)
  eg.TextKeyPress KeyAscii
  If KeyAscii = 13 Then KeyAscii = 0
End Sub
    
Private Sub txtGridEdit_LostFocus()
  eg.TextLostFocus
End Sub

Private Sub cmdClear_Click()
  eg.ClearSelected
End Sub

Private Sub cmdDelete_Click()
  eg.DeleteRow
End Sub

Private Sub cmdInsert_Click()
  eg.InsertRow
End Sub


