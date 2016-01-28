VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTBSBDReceptors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Discrete Receptors"
   ClientHeight    =   4395
   ClientLeft      =   1665
   ClientTop       =   1650
   ClientWidth     =   9480
   HelpContextID   =   1496
   Icon            =   "TBSBDRCP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4395
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImport 
      Caption         =   "I&mport"
      Height          =   375
      HelpContextID   =   1496
      Left            =   2640
      TabIndex        =   5
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "E&xport"
      Height          =   375
      HelpContextID   =   1496
      Left            =   3600
      TabIndex        =   6
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtEditGrid 
      BorderStyle     =   0  'None
      Height          =   285
      HelpContextID   =   1496
      Left            =   960
      TabIndex        =   7
      Text            =   "txtEditGrid"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clea&r"
      Height          =   255
      HelpContextID   =   1496
      Left            =   1800
      TabIndex        =   4
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   255
      HelpContextID   =   1496
      Left            =   960
      TabIndex        =   3
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   255
      HelpContextID   =   1496
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1496
      Left            =   8520
      TabIndex        =   1
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1496
      Left            =   7560
      TabIndex        =   0
      Top             =   3960
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grdTable 
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      WordWrap        =   -1  'True
      Appearance      =   0
   End
End
Attribute VB_Name = "frmTBSBDReceptors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public eg As New clsEditGrid
Public Cancelled As Boolean

'An escape key means Cancel. Press it and the form goes
'away. Normally, we would set the Cancel property to True
'for the Cancel button and we would be done, but this form
'contain EditGrids. EditGrids rely on the Escape key to
'cancel an edit. If the Cancel property of the Cancel
'button is true, this behavior doesn't work. The desired
'behavior is for the Escape key to cancel an edit operation
'in an EditGrid, and to dismiss the form in all other cases.
'To that end, we employ this method:
'- Set the Cancel property to False for the Cancel button
'- Set the KeyPreview property to True for the form
'- Examine KeyPress events at the form level and pass Escapes
'  through to EditGrid text boxes, and dismiss the form for
'  all other cases.
'Here we define a collection to hold all EditGrid text boxes
'for this form. If, when an escape key is pressed, one of
'the controls in this collection is the ActiveControl, the
'program continues normally and the Text control receives a
'KeyPress event. If the ActiveControl is not one in the
'collection, the cmdCancel_Click event routine is invoked.
'See Form_KeyPress below.
Private ControlsThatMayReceiveEscape As New Collection

Private Sub ImportReceptors()
'Import a new set of discrete receptors
'from a two-column text file and load it into
'the grid. The file is assumed to be in metric.

  Dim fn As String
  Dim n As Integer
  Dim NumDisc As Long        'Number of Discrete Receptors
  Dim DiscType(99) As Single 'Receptor Type 0=
  Dim DiscX(99) As Single    'Receptor X Position (m)
  Dim DiscY(99) As Single    'Receptor Y Position (m)
  Dim DiscZ(99) As Single    'Receptor Z Position (m)
  Dim DiscI(99) As Single    'Receptor Normal X component
  Dim DiscJ(99) As Single    'Receptor Normal Y component
  Dim DiscK(99) As Single    'Receptor Normal Z component
  Dim DiscSize(99) As Single 'Receptor size (m)
  
  If FileDialog(FD_OPEN, FD_TYPE_TEXT, fn) Then  'get a filename
    'Open the file
    On Error GoTo ImportReceptorErrHand
    OpenFileAndSkipComments fn, 1
    n = 0
    While Not EOF(1)
      Input #1, DiscType(n), _
                DiscX(n), DiscY(n), DiscZ(n), _
                DiscI(n), DiscJ(n), DiscK(n), _
                DiscSize(n)
      n = n + 1
    Wend
    NumDisc = n
    Close #1

    eg.ClearAll
    eg.ArrayToGrid 1, n, DiscType()
    eg.ArrayToGrid 2, n, DiscX(), UN_LENGTH
    eg.ArrayToGrid 3, n, DiscY(), UN_LENGTH
    eg.ArrayToGrid 4, n, DiscZ(), UN_LENGTH
    eg.ArrayToGrid 5, n, DiscI()
    eg.ArrayToGrid 6, n, DiscJ()
    eg.ArrayToGrid 7, n, DiscK()
    eg.ArrayToGrid 8, n, DiscSize(), UN_SMLENGTH
  End If
  Exit Sub

ImportReceptorErrHand:
  Close #1
  MsgBox "Error importing file: " + fn + vbCr + Error$(Err), _
    vbCritical + vbOKOnly
  Exit Sub
End Sub

Private Sub ExportReceptors()
'Export a set of discrete receptors

  Dim fn As String
  Dim n As Integer
  Dim i As Integer
  Dim NumDisc As Long        'Number of Discrete Receptors
  Dim DiscType(99) As Single 'Receptor Type 0=
  Dim DiscX(99) As Single    'Receptor X Position (m)
  Dim DiscY(99) As Single    'Receptor Y Position (m)
  Dim DiscZ(99) As Single    'Receptor Z Position (m)
  Dim DiscI(99) As Single    'Receptor Normal X component
  Dim DiscJ(99) As Single    'Receptor Normal Y component
  Dim DiscK(99) As Single    'Receptor Normal Z component
  Dim DiscSize(99) As Single 'Receptor size (m)
  
  If FileDialog(FD_OPEN, FD_TYPE_TEXT, fn) Then  'get a filename
    On Error GoTo ExportReceptorErrHand
    'Open the file
    eg.GridToArray 1, n, DiscType()
    eg.GridToArray 2, n, DiscX(), UN_LENGTH
    eg.GridToArray 3, n, DiscY(), UN_LENGTH
    eg.GridToArray 4, n, DiscZ(), UN_LENGTH
    eg.GridToArray 5, n, DiscI()
    eg.GridToArray 6, n, DiscJ()
    eg.GridToArray 7, n, DiscK()
    eg.GridToArray 8, n, DiscSize(), UN_SMLENGTH
    NumDisc = n
    
    Open fn For Output As #1
    For n = 0 To NumDisc - 1
      Print #1, DiscType(n); _
                DiscX(n); DiscY(n); DiscZ(n); _
                DiscI(n); DiscJ(n); DiscK(n); _
                DiscSize(n)
    Next
    Close #1

  End If
  Exit Sub

ExportReceptorErrHand:
  Close #1
  MsgBox "Error exporting file: " + fn + vbCr + Error$(Err), _
    vbCritical + vbOKOnly
  Exit Sub
End Sub

Private Sub cmdCancel_Click()
  Cancelled = True
  Me.Hide
End Sub

Private Sub cmdClear_Click()
  eg.ClearSelected
End Sub

Private Sub cmdDelete_Click()
  eg.DeleteRow
End Sub

Private Sub cmdExport_Click()
  ExportReceptors
End Sub

Private Sub cmdImport_Click()
  ImportReceptors
End Sub

Private Sub cmdInsert_Click()
  eg.InsertRow
End Sub

Private Sub cmdOk_Click()
  Cancelled = False
  Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim c As Control
  If KeyAscii = 27 Then
    For Each c In ControlsThatMayReceiveEscape
      If c Is Me.ActiveControl Then
        Exit Sub
      End If
    Next
    cmdCancel_Click
  End If
End Sub

Private Sub Form_Load()
  CenterForm Me

  'Initialize the collection of controls that may receive
  'an escape character. This allows Escape to dismiss the
  'form OR abort an EditGrid edit.
  With ControlsThatMayReceiveEscape
    .Add txtEditGrid
  End With
  
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

Private Sub grdTable_Scroll()
  eg.GridScroll
End Sub

Private Sub txtEditGrid_KeyDown(KeyCode As Integer, Shift As Integer)
  eg.TextKeyDown KeyCode, Shift
End Sub
    
Private Sub txtEditGrid_KeyPress(KeyAscii As Integer)
  eg.TextKeyPress KeyAscii
  If KeyAscii = 13 Then KeyAscii = 0
  If KeyAscii = 27 Then KeyAscii = 0
End Sub
    
Private Sub txtEditGrid_LostFocus()
  eg.TextLostFocus
End Sub

