VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTBSBDBoundary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spray Block/Area Coverage Boundary"
   ClientHeight    =   4740
   ClientLeft      =   4290
   ClientTop       =   1440
   ClientWidth     =   4680
   Icon            =   "TBSBDBND.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4740
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExport 
      Caption         =   "E&xport"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "I&mport"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtEditGrid 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "txtEditGrid"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clea&r"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   4320
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grdTable 
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6588
      _Version        =   393216
      WordWrap        =   -1  'True
      Appearance      =   0
   End
End
Attribute VB_Name = "frmTBSBDBoundary"
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

'Since this form is used for more than one purpose, we must
'assign HelpContextID programmatically rather than hard-
'wiring it. This collection helps us do that.
Private ControlsForHelpContextID As New Collection

Public Sub SetHelpContextID(HCID As Integer)
'Set the HelpContextId of all the appropriate controls
  Dim c As Control
  For Each c In ControlsForHelpContextID
    c.HelpContextID = HCID
  Next
End Sub

Private Sub ImportBound()
'Import a new Spray Block Boundary
'from a two-column text file and load it into
'the grid. The file is assumed to be in metric.

  Dim fn As String
  Dim n As Integer
  Dim X(99) As Single
  Dim Y(99) As Single
  
  If FileDialog(FD_OPEN, FD_TYPE_TEXT, fn) Then  'get a filename
    'Open the file
    On Error GoTo ImportBoundErrHand
    OpenFileAndSkipComments fn, 1
    
    n = 0
    While Not EOF(1)
      Input #1, X(n), Y(n)
      n = n + 1
    Wend
    Close #1

    eg.ArrayToGrid 1, n, X(), UN_LENGTH
    eg.ArrayToGrid 2, n, Y(), UN_LENGTH
  End If
  Exit Sub

ImportBoundErrHand:
  Close #1
  MsgBox "Error importing file: " + fn + vbCr + Error$(Err), _
    vbCritical + vbOKOnly
  Exit Sub
End Sub

Private Sub ExportBound()
'Import a new Spray Block Boundary
'from a two-column text file and load it into
'the grid. The file is assumed to be in metric.

  Dim fn As String
  Dim n As Integer
  Dim i As Integer
  Dim X(99) As Single
  Dim Y(99) As Single
  
  If FileDialog(FD_OPEN, FD_TYPE_TEXT, fn) Then  'get a filename
    On Error GoTo ExportBoundErrHand
    'Open the file
    eg.GridToArray 1, n, X(), UN_LENGTH
    eg.GridToArray 2, n, Y(), UN_LENGTH
    Open fn For Output As #1
    For i = 0 To n - 1
      Print #1, X(i); Y(i)
    Next
    Close #1

  End If
  Exit Sub

ExportBoundErrHand:
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
  ExportBound
End Sub

Private Sub cmdImport_Click()
  ImportBound
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
  
  'Initialize the collection of controls that will have
  'their HelpContextID set
  With ControlsForHelpContextID
    .Add cmdInsert
    .Add cmdDelete
    .Add cmdClear
    .Add cmdImport
    .Add cmdExport
    .Add cmdOK
    .Add cmdCancel
    .Add grdTable
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
  If KeyAscii = 13 Then KeyAscii = 0 'absorb carriage returns
  If KeyAscii = 27 Then KeyAscii = 0 'absorb escape
End Sub
    
Private Sub txtEditGrid_LostFocus()
  eg.TextLostFocus
End Sub

