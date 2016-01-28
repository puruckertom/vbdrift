VERSION 5.00
Begin VB.Form frmNotes 
   Caption         =   "Notes"
   ClientHeight    =   5040
   ClientLeft      =   1275
   ClientTop       =   1890
   ClientWidth     =   4905
   ForeColor       =   &H80000008&
   HelpContextID   =   1182
   Icon            =   "NOTES.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5040
   ScaleWidth      =   4905
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      HelpContextID   =   1182
      Left            =   1320
      TabIndex        =   4
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      HelpContextID   =   1182
      Left            =   2160
      TabIndex        =   2
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1182
      Left            =   3960
      TabIndex        =   1
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1182
      Left            =   3000
      TabIndex        =   0
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtNotes 
      Height          =   4335
      HelpContextID   =   1182
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "NOTES.frx":030A
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: notes.frm,v 1.8 2001/05/24 20:16:22 tom Exp $

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  FormToData
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim fn As String
  
  If Not FileDialog(FD_SAVEAS, FD_TYPE_TEXT, fn) Then
    Exit Sub
  End If
  On Error GoTo ErrHandcmdSave
  Open fn For Output As #1
  Print #1, GenFormData()
  Close #1
ExitcmdSave:
  Exit Sub
  
ErrHandcmdSave:
  MsgBox "Error writing file: " + fn + vbCr + Error$(Err)
  Resume ExitcmdSave
End Sub

Private Sub cmdPrint_Click()
'print the current UserData
  Dim BeginPage As Integer
  Dim EndPage As Integer
  Dim NumCopies As Integer
  Dim ReportText As String
  Dim i As Integer
  Dim pages As Integer

  If PrinterExists() Then
    If PrintDialog(BeginPage, EndPage, NumCopies) Then
      ReportText = GenFormData()
      For i = 1 To NumCopies
        PrintData ReportText, False, pages, Mag
      Next
    End If
  End If
End Sub

Private Sub DataToForm()
'Transfer stored data to form controls
  txtNotes.Text = UD.Notes
End Sub

Private Sub Form_Load()
  InitForm
End Sub

Private Sub Form_Resize()
  ResizeForm
End Sub

Private Sub FormToData()
  UD.Notes = txtNotes.Text
  UpdateDataChangedFlag True 'Data was changed
End Sub

Private Function GenFormData() As String
'Generate report text for this form to be used for printing
  
  Dim gfd As String  'temporary storage for report text
  Dim s As String        'workspace string

  gfd = "" 'start with a blank string
  
  AppendStr gfd, "AgDRIFT® Notes", True
  AppendStr gfd, "", True
  
  AppendStr gfd, CStr(txtNotes.Text), True

  GenFormData = gfd
End Function

Private Sub InitForm()
'Initialize this form
  CenterForm Me
  DataToForm
End Sub

Private Sub ResizeForm()
'position and size all the controls on the form
  Const MRGN = 120
  'guard against too small a form
  If Me.Height < 2000 Then Me.Height = 2000
  If Me.Width < 2120 Then Me.Width = 2120
  
  cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - MRGN
  cmdCancel.Top = Me.ScaleHeight - cmdCancel.Height - MRGN
  cmdOk.Left = cmdCancel.Left - cmdOk.Width - MRGN
  cmdOk.Top = cmdCancel.Top
  cmdPrint.Left = cmdOk.Left - cmdPrint.Width - MRGN
  cmdPrint.Top = cmdOk.Top
  cmdSave.Left = cmdPrint.Left - cmdSave.Width - MRGN
  cmdSave.Top = cmdPrint.Top
  txtNotes.Top = MRGN
  txtNotes.Left = MRGN
  txtNotes.Height = cmdCancel.Top - MRGN - txtNotes.Top
  txtNotes.Width = Me.ScaleWidth - MRGN - MRGN
End Sub

