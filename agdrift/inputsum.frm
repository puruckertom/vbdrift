VERSION 5.00
Begin VB.Form frmInputSummary 
   Caption         =   "Input Summary"
   ClientHeight    =   6225
   ClientLeft      =   1350
   ClientTop       =   3435
   ClientWidth     =   9480
   ForeColor       =   &H80000008&
   Icon            =   "INPUTSUM.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6225
   ScaleWidth      =   9480
   Begin VB.TextBox txtSummary 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "INPUTSUM.frx":030A
      Top             =   120
      Width           =   9255
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      HelpContextID   =   1173
      Left            =   6840
      TabIndex        =   2
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      HelpContextID   =   1173
      Left            =   7680
      TabIndex        =   1
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1173
      Left            =   8520
      TabIndex        =   0
      Top             =   5760
      Width           =   855
   End
End
Attribute VB_Name = "frmInputSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: inputsum.frm,v 1.6 2001/05/24 20:16:21 tom Exp $

Private Sub cmdOk_Click()
  Unload Me
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
      ReportText = GenReportText()
      For i = 1 To NumCopies
        PrintData ReportText, False, pages, Mag
      Next
    End If
  End If
End Sub

Private Sub cmdSave_Click()
  Dim fn As String
  
  If Not FileDialog(FD_SAVEAS, FD_TYPE_TEXT, fn) Then
    Exit Sub
  End If
  On Error GoTo ErrHandcmdSave
  Open fn For Output As #1
  Print #1, GenReportText()
  Close #1
ExitcmdSave:
  Exit Sub
  
ErrHandcmdSave:
  MsgBox "Error writing file: " + fn + vbCr + Error$(Err)
  Resume ExitcmdSave
End Sub

Private Sub Form_Load()
  InitForm
End Sub

Private Sub Form_Resize()
  ResizeForm
End Sub

Private Sub InitForm()
'initialize this form
  
  'Center the form on the screen
  CenterForm Me
  
  'initialize the data
  txtSummary = GenReportText
End Sub

Private Sub ResizeForm()
'Resize the controls on the form to match the new form size
  'guard against too small a form
  If Me.Height < 2000 Then Me.Height = 2000
  If Me.Width < 1000 Then Me.Width = 1000
  
  'Ok button
  cmdOk.Top = Me.ScaleHeight - cmdOk.Height - 120
  cmdOk.Left = Me.ScaleWidth - cmdOk.Width - 120

  'Print button
  cmdPrint.Top = cmdOk.Top
  cmdPrint.Left = cmdOk.Left - cmdPrint.Width - 120

  'Save button
  cmdSave.Top = cmdPrint.Top
  cmdSave.Left = cmdPrint.Left - cmdSave.Width - 120

  'List box
  txtSummary.Top = 120
  txtSummary.Left = 120
  txtSummary.Width = Me.ScaleWidth - 120 - 120
  txtSummary.Height = cmdOk.Top - txtSummary.Top - 120
End Sub

