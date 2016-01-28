VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExportToolbox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Toolbox Data"
   ClientHeight    =   4755
   ClientLeft      =   2880
   ClientTop       =   1920
   ClientWidth     =   6090
   ForeColor       =   &H80000008&
   HelpContextID   =   1136
   Icon            =   "EXPORTTB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4755
   ScaleWidth      =   6090
   Begin VB.Frame Frame1 
      Caption         =   "Notes"
      Height          =   1815
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtNotes 
         Height          =   1575
         HelpContextID   =   1537
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Text            =   "EXPORTTB.frx":030A
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame fraDelim 
      Caption         =   "Delimiter"
      Height          =   2895
      Left            =   0
      TabIndex        =   11
      Top             =   1800
      Width           =   3015
      Begin VB.TextBox txtColumns 
         Height          =   285
         HelpContextID   =   1136
         Left            =   1200
         TabIndex        =   9
         Top             =   2280
         Width           =   855
      End
      Begin VB.OptionButton optDelim 
         Caption         =   "&Fixed Width"
         Height          =   255
         HelpContextID   =   1136
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtDelim 
         Height          =   285
         HelpContextID   =   1136
         Left            =   1320
         TabIndex        =   7
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton optDelim 
         Caption         =   "&Other"
         Height          =   255
         HelpContextID   =   1136
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton optDelim 
         Caption         =   "Co&mma"
         Height          =   255
         HelpContextID   =   1136
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optDelim 
         Caption         =   "&Space"
         Height          =   255
         HelpContextID   =   1136
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optDelim 
         Caption         =   "&Tab"
         Height          =   255
         HelpContextID   =   1136
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Character:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblColumns 
         Caption         =   "Columns:"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1136
      Left            =   4320
      TabIndex        =   0
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1136
      Left            =   5160
      TabIndex        =   1
      Top             =   4320
      Width           =   855
   End
   Begin VB.Frame fraOpt 
      Caption         =   "Options"
      Height          =   855
      Left            =   3120
      TabIndex        =   12
      Top             =   1920
      Width           =   2895
      Begin VB.CheckBox cbxIncludeHeaders 
         Caption         =   "Include &Headers"
         Height          =   255
         HelpContextID   =   1136
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmExportToolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: exporttb.frm,v 1.9 2001/08/13 17:40:02 tom Exp $
'
' Toolbox Data Export Form

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  If ExportData() Then Unload Me
End Sub

Private Function ExportData() As Integer
'generate and export Toolbox plot data.
'Since this data is stored in a special place,
'we can access it directly.
'This function uses the same functions as the plotting form

  Dim i As Integer
  Dim s As String
  ReDim buffer(9) As String
  Dim fn As String
  Dim nsect As Integer
  Dim isect As Integer
  Dim np As Integer
  Dim ip As Integer
  Dim idelim As Integer
  Dim nycol As Integer
  Dim icol As Integer
  Dim slot As Integer
  Dim dlm As Integer
  Dim stat As Integer
  Dim cols As Integer
  Dim start As Integer
  Dim warnheaders As Integer
  Dim warnvalues As Integer
  Dim DoHeaders As Integer

  '
  'recover form settings
  '
  
  'Delimiter
  For i = 0 To 4
    If optDelim(i).Value Then idelim = i
  Next
  Select Case idelim
    Case 0  'tab
     dlm = 9
    Case 1  'space
      dlm = 32
    Case 2  'comma
      dlm = 44
    Case 3  'user-defined
      'the text is either a single delimiter character
      'or a two-char sequence representing a nonprintable
      Select Case Len(txtDelim.Text)
        Case 1
          dlm = Asc(txtDelim.Text)
        Case 2
          dlm = Asc(Right$(txtDelim.Text, 1)) - 64
        Case Else 'space
          dlm = 32
      End Select
    Case 4  'fixed columns
      cols = Val(txtColumns.Text)
      If cols > 0 Then
        dlm = -1  'negative dlm signals fixed columns
      Else
        MsgBox "Bad column width.", vbCritical + vbOKOnly
        ExportData = False
        Exit Function
      End If
  End Select
  'include header info?
  DoHeaders = False
  If cbxIncludeHeaders.Value = 1 Then DoHeaders = True
  '
  'get an output file
  '
  If Not FileDialog(FD_SAVEAS, FD_TYPE_TEXT, fn) Then
    ExportData = False
    Exit Function
  End If
  On Error GoTo ErrHandlerED
  Open fn For Output As #1

  'include the RunID, regardless of the header switch setting
  Print #1, "# "; GetRunID()
  
  'output the notes
  If txtNotes.Text <> "" Then
    start = 1
    Do
      Print #1, "# "; LineFromString(txtNotes.Text, start)
    Loop While start > 0
  End If
  '
  'format and export data
  '
  'Toolbox Plot Data come in two flavors: one X column that applies
  'to all Y columns (TPD.X1D=True), and a separate X column for each
  'Y column (TPD.X1D=False). For the first flavor, the exported format
  'goes like: X Y1 Y2 Y3 etc. For the second flavor, we must export
  'the data in sections, each with its own headers and X Y format.
  
  'determinte the number of "sections" that the output file will contain,
  'as well as the number of Y columns in each section
  If TPD.X1D Then
    nsect = 1
    nycol = TPD.NC
  Else
    nsect = TPD.NC
    nycol = 1
  End If
  
  'init warning flags. If these flags end up true, display
  'a warning message.
  warnheaders = False
  warnvalues = False
  
  For isect = 0 To nsect - 1  'loop over number of sections
    'headers
    If DoHeaders Then
      Print #1, "# "; PS.PlotTitle.Text
      'first column, X title
      buffer(0) = PS.XTitle.Text + PS.Xunits
      'Y titles
      For i = 0 To nycol - 1
        buffer(i + 1) = PS.YTitle.Text + PS.Yunits
      Next
      'build a single, formatted string for output from the individual buffers
      If dlm < 0 Then  'fixed width columns
        s = ""
        For i = 0 To nycol
          If Len(buffer(i)) > cols Then
            buffer(i) = Left$(buffer(i), cols) 'truncate header
            warnheaders = True  'set flag for warning display
          ElseIf Len(buffer(i)) < cols Then
            buffer(i) = Space$(cols - Len(buffer(i))) + buffer(i) 'pad header
          End If
          s = s + buffer(i)
        Next
      ElseIf dlm = 32 Then 'for a space delimiter, quote the headers
        s = Chr$(34) + buffer(0) + Chr$(34)
        For i = 0 To nycol
          s = s + Chr$(dlm) + Chr$(34) + buffer(i + 1) + Chr$(34)
        Next
      Else  'other single-character delimiter
        s = buffer(0)
        For i = 0 To nycol
          s = s + Chr$(dlm) + buffer(i + 1)
        Next
      End If
      Print #1, "# "; s
    End If
  
    'data
    If TPD.X1D Then
      np = TPD.np(0)
    Else
      np = TPD.np(isect)
    End If
    For ip = 0 To np - 1
      If TPD.X1D Then
        buffer(0) = Format$(TPD.X(ip))
        For i = 0 To nycol - 1
          buffer(i + 1) = Format$(TPD.Y(ip, i))
        Next
      Else
        buffer(0) = Format$(TPD.X(ip, isect))
        buffer(1) = Format$(TPD.Y(ip, isect))
      End If
      'build a single, formatted string for output from the individual buffers
      If dlm < 0 Then  'fixed width columns
        s = ""
        For i = 0 To nycol
          If Len(buffer(i)) > cols Then
            buffer(i) = String$(cols, "#") 'display overflow
            warnvalues = True  'set flag for warning display
          ElseIf Len(buffer(i)) < cols Then
            buffer(i) = Space$(cols - Len(buffer(i))) + buffer(i) 'pad header
          End If
          s = s + buffer(i)
        Next
      ElseIf dlm = 32 Then 'for a space delimiter, quote the headers
        s = Chr$(34) + buffer(0) + Chr$(34)
        For i = 0 To nycol
          s = s + Chr$(dlm) + Chr$(34) + buffer(i + 1) + Chr$(34)
        Next
      Else  'other single-character delimiter
        s = buffer(0)
        For i = 0 To nycol
          s = s + Chr$(dlm) + buffer(i + 1)
        Next
      End If
      Print #1, s
    Next
  Next
  
  'check for warnings
  If warnheaders Then
    s = "Warning: Some headers were too wide for the number of columns specified."
    MsgBox s, vbExclamation + vbOKOnly
  End If
  If warnvalues Then
    s = "Warning: Some values were too wide for the number of columns specified."
    MsgBox s, vbExclamation + vbOKOnly
  End If

  'success!
  Close #1
  ExportData = True
  Exit Function

ErrHandlerED:
  s = "Error writing file: " + fn + Chr$(13) + Error$(Err)
  t% = vbCritical + vbOKOnly
  MsgBox s, t%
  ExportData = False
  Exit Function
End Function

Private Sub Form_Load()
  InitForm
End Sub

Private Sub InitForm()
  CenterForm Me  'center the form
  txtNotes.Text = ""
  txtColumns.Text = "10"
End Sub


Private Sub txtDelim_Click()
  optDelim(3).Value = True
End Sub

Private Sub txtDelim_KeyPress(KeyAscii As Integer)
'Trap all ascii characters and place them in the text
'make the non-printing one printable
  Select Case KeyAscii
    Case 0 To 31  'nonprintables
      txtDelim.Text = "^" + Chr$(KeyAscii + 64)
      KeyAscii = 0
    Case 32 To 126
      txtDelim.Text = Chr$(KeyAscii)
      KeyAscii = 0
  End Select
End Sub

