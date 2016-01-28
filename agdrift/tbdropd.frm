VERSION 5.00
Begin VB.Form frmTBDropDist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Drop Distance Calculator"
   ClientHeight    =   3450
   ClientLeft      =   5235
   ClientTop       =   1710
   ClientWidth     =   5010
   ForeColor       =   &H80000008&
   HelpContextID   =   1085
   Icon            =   "TBDROPD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3450
   ScaleWidth      =   5010
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      HelpContextID   =   1085
      Left            =   2400
      TabIndex        =   2
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calc"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1085
      Left            =   3240
      TabIndex        =   1
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Clos&e"
      Height          =   375
      HelpContextID   =   1085
      Left            =   4080
      TabIndex        =   0
      Top             =   3000
      Width           =   855
   End
   Begin VB.Frame fraDisp 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtRelHgt 
         Height          =   285
         HelpContextID   =   1060
         Left            =   2040
         TabIndex        =   4
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtDiam 
         Height          =   285
         HelpContextID   =   1363
         Left            =   2040
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         Caption         =   "Message"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label lblInput1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Release Height:"
         Height          =   195
         Left            =   600
         TabIndex        =   18
         Top             =   645
         Width           =   1380
      End
      Begin VB.Label lblRelHgtUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2970
         TabIndex        =   17
         Top             =   645
         Width           =   420
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Caption         =   "000000.00"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblUnits6 
         AutoSize        =   -1  'True
         Caption         =   "sec"
         Height          =   195
         Left            =   3600
         TabIndex        =   15
         Top             =   2280
         Width           =   315
      End
      Begin VB.Label lblResults2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Time to Impact:"
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lblResults0 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Drop Size at Impact:"
         Height          =   195
         Left            =   735
         TabIndex        =   8
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label lblUnits4 
         AutoSize        =   -1  'True
         Caption         =   "µm"
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   1590
         Width           =   255
      End
      Begin VB.Label lblResults1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Distance Traveled:"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblDistUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3600
         TabIndex        =   11
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label lblSize 
         Alignment       =   1  'Right Justify
         Caption         =   "000000.00"
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   1590
         Width           =   1215
      End
      Begin VB.Label lblDist 
         Alignment       =   1  'Right Justify
         Caption         =   "000000.00"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblUnits0 
         AutoSize        =   -1  'True
         Caption         =   "µm"
         Height          =   195
         Left            =   2970
         TabIndex        =   7
         Top             =   285
         Width           =   255
      End
      Begin VB.Label lblInput0 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Drop Size:"
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   285
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmTBDropDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tbdropd.frm,v 1.9 2001/08/13 17:40:04 tom Exp $

Private Sub Calculate()
'Calculate the distances
  Dim BgnSize As Single
  Dim RelHgt As Single
  Dim FinSize As Single
  Dim Dist As Single
  Dim TimeImpact As Single
  Dim WarnStr As String
  Dim WarnLen As Long

  ' Change the form mouse pointer
  Me.MousePointer = vbHourglass 'hourglass
  
  'extract input data from form controls
  BgnSize = Val(txtDiam.Text)
  RelHgt = UnitsInternal(Val(txtRelHgt.Text), UN_LENGTH)

  'Clear the Warning message
  lblMessage.Caption = ""

  WarnStr = Space$(40)
  Call agdrp(UD, BgnSize, RelHgt, FinSize, Dist, TimeImpact, WarnStr, WarnLen)
  
  'Check for a warning from agdrp.
  'If there is one, display it and recover the
  'possibly "adjusted" input params
  If WarnLen > 0 Then
    lblMessage.Caption = "Warning: Limits reached on " & Left$(WarnStr, WarnLen)
    txtDiam.Text = AGFormat$(BgnSize)
    txtRelHgt.Text = AGFormat$(UnitsDisplay(RelHgt, UN_LENGTH))
  End If

  'stuff the results into the form controls
  'If FinSize is -1, there was a problem
  If FinSize >= 0 Then
    lblSize.Caption = AGFormat$(FinSize)
    lblDist.Caption = AGFormat$(UnitsDisplay(Dist, UN_LENGTH))
    lblTime.Caption = AGFormat$(TimeImpact)
  Else
    lblSize.Caption = "out of range!"
    lblDist.Caption = ">" & AGFormat$(UnitsDisplay(Dist, UN_LENGTH))
    lblTime.Caption = "out of range!"
  End If

  Me.MousePointer = vbDefault 'default
End Sub

Private Sub ClearOutput()
'clear all output fields
  lblSize.Caption = ""
  lblDist.Caption = ""
  lblTime.Caption = ""
End Sub

Private Sub cmdCalc_Click()
  Calculate
End Sub

Private Sub cmdOk_Click()
  Hide
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

Private Sub Form_Activate()
  UpdateUnitsLabels
  Calculate
End Sub

Private Sub Form_Load()
  InitForm
End Sub

Private Function GenFormData() As String
'Generate report text for this form to be used for printing
  
  Dim gfd As String  'temporary storage for report text
  Dim s As String        'workspace string

  gfd = "" 'start with a blank string
  
  AppendStr gfd, "AgDRIFT® Drop Distance Toolbox", True
  AppendStr gfd, "", True
  
  AppendStr gfd, lblInput0.Caption & " " & txtDiam.Text & " " & lblUnits0.Caption, True
  AppendStr gfd, lblInput1.Caption & " " & txtRelHgt.Text & " " & lblRelHgtUnits.Caption, True
  AppendStr gfd, "", True
  AppendStr gfd, lblResults0.Caption & " " & lblSize.Caption & " " & lblUnits4.Caption, True
  AppendStr gfd, lblResults1.Caption & " " & lblDist.Caption & " " & lblDistUnits.Caption, True
  AppendStr gfd, lblResults2.Caption & " " & lblTime.Caption & " " & lblUnits6.Caption, True
  AppendStr gfd, "", True
  
  AppendStr gfd, "Tier: " & String$(UD.Tier, "I"), True
  AppendStr gfd, "RunID:", True
  AppendStr gfd, "  " & GetRunID(), True
  AppendStr gfd, "", True
  
  GenFormData = gfd
End Function

Private Sub InitForm()
'initialize this form
  'Center the form on the screen
  CenterForm Me
  
  'init controls
  lblMessage.Caption = ""
  txtDiam.Text = "1000"
  txtRelHgt.Text = AGFormat$(UnitsDisplay(UD.CTL.Height, UN_LENGTH))
End Sub

Private Sub txtDiam_Change()
  ClearOutput
End Sub

Private Sub txtRelHgt_Change()
  ClearOutput
End Sub

Private Sub UpdateUnitsLabels()
  lblRelHgtUnits.Caption = UnitsName(UN_LENGTH)
  lblDistUnits.Caption = UnitsName(UN_LENGTH)
End Sub

