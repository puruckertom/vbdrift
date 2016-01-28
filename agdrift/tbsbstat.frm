VERSION 5.00
Begin VB.Form frmTBSBStats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spray Block Statistics"
   ClientHeight    =   3090
   ClientLeft      =   555
   ClientTop       =   3015
   ClientWidth     =   5010
   ForeColor       =   &H80000008&
   HelpContextID   =   1482
   Icon            =   "TBSBSTAT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3090
   ScaleWidth      =   5010
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      HelpContextID   =   1482
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calc"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1482
      Left            =   3240
      TabIndex        =   1
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Clos&e"
      Height          =   375
      HelpContextID   =   1482
      Left            =   4080
      TabIndex        =   0
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame fraDisp 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtCalc 
         Height          =   285
         HelpContextID   =   1536
         Index           =   2
         Left            =   2280
         TabIndex        =   5
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         HelpContextID   =   1260
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtCalc 
         Height          =   285
         HelpContextID   =   1350
         Index           =   0
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblEMD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mean Deposition:"
         Height          =   195
         Left            =   1005
         TabIndex        =   10
         Top             =   1845
         Width           =   1245
      End
      Begin VB.Label lblESW 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Effective Swath Width:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1125
         Width           =   1980
      End
      Begin VB.Label lblESWUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3840
         TabIndex        =   9
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label lblCOV 
         AutoSize        =   -1  'True
         Caption         =   "COV:"
         Height          =   195
         Left            =   1800
         TabIndex        =   7
         Top             =   405
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmTBSBStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: tbsbstat.frm,v 1.7 2001/05/24 20:16:26 tom Exp $
Option Explicit

Dim PropTakeAction As Integer
Dim NeedCalcs As Integer     'tracks calculation status
Dim CalcOutputMarker As Integer 'tracks output selection
Dim PreviousUnits As Integer 'Tracks units setting

Private Sub Calculate()
'Calculate the distances
  Dim COV As Single
  Dim ESW As Single
  Dim EMD As Single
  Dim INTYPE As Long

  ' Change the form mouse pointer
  Me.MousePointer = vbHourglass 'hourglass
  
  'extract input data from form controls
  COV = Val(txtCalc(0).Text)
  ESW = UnitsInternal(Val(txtCalc(1).Text), UN_LENGTH)
  EMD = Val(txtCalc(2).Text)
  INTYPE = CLng(CalcOutputMarker)
  
  Call agcov(CLng(UC.NumCOV), _
             UC.COVVal(0), UC.COVESW(0), UC.COVMVal(0), _
             INTYPE, COV, ESW, EMD)
  
  'stuff the results into the form controls
  'Place calculated values back in the controls
  PropTakeAction = False
  Select Case CalcOutputMarker
  Case 0
    If ESW >= 0 Then
      txtCalc(1) = AGFormat$(UnitsDisplay(ESW, UN_LENGTH))
    Else
      txtCalc(1) = "out of range!"
    End If
    If EMD >= 0 Then
      txtCalc(2) = AGFormat$(EMD)
    Else
      txtCalc(2) = "out of range!"
    End If
  Case 1
    If COV >= 0 Then
      txtCalc(0) = AGFormat$(COV)
    Else
      txtCalc(0) = "out of range!"
    End If
    If EMD >= 0 Then
      txtCalc(2) = AGFormat$(EMD)
    Else
      txtCalc(2) = "out of range!"
    End If
  Case 2
    If COV >= 0 Then
      txtCalc(0) = AGFormat$(COV)
    Else
      txtCalc(0) = "out of range!"
    End If
    If ESW >= 0 Then
      txtCalc(1) = AGFormat$(UnitsDisplay(ESW, UN_LENGTH))
    Else
      txtCalc(1) = "out of range!"
    End If
  End Select
  PropTakeAction = True
  
  NeedCalcs = False
  
  Me.MousePointer = vbDefault 'default
End Sub

Private Sub ClearOutputFields()
'clear all output fields
'Don't clear the one pointed to by CalcOutputMarker
  Dim PTAsave As Integer
  Dim i As Integer
  
  PTAsave = PropTakeAction
  PropTakeAction = False 'desensitize controls
  For i = 0 To 2
    If i = CalcOutputMarker Then
      txtCalc(i).ForeColor = vbRed
    Else
      txtCalc(i).Text = ""
      txtCalc(i).ForeColor = vbBlack
    End If
  Next
  PropTakeAction = PTAsave 'restore control sensitivity
End Sub

Private Sub cmdCalc_Click()
  If NeedCalcs Then Calculate
End Sub

Private Sub cmdOk_Click()
  'Reset calc flag, since we don't know if new
  'main calcs will be performed or loaded
  NeedCalcs = True
  ClearOutputFields
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
  Dim Mag As Variant

  If NeedCalcs Then Calculate
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
'This routine is executed each time the form is shown

  'If the units have changed since the last time this
  'form was shown, update a few things
  If UP.Units <> PreviousUnits Then
    'Update units lablels
    lblESWUnits.Caption = UnitsName(UN_LENGTH)
    
    'Convert any existing user-defined values
    '(if this is not the first time this form has been shown)
    If PreviousUnits <> -1 Then
      If CalcOutputMarker = 1 Then
        txtCalc(1).Text = _
          AGFormat$(UnitsDisplay(UnitsInternalSys(Val(txtCalc(1).Text), _
          UN_LENGTH, PreviousUnits), UN_LENGTH))
      End If
    End If
    PreviousUnits = UP.Units 'save the new setting
  End If
  Calculate
End Sub

Private Sub Form_Load()
'initialize this form
  'Center the form on the screen
  CenterForm Me
  
  PropTakeAction = True         'Activate control reactions
  txtCalc(0).Text = "0.3" 'COV
  PreviousUnits = -1
End Sub

Private Function GenFormData() As String
'Generate report text for this form to be used for printing
  
  Dim gfd As String  'temporary storage for report text
  Dim s As String        'workspace string

  gfd = "" 'start with a blank string
  
  AppendStr gfd, "AgDRIFT® Spray Block Statistics Toolbox", True
  AppendStr gfd, "", True
  
  AppendStr gfd, lblCOV.Caption & " " & txtCalc(0).Text, True
  AppendStr gfd, "", True
  AppendStr gfd, lblESW.Caption & " " & txtCalc(1).Text & " " & lblESWUnits.Caption, True
  AppendStr gfd, "", True
  AppendStr gfd, lblEMD.Caption & " " & txtCalc(2).Text, True
  AppendStr gfd, "", True
  
  AppendStr gfd, "Tier: " & String$(UD.Tier, "I"), True
  AppendStr gfd, "RunID:", True
  AppendStr gfd, "  " & GetRunID(), True
  AppendStr gfd, "", True
  
  GenFormData = gfd
End Function

Private Sub txtCalc_Change(Index As Integer)
'When this control changes, clear the other
'members of the array
  If PropTakeAction Then
    PropTakeAction = False
    If Trim$(txtCalc(Index).Text) = "" Then
      CalcOutputMarker = -1
    Else
      CalcOutputMarker = Index
    End If
    NeedCalcs = True
    ClearOutputFields
    PropTakeAction = True
  End If
End Sub
