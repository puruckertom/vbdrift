VERSION 4.00
Begin VB.Form frmTier2Gnd 
   BorderStyle     =   0  'None
   Caption         =   "Tier II Ground Input"
   ClientHeight    =   5970
   ClientLeft      =   1335
   ClientTop       =   2100
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Height          =   6375
   HelpContextID   =   1280
   Icon            =   "TIER2GND.frx":0000
   Left            =   1275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7305
   Tag             =   "tier1"
   Top             =   1755
   Width           =   7425
   Begin VB.Frame Frame2 
      Caption         =   "Swath Range"
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   3375
      Begin VB.TextBox txtSwaths 
         Height          =   285
         HelpContextID   =   1280
         Left            =   1560
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         HelpContextID   =   1280
         Left            =   1560
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblSwaths 
         AutoSize        =   -1  'True
         Caption         =   "Starting Swath:"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ending Swath:"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   1035
      End
   End
   Begin VB.Frame fraDropSize 
      Caption         =   "Drop Size Distribution"
      Height          =   1575
      Left            =   3600
      TabIndex        =   9
      Top             =   840
      Width           =   3615
      Begin VB.OptionButton optDropSize 
         Caption         =   "&Very Coarse"
         Height          =   255
         HelpContextID   =   1280
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optDropSize 
         Caption         =   "&Coarse"
         Height          =   255
         HelpContextID   =   1280
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optDropSize 
         Caption         =   "&Medium (default)"
         Height          =   255
         HelpContextID   =   1280
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fraBoomHeight 
      Caption         =   "Boom Height"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   3375
      Begin VB.OptionButton optBoomHeight 
         Caption         =   "&Low Boom"
         Height          =   255
         HelpContextID   =   1280
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optBoomHeight 
         Caption         =   "&High Boom"
         Height          =   255
         HelpContextID   =   1280
         Index           =   1
         Left            =   1080
         TabIndex        =   7
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.PictureBox picLogo 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3960
      ScaleHeight     =   735
      ScaleWidth      =   3135
      TabIndex        =   3
      Top             =   4920
      Width           =   3135
      Begin VB.Label lblTM 
         AutoSize        =   -1  'True
         Caption         =   "TM"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   -1  'True
            strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblTier 
         AutoSize        =   -1  'True
         Caption         =   "Tier II"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   12
            underline       =   0   'False
            italic          =   -1  'True
            strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Line linLogo 
         BorderColor     =   &H000000FF&
         BorderWidth     =   4
         X1              =   720
         X2              =   2400
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblLogo 
         Caption         =   "AgDRIFT"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   24
            underline       =   0   'False
            italic          =   -1  'True
            strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame fraRunTitle 
      Caption         =   "Title"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtRunTitle 
         Alignment       =   2  'Center
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   1300
         Left            =   120
         TabIndex        =   0
         Text            =   "Untitled"
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmTier2Gnd"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' $Id: tier2gnd.frm,v 1.1 2000/03/06 21:29:24 tom Exp $

Private Sub DataToForm()
'Places user data values in form controls
' File name
  UpdateInputFormCaption
' Title
  txtRunTitle.Text = UD.Title               'Title
' Application Method
'  If UD.ApplMethod = 0 Then 'Aerial
'    optDropSize(UD.DSD(0).BasicType) = True 'Drop Distribution
'    For i = 0 To 4
'      optBoomHeight(i) = False
'    Next
'  Else
'    optBoomHeight(UD.GA.BasicType) = True
'    For i = 0 To 3
'      optDropSize(i) = False
'    Next
'  End If
'' Number of swaths for ground sprayer
'  txtSwaths.Text = UD.GA.NumSwaths
End Sub

Private Sub Form_Load()
  InitForm  'Initialize the form objects
End Sub

Private Sub Form_Resize()
  ResizeForm
End Sub

Private Sub InitForm()
'Initialize the data entry form
  Dim SaveDC As Integer
  
  'Save the current state of DataChanged.
  'We need to do this because by loading a new form
  'and updating its controls, DataChanged will be
  'set.
  SaveDC = UI.DataChanged
  
  'Transfer User data to form controls
  DataToForm
  
  'Make sure the drop distribution is set
  If UD.ApplMethod = 0 Then 'Aerial
    GetBasicDataDSD UD.DSD(0).BasicType, UD.DSD(0)
  End If
  
  UpdateDataChangedFlag SaveDC 'restore DataChanged

End Sub

Private Sub lblLogo_Click()
  frmAbout.Show 1
End Sub

Private Sub optBoomHeight_Click(Index As Integer)
  If (Index <> UD.GA.BasicType) Or UD.ApplMethod = 0 Then
    If Index <> UD.GA.BasicType Then
      optBoomHeight(UD.GA.BasicType) = False 'turn off old one
    End If
    UD.GA.BasicType = Index
    UD.ApplMethod = 1
    optDropSize(UD.DSD(0).BasicType) = False 'turn off aerial
    UpdateSwathControls
    LoadTier1Data UD, UC
    UpdateDataChangedFlag True 'Data was changed
  End If
End Sub

Private Sub optDropSize_Click(Index As Integer)
  If (Index <> UD.DSD(0).BasicType) Or UD.ApplMethod = 1 Then
    UD.DSD(0).BasicType = Index
    UD.ApplMethod = 0
    optBoomHeight(UD.GA.BasicType) = False 'turn off ground
    UpdateSwathControls
    LoadTier1Data UD, UC
    UpdateDataChangedFlag True 'Data was changed
  End If
End Sub

Private Sub ResizeForm()
'relocate controls when the form is resized
  'position agdrift logo
  'the top must not go above the DSD frame
  toplimit = fraDropSize.Top + fraDropSize.Height + 300
  leftlimit = 300
  logotop = Me.ScaleHeight - picLogo.Height - 300
  logoleft = Me.ScaleWidth - picLogo.Width - 300
  If logotop < toplimit Then logotop = toplimit
  If logoleft < leftlimit Then logoleft = leftlimit
  picLogo.Top = logotop
  picLogo.Left = logoleft
  
  'position the title frame and text box
  'it must not get narrower than the Orchard frame
  widlimit = fraDropSize.Left + fraDropSize.Width
  titlewidth = Me.ScaleWidth - fraRunTitle.Left - 100
  If titlewidth < widlimit Then titlewidth = widlimit
  fraRunTitle.Width = titlewidth
  'text box
  txtRunTitle.Width = fraRunTitle.Width - txtRunTitle.Left - 120
End Sub

Private Sub txtRunTitle_Change()
  UD.Title = txtRunTitle.Text
  UpdateDataChangedFlag True 'Data was changed
End Sub



Private Sub txtSwaths_Change()
'update the value of NumSwaths, but don't reload the
'Tier 1 data yet, since the user may not be finished typing.
'Do the reloading of data on the LostFocus method
  UD.GA.NumSwaths = Val(txtSwaths.Text)
  LoadTier1Data UD, UC
  UpdateDataChangedFlag True 'Data was changed
End Sub


Private Sub txtSwaths_LostFocus()
'update the control in case it was changed
'during the last LoadTier1Data
  txtSwaths.Text = AGFormat$(UD.GA.NumSwaths)
End Sub



Private Sub UpdateSwathControls()
'Update the state of the Number of Swaths controls
  If ((UD.GA.BasicType <= 1) And (UD.ApplMethod = 1)) Then
    lblSwaths.Enabled = True
    txtSwaths.Enabled = True 'Turn on the swaths control
  Else
    lblSwaths.Enabled = False
    txtSwaths.Enabled = False 'Turn off the swaths control
  End If
End Sub
