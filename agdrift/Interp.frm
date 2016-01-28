VERSION 5.00
Begin VB.Form frmInterpolateDSD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interpolation Method"
   ClientHeight    =   1785
   ClientLeft      =   3615
   ClientTop       =   2385
   ClientWidth     =   3000
   HelpContextID   =   1174
   Icon            =   "INTERP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1785
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optMethod 
      Caption         =   "&Root Normal"
      Height          =   255
      HelpContextID   =   1174
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton optMethod 
      Caption         =   "Rosin-R&ammler"
      Height          =   255
      HelpContextID   =   1174
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton optMethod 
      Caption         =   "&Log Normal"
      Height          =   255
      HelpContextID   =   1174
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1174
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1174
      Left            =   2040
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "frmInterpolateDSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: Interp.frm,v 1.5 2001/04/26 16:20:54 tom Exp $

Private Sub cmdCancel_Click()
  Me.Hide
End Sub


Private Sub cmdOk_Click()
  'find the current method selection and
  'return it in the tag
  For i = 0 To 2
    If optMethod(i).Value Then
      Me.Tag = Format$(i)
    End If
  Next
  Me.Hide
End Sub

Private Sub Form_Load()
  InitForm
End Sub

Public Sub InitForm()
  CenterForm Me
  Me.Tag = "-1" 'default return value
End Sub
