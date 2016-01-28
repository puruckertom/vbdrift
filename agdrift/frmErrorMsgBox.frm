VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmErrorMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   6945
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmErrorMsgBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   165
      ScaleHeight     =   435
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   2625
      Visible         =   0   'False
      Width           =   360
   End
   Begin RichTextLib.RichTextBox rtbErrorSummary 
      Height          =   2055
      Left            =   975
      TabIndex        =   4
      Top             =   465
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmErrorMsgBox.frx":030A
   End
   Begin RichTextLib.RichTextBox rtbErrorDetails 
      Height          =   3390
      Left            =   975
      TabIndex        =   3
      Top             =   3360
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   5980
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmErrorMsgBox.frx":0385
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "View Details >>>"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   960
      X2              =   5880
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   975
      X2              =   5895
      Y1              =   3210
      Y2              =   3210
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Error:"
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "frmErrorMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'File: frmErrorMsgBox
'
'----------------------------------------------------------------------
'Re:
'    Display a message for user and also have available for tech support a detail listing of complete error an
'    and stack trace.
'
'----------------------------------------------------------------------
'
' Application defined Errors (vbObjectError + # )
' Number         Description
' ============   =================================================
'
'
'----------------------------------------------------------------------
' CONSTANTS:
'----------------------------------------------------------------------
'
'
'
'----------------------------------------------------------------------
' PRIVATE VARIABLES:
'----------------------------------------------------------------------
    Private dummyVar As Integer      'Not used, please remove when first var defined



Private Const VIEW_DETAILS As String = "View Details >>>"
Private Const HIDE_DETAILS As String = "<<< Hide Details"

'---------------------------------------------------------------------------
' cmdDetails_Click:
'    Either display detail error or hide it. This depends on the current state of from. This function should just toggle fr
'    from view to hide.
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2003-08-22  AED  Created
'
'---------------------------------------------------------------------------
Private Sub cmdDetails_Click()
    If cmdDetails.Caption = VIEW_DETAILS Then
        cmdDetails.Caption = HIDE_DETAILS
        Me.Height = 7500
        Me.Refresh
    Else
        cmdDetails.Caption = VIEW_DETAILS
        Me.Height = 3660
        Me.Refresh
    End If
End Sub

'---------------------------------------------------------------------------
' Form_Load:
'    Start dialog and initialially display only user message box (Hide details message box.)
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2003-08-22  AED  Created
'
'---------------------------------------------------------------------------
Private Sub Form_Load()
    On Error GoTo Error_Handler
    'Don't display detail box yet.
    cmdDetails.Caption = HIDE_DETAILS
    cmdDetails.Value = True
    


'====================================================
'Exit Point for Form_Load
'====================================================
Exit_Form_Load:
    Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
    'Err.Raise Err.Number, Err.Source, Err.Description
    MsgBox Err.Description, vbInformation, "Form_Load"
    Resume Exit_Form_Load

End Sub


'---------------------------------------------------------------------------
' OKButton_Click:
'    Exit this dialog box
'
' Modified:
' Date        Ini  Description
'===========  ===  =========================================================
' 2003-08-22  AED  Created
'
'---------------------------------------------------------------------------
Private Sub OKButton_Click()
    On Error GoTo Error_Handler
    Unload Me


'====================================================
'Exit Point for OKButton_Click
'====================================================
Exit_OKButton_Click:
    Exit Sub


'====================================================
'            ERROR HANDLER ROUTINE(S)
'====================================================
Error_Handler:
    'Err.Raise Err.Number, Err.Source, Err.Description
    MsgBox Err.Description, vbInformation, "OKButton_Click"
    Resume Exit_OKButton_Click

End Sub


