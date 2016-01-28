VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlotEditTitle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Title"
   ClientHeight    =   1050
   ClientLeft      =   1425
   ClientTop       =   3195
   ClientWidth     =   6480
   ForeColor       =   &H80000008&
   Icon            =   "PLOTTEDT.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1050
   ScaleWidth      =   6480
   Begin VB.CommandButton cmdFont 
      Caption         =   "&Font..."
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtTitle 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPreview 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmPlotEditTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: plottedt.frm,v 1.4 2001/04/26 16:21:59 tom Exp $
'NAME:
'frmPlotEditTitle
'
'PURPOSE:
'label editor.
'This form allows the text and font properties of
'a label control to be edited.
'
'INPUT:
'lblData.Caption
'lblData.FontName
'lblData.FontSize
'lblData.FontBold
'lblData.FontItalic
'lblData.FontUnderline
'lblData.FontStrikethru
'lblData.ForeColor
'
'OUTPUT:
'lblData.Caption
'lblData.FontName
'lblData.FontSize
'lblData.FontBold
'lblData.FontItalic
'lblData.FontUnderline
'lblData.FontStrikethru
'lblData.ForeColor
'
'Tag property returns status true=OK false=cancel
'
'USAGE:
'The user modifies the properties of the lblData label
'control of this form and shows the form as modal. Upon
'return the updated values are unloaded from the lblData
'label control and copied to the desired target control.
'
'EXAMPLE:
'  'Create an instance of the form
'  Dim titleform As New frmPlotEditTitle
'  titleform.Caption = "Title"   'Title the window
'  'Load the text attributes into the form's data control
'  titleform!lblData.Caption = lblTitle.Caption
'  titleform!lblData.FontName = lblTitle.FontName
'  titleform!lblData.FontSize = lblTitle.FontSize
'  titleform!lblData.FontBold = lblTitle.FontBold
'  titleform!lblData.FontItalic = lblTitle.FontItalic
'  titleform!lblData.FontUnderline = lblTitle.FontUnderline
'  titleform!lblData.FontStrikethru = lblTitle.FontStrikethru
'  titleform!lblData.ForeColor = lblTitle.ForeColor
'  'Show the form as modal to collect the changes
'  titleform.show vbmodal
'  'Store the new parameters in the title label
'  lblTitle.Caption = titleform!lblData.Caption
'  lblTitle.FontName = titleform!lblData.FontName
'  lblTitle.FontSize = titleform!lblData.FontSize
'  lblTitle.FontBold = titleform!lblData.FontBold
'  lblTitle.FontItalic = titleform!lblData.FontItalic
'  lblTitle.FontUnderline = titleform!lblData.FontUnderline
'  lblTitle.FontStrikethru = titleform!lblData.FontStrikethru
'  lblTitle.ForeColor = titleform!lblData.ForeColor
'  'Dump the form
'  Unload titleform
'----------------------------------------------------------
Dim DataSaved As Integer 'True if lblData was copied to lblSave

Private Sub cmdCancel_Click()
  'Restore the data to its original state and hide the form
  DataRestore
  Me.Tag = "False"
  Hide
End Sub

Private Sub cmdFont_Click()
  'Load up Dialog Selections from Data label on form
  CMDialog1.FontName = lblData.FontName
  CMDialog1.FontSize = lblData.FontSize
  CMDialog1.FontBold = lblData.FontBold
  CMDialog1.FontItalic = lblData.FontItalic
  CMDialog1.FontUnderline = lblData.FontUnderline
  CMDialog1.FontStrikethru = lblData.FontStrikethru
  CMDialog1.Color = lblData.ForeColor

  On Error GoTo ErrHandlerCancel
  'Set cancel to true
  CMDialog1.CancelError = True
  'Set the cdlCFBoth and cdlCFEffects flags
  CMDialog1.Flags = cdlCFBoth Or cdlCFEffects
  'Display the Font dialog box
  CMDialog1.Action = 4
  'Set text properties according to user's selections
  On Error GoTo ErrHandlerFont
  lblData.FontName = CMDialog1.FontName
  lblData.FontSize = CMDialog1.FontSize
  lblData.FontBold = CMDialog1.FontBold
  lblData.FontItalic = CMDialog1.FontItalic
  lblData.FontUnderline = CMDialog1.FontUnderline
  lblData.FontStrikethru = CMDialog1.FontStrikethru
  lblData.ForeColor = CMDialog1.Color
  Exit Sub

ErrHandlerCancel:
  'User pressed Cancel
  Exit Sub

ErrHandlerFont:
  'Error setting the font
  m$ = "Could not set display font to " + Chr$(34) + CMDialog1.FontName + Chr$(34) + "." + Chr$(13)
  m$ = m$ + "Try selecting a TrueType font instead."
  t% = vbExclamation + vbOKOnly
  MsgBox m$, t%
  Resume Next
End Sub

Private Sub cmdOk_Click()
  Me.Tag = "True"
  Hide
End Sub

Private Sub DataRestore()
  'Copy the data stored in the lblSave control back into
  'the lblData control. Undoes any changes.
  lblData.Caption = lblSave.Caption
  lblData.FontName = lblSave.FontName
  lblData.FontSize = lblSave.FontSize
  lblData.FontBold = lblSave.FontBold
  lblData.FontItalic = lblSave.FontItalic
  lblData.FontUnderline = lblSave.FontUnderline
  lblData.FontStrikethru = lblSave.FontStrikethru
  lblData.ForeColor = lblSave.ForeColor
End Sub

Private Sub DataSave()
  'Save lblData data in lblSave control in case it's
  'needed later.
  lblSave.Caption = lblData.Caption
  lblSave.FontName = lblData.FontName
  lblSave.FontSize = lblData.FontSize
  lblSave.FontBold = lblData.FontBold
  lblSave.FontItalic = lblData.FontItalic
  lblSave.FontUnderline = lblData.FontUnderline
  lblSave.FontStrikethru = lblData.FontStrikethru
  lblSave.ForeColor = lblData.ForeColor
End Sub

Private Sub Form_Load()
  'Center the form on the screen
  Me.Left = (Screen.Width / 2) - (Me.Width / 2)
  Me.Top = (Screen.Height / 2) - (Me.Height / 2)
  'set default return value
  Me.Tag = "False"
End Sub

Private Sub Form_Paint()
'This code is located here, rather than Form_Load because
'The controls have not been set up until after load time.
  'Save the input data in the lblSave control in case
  'the user presses the cancel button later.
  If (Not DataSaved) Then
    DataSave
    DataSaved = True
  End If
  'Select the entire text in txtTitle
  txtTitle.SelStart = 0
  txtTitle.SelLength = Len(txtTitle.Text)
End Sub

Private Sub lblData_Change()
  'synchronize text control caption with Data control
  'see also txtTitle_Change
  txtTitle.Text = lblData.Caption
End Sub

Private Sub txtTitle_Change()
  'synchronize text control caption with Data control
  'see also lblData_Change
  lblData.Caption = txtTitle.Text
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
'Allow carriage returns as string input
  If KeyAscii = 13 Then
    txtTitle.SelText = Chr$(KeyAscii) 'insert the character
    KeyAscii = 0                      'throw away the key
  End If
End Sub

