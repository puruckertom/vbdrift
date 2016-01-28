VERSION 5.00
Begin VB.Form frmPrintPreview 
   AutoRedraw      =   -1  'True
   Caption         =   "Print Preview"
   ClientHeight    =   5115
   ClientLeft      =   1590
   ClientTop       =   1545
   ClientWidth     =   5190
   ForeColor       =   &H80000008&
   HelpContextID   =   1541
   Icon            =   "PRTPREVW.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5115
   ScaleWidth      =   5190
   Begin VB.PictureBox picView 
      BorderStyle     =   0  'None
      Height          =   3855
      HelpContextID   =   1541
      Left            =   0
      ScaleHeight     =   3855
      ScaleWidth      =   4695
      TabIndex        =   7
      Top             =   720
      Width           =   4695
      Begin VB.PictureBox picPage 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         HelpContextID   =   1541
         Index           =   1
         Left            =   0
         ScaleHeight     =   7.104
         ScaleMode       =   0  'User
         ScaleWidth      =   9.19
         TabIndex        =   8
         Top             =   0
         Width           =   4455
      End
   End
   Begin VB.HScrollBar hscPic 
      Height          =   255
      HelpContextID   =   1541
      LargeChange     =   1440
      Left            =   0
      SmallChange     =   360
      TabIndex        =   5
      Top             =   4680
      Width           =   4815
   End
   Begin VB.VScrollBar vscPic 
      Height          =   3975
      HelpContextID   =   1541
      LargeChange     =   1440
      Left            =   4800
      SmallChange     =   360
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox picCmd 
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   4995
      TabIndex        =   6
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton cmdZoom 
         Caption         =   "&Zoom"
         Height          =   390
         HelpContextID   =   1541
         Left            =   3000
         TabIndex        =   3
         Top             =   120
         Width           =   840
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   390
         HelpContextID   =   1541
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   840
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "&Prev"
         Height          =   390
         HelpContextID   =   1541
         Left            =   1080
         TabIndex        =   1
         Top             =   120
         Width           =   840
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   390
         HelpContextID   =   1541
         Left            =   2040
         TabIndex        =   2
         Top             =   120
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: prtprevw.frm,v 1.5 2001/08/30 14:00:34 tom Exp $
'This form is used to preview printout that is
'to be formatted with PrintData. The raw text is
'passwd to this form via the Tag property
'
'Me.Tag  i  raw text to be formatted
'
Dim pages As Integer           'total number of pages
Dim Mag As Single              'Magnification
Dim Zoom As Integer            'Zoom flag
Dim CurrentPageNum As Integer  'current page number
Dim CurrentPage As Control     'current page picturebox

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdNext_Click()
If CurrentPageNum < pages Then
  CurrentPageNum = CurrentPageNum + 1
  ShowPage
End If
End Sub

Private Sub cmdPrev_Click()
If CurrentPageNum > 1 Then
  CurrentPageNum = CurrentPageNum - 1
  ShowPage
End If
End Sub

Private Sub cmdZoom_Click()
  ZoomToggle
End Sub

Private Sub Form_Activate()
  GetFullPageMag
  GenPages
End Sub

Private Sub Form_Load()
  InitForm
End Sub

Private Sub Form_Resize()
  ResizeForm
End Sub

Private Sub GenPages()
'generate the print preview pages
  
  'delete preveously generated pages
  ResetPages
  
  'print the pages to the picturebox array
  PrintData CStr(Me.Tag), True, pages, Mag
  
  'show the new current page and set up the scrollbars
  ShowPage
  SetScrollbars
  PositionPages
End Sub

Private Sub GetFullPageMag()
'figure out the magnification for full-page
'display and stuff it in the combo box
  Dim Magv As Single
  Dim Magh As Single

  Magv = picView.Height / Printer.Height
  Magh = picView.Width / Printer.Width

  If Magv < Magh Then
    Mag = Magv
  Else
    Mag = Magh
  End If
End Sub

Private Sub hscPic_Change()
'scroll all pages
  PositionPages
End Sub

Private Sub hscPic_Scroll()
'scroll all pages
  PositionPages
End Sub

Private Sub InitForm()
'initialize this form
  
  'init current page variables
  Set CurrentPage = picPage(1)
  pages = 1
  CurrentPageNum = 1
  
  'init zoom amd mag
  Zoom = False
  Mag = 1

  'grow the form to full screen
  Me.Left = 0
  Me.Top = 0
  Me.Width = Screen.Width
  Me.Height = Screen.Height

End Sub

Private Sub picPage_Click(Index As Integer)
  ZoomToggle
End Sub

Private Sub PositionPages()
  'center the pages, if narrow
  For i = 1 To pages
    If Not hscPic.Enabled Then
      picPage(i).Left = hscPic.Width / 2 - picPage(i).Width / 2
    Else
      picPage(i).Left = -hscPic.Value
    End If
    If Not vscPic.Enabled Then
      picPage(i).Top = vscPic.Height / 2 - picPage(i).Height / 2
    Else
      picPage(i).Top = -vscPic.Value
    End If
  Next
End Sub

Private Sub ResetPages()
  'clear page 1, but don't remove
  picPage(1).Cls
  'unload additional pages
  For i = 2 To pages
    Unload picPage(i)
  Next
  'reset the Current page
  Set CurrentPage = picPage(1) 'set the current page control
End Sub

Private Sub ResizeForm()
'Adjust form controls to fit the current form size
  
  'controls Picture box
  picCmd.Top = 0
  picCmd.Left = 0
  picCmd.Width = Me.ScaleWidth
  
  'scroll bars
  vscPic.Top = picCmd.Top + picCmd.Height
  vscPic.Left = Me.ScaleWidth - vscPic.Width
  vscPic.Height = Me.ScaleHeight - picCmd.Height - hscPic.Height
  '
  hscPic.Top = Me.ScaleHeight - hscPic.Height
  hscPic.Left = 0
  hscPic.Width = Me.ScaleWidth - vscPic.Width

  'viewport picture box
  picView.Top = picCmd.Height
  picView.Left = 0
  picView.Height = Me.ScaleHeight - picCmd.Height - hscPic.Height
  picView.Width = Me.ScaleWidth - vscPic.Width

  'set up Scroll bars
  SetScrollbars
  PositionPages

End Sub

Private Sub SetScrollbars()
'reconfigure the scrollbars according to current data

  ' Set the Max value for the scroll bars.
  hscPic.Max = CurrentPage.Width - picView.Width
  vscPic.Max = CurrentPage.Height - picView.Height
  
  ' Determine if child picture will fill up screen.
  ' If so, then there is no need to use scroll bars.
  vscPic.Enabled = (picView.Height < CurrentPage.Height)
  hscPic.Enabled = (picView.Width < CurrentPage.Width)
End Sub

Private Sub ShowPage()
'show the current page by making its control visible
'and the control of the previous page invisible
  CurrentPage.Visible = False
  Set CurrentPage = picPage(CurrentPageNum)
  CurrentPage.Visible = True
End Sub

Private Sub vscPic_Change()
'scroll all pages
  PositionPages
End Sub

Private Sub vscPic_Scroll()
'scroll all pages
  PositionPages

End Sub

Private Sub ZoomToggle()
  Static Zoom As Integer
  If Zoom Then
    GetFullPageMag
    GenPages
    Zoom = False
  Else
    Mag = 1
    GenPages
    Zoom = True
  End If
End Sub

