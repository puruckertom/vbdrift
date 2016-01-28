VERSION 5.00
Begin VB.Form frmEasterEgg 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AgDRIFT Developers"
   ClientHeight    =   3750
   ClientLeft      =   2295
   ClientTop       =   2310
   ClientWidth     =   9480
   HelpContextID   =   1538
   Icon            =   "EASTER.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3750
   ScaleWidth      =   9480
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picPlane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      HelpContextID   =   1538
      Left            =   8280
      Picture         =   "EASTER.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   810
      TabIndex        =   1
      Top             =   1920
      Width           =   810
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Enough Already!"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1538
      Left            =   7800
      TabIndex        =   0
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblBanner 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "lblBanner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   7800
      TabIndex        =   2
      Top             =   2760
      Width           =   1005
   End
End
Attribute VB_Name = "frmEasterEgg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: easter.frm,v 1.7 2001/08/13 17:40:01 tom Exp $
Option Explicit

Private num_banners As Integer     'total number of banners
Private banner_status() As Integer '0=waiting 1=towed 2=falling 3=done
Private flight_phase As Integer    'flight phase of the airplane
Private towed_banner As Integer    'index of banner currently towed
Private banners_waiting As Boolean

'animation values
Private DELTA_X As Long   'forward motion step
Private DELTA_Y As Long   'falling step
Private X_DESCEND As Long 'transition from high cruise to descent
Private X_ASCEND As Long  'transition from low cruise to ascent
Private Y_HIGH As Long    'height of high cruise
Private Y_LOW As Long     'height of low cruise
  

Private Sub Form_Load()
  Dim i As Integer
  
  CenterForm Me
  picPlane.Left = Me.ScaleWidth 'start plane off the screen
  
  'initialize the banners
  num_banners = 20
  ReDim banner_status(num_banners - 1)
  For i = 0 To num_banners - 1
    If i > 0 Then Load lblBanner(i)
    lblBanner(i).Visible = True
    lblBanner(i).Left = Me.Width + 10
    banner_status(i) = 0
  Next
  i = 0
  lblBanner(i) = "Presenting the AgDRIFT® Development Team:": i = i + 1
  lblBanner(i) = "John W. Barry - USDA Forest Service (retired)": i = i + 1
  lblBanner(i) = "Sandra L. Bird - U. S. Environmental Protection Agency": i = i + 1
  lblBanner(i) = "L. F. Bouse - USDA Agricultural Research Service (retired)": i = i + 1
  lblBanner(i) = "Thomas B. Curbishley - Continuum Dynamics, Inc.": i = i + 1
  lblBanner(i) = "Robert B. Ekblad - USDA Forest Service (retired)": i = i + 1
  lblBanner(i) = "David M. Esterly - Environmental Focus, Inc.": i = i + 1
  lblBanner(i) = "David I. Gustafson - Monsanto Agricultural Company": i = i + 1
  lblBanner(i) = "Clarence G. Hermansky - DuPont Agricultural Products": i = i + 1
  lblBanner(i) = "Andrew J. Hewitt - Stewart Agricultural Research Services, Inc.": i = i + 1
  lblBanner(i) = "George G. Ice - National Council for Air and Stream Improvement, Inc.": i = i + 1
  lblBanner(i) = "I. W. Kirk - USDA Agricultural Research Service": i = i + 1
  lblBanner(i) = "Theodore C. Kuchnicki - PMRA Canada": i = i + 1
  lblBanner(i) = "James C. Lin - U. S. Environmental Protection Agency": i = i + 1
  lblBanner(i) = "Robert E. Mickle - REMSpC Spray Consulting": i = i + 1
  lblBanner(i) = "Steven G. Perry - U. S. Environmental Protection Agency": i = i + 1
  lblBanner(i) = "Scott L. Ray - Dow AgroSciences LLC": i = i + 1
  lblBanner(i) = "Milton E. Teske - Continuum Dynamics, Inc.": i = i + 1
  lblBanner(i) = "Harold W. Thistle - USDA Forest Service": i = i + 1
  lblBanner(i) = "David L. Valcore - Dow AgroSciences LLC": i = i + 1
  
  towed_banner = -1
  banners_waiting = True
End Sub

Private Sub Form_Resize()
  DELTA_X = -144    'forward increment (- for left, + for right)
  DELTA_Y = 15      'falling increment (+ for down, - for up)
  X_DESCEND = Me.ScaleWidth * 0.8  'change from high cruise to descent
  X_ASCEND = Me.ScaleWidth * 0.1   'change from low cruise to ascent
  Y_HIGH = 10      'height of high cruise
  Y_LOW = Me.ScaleHeight * 0.6 - picPlane.Height 'height of low cruise
End Sub

Private Sub Timer1_Timer()
'Do the next animation step
  Static Working As Boolean 'falg to prevent frame overlap
  
  Dim i As Integer
  Dim j As Integer
  
  If Working Then Exit Sub
  
  Working = True
  
  'Update the airplane
  Select Case flight_phase
  Case 0 'begin the flight pass
    picPlane.Left = Me.ScaleWidth
    picPlane.Top = Y_HIGH
    flight_phase = flight_phase + 1
      
    If banners_waiting Then
      towed_banner = towed_banner + 1 'pick up the next banner
      If towed_banner >= num_banners Then towed_banner = 0
      banner_status(towed_banner) = 1 'attach the banner to the airplane
    End If
  Case 1 'cruise high
    picPlane.Left = picPlane.Left + DELTA_X
    If picPlane.Left <= X_DESCEND Then
      flight_phase = flight_phase + 1
    End If
  Case 2 'descend
    picPlane.Left = picPlane.Left + DELTA_X
    picPlane.Top = picPlane.Top - DELTA_X
    If picPlane.Top >= Y_LOW Then
      flight_phase = flight_phase + 1
    End If
  Case 3 'cruise low
    picPlane.Left = picPlane.Left + DELTA_X
    If picPlane.Left <= X_ASCEND Then
      flight_phase = flight_phase + 1
    End If
  Case 4 'ascend
    picPlane.Left = picPlane.Left + DELTA_X
    picPlane.Top = picPlane.Top + DELTA_X
    If picPlane.Top <= Y_HIGH Then
      flight_phase = flight_phase + 1
    End If
    If picPlane.Left + picPlane.Width < 0 Then
      flight_phase = 0
    End If
  Case 5 'depart
    picPlane.Left = picPlane.Left + DELTA_X
    If picPlane.Left + picPlane.Width < 0 Then
      flight_phase = 0
    End If
  End Select
  
  'Banners
  banners_waiting = False
  For i = 0 To num_banners - 1
    With lblBanner(i)
      Select Case banner_status(i)
      Case 0 'waiting
        banners_waiting = True
      Case 1 'towing
        'follow the airplane
        .Left = picPlane.Left + picPlane.Width
        .Top = picPlane.Top + picPlane.Height * 0.5 - .Height * 0.5
        'follow the airplane until the name is centered
        If .Left + (.Width / 2) <= Me.ScaleWidth / 2 Then
          banner_status(i) = banner_status(i) + 1
        End If
        banners_waiting = False
      Case 2 'falling/rising
        'fall
        .Top = .Top + DELTA_Y
        'Stop at the ground
        If (.Top >= Me.ScaleHeight - .Height) Or _
           (.Top + .Height < 0) Then
          banner_status(i) = banner_status(i) + 1
        End If
        banners_waiting = False
      Case 3 'done
        'reset banner for the next time
        '.Left = Me.ScaleWidth + 10
        'if this is the last banner to be done, reset them all
        If i >= num_banners - 1 Then
          For j = 0 To num_banners - 1
            banner_status(j) = 0 'waiting
            lblBanner(j).Left = Me.ScaleWidth + 10
          Next
        End If
      End Select
    End With
  Next
  
  Working = False
End Sub

Private Sub cmdOk_Click()
  Unload Me
End Sub

