VERSION 2.00
Begin Form Form1 
   BorderStyle     =   3  'Fixed Double
   Caption         =   "Line Style"
   ClientHeight    =   3840
   ClientLeft      =   3420
   ClientTop       =   2640
   ClientWidth     =   2880
   Height          =   4245
   Left            =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   2880
   Top             =   2295
   Width           =   3000
   Begin CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin Shape shpColorSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   7
      Left            =   840
      Top             =   2040
      Width           =   375
   End
   Begin Shape shpColor 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   7
      Left            =   840
      Top             =   2040
      Width           =   375
   End
   Begin Shape shpColorSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   6
      Left            =   840
      Top             =   1560
      Width           =   375
   End
   Begin Shape shpColor 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   6
      Left            =   840
      Top             =   1560
      Width           =   375
   End
   Begin Shape shpColorSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   5
      Left            =   840
      Top             =   1080
      Width           =   375
   End
   Begin Shape shpColor 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   840
      Top             =   1080
      Width           =   375
   End
   Begin Line Line1 
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   4
      X1              =   1680
      X2              =   2520
      Y1              =   2700
      Y2              =   2700
   End
   Begin Shape shpColorSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   4
      Left            =   840
      Top             =   600
      Width           =   375
   End
   Begin Shape shpStyleSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   4
      Left            =   1560
      Top             =   2520
      Width           =   1095
   End
   Begin Shape shpColor 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   840
      Top             =   600
      Width           =   375
   End
   Begin Line Line1 
      BorderStyle     =   4  'Dash-Dot
      Index           =   3
      X1              =   1680
      X2              =   2520
      Y1              =   2220
      Y2              =   2220
   End
   Begin Shape shpColorSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   3
      Left            =   240
      Top             =   2040
      Width           =   375
   End
   Begin Shape shpStyleSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   3
      Left            =   1560
      Top             =   2040
      Width           =   1095
   End
   Begin Shape shpColor 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   240
      Top             =   2040
      Width           =   375
   End
   Begin Line Line1 
      BorderStyle     =   3  'Dot
      Index           =   2
      X1              =   1680
      X2              =   2520
      Y1              =   1740
      Y2              =   1740
   End
   Begin Shape shpColorSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   2
      Left            =   240
      Top             =   1560
      Width           =   375
   End
   Begin Shape shpStyleSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   2
      Left            =   1560
      Top             =   1560
      Width           =   1095
   End
   Begin Shape shpColor 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   240
      Top             =   1560
      Width           =   375
   End
   Begin Line Line1 
      BorderStyle     =   2  'Dash
      Index           =   1
      X1              =   1680
      X2              =   2520
      Y1              =   1260
      Y2              =   1260
   End
   Begin Shape shpColorSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   1
      Left            =   240
      Top             =   1080
      Width           =   375
   End
   Begin Shape shpStyleSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   1
      Left            =   1560
      Top             =   1080
      Width           =   1095
   End
   Begin Shape shpColor 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   240
      Top             =   1080
      Width           =   375
   End
   Begin Line Line1 
      Index           =   0
      X1              =   1680
      X2              =   2520
      Y1              =   780
      Y2              =   780
   End
   Begin Shape shpColorSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   0
      Left            =   240
      Top             =   600
      Width           =   375
   End
   Begin Label lblStyle 
      AutoSize        =   -1  'True
      Caption         =   "Line Style"
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin Label lblColor 
      AutoSize        =   -1  'True
      Caption         =   "Line Color"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   870
   End
   Begin Shape shpStyleSelect 
      BorderWidth     =   3
      Height          =   375
      Index           =   0
      Left            =   1560
      Top             =   600
      Width           =   1095
   End
   Begin Shape shpColor 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   240
      Top             =   600
      Width           =   375
   End
End
