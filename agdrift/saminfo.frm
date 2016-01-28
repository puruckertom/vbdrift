VERSION 4.00
Begin VB.Form frmSamsonInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library Site Information"
   ClientHeight    =   4785
   ClientLeft      =   1950
   ClientTop       =   4950
   ClientWidth     =   8745
   Height          =   5190
   Icon            =   "SAMINFO.frx":0000
   Left            =   1890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8745
   Top             =   4605
   Width           =   8865
   Begin VB.Frame fraInfo 
      Caption         =   "Info"
      Height          =   2295
      Left            =   5880
      TabIndex        =   2
      Top             =   0
      Width           =   2775
      Begin VB.Label lblInfoElevUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   1920
         TabIndex        =   32
         Top             =   1800
         Width           =   330
      End
      Begin VB.Label lblInfoElev 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   31
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblInfoElevLabel 
         AutoSize        =   -1  'True
         Caption         =   "Elevation"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label lblInfoLon 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblInfoLonLabel 
         AutoSize        =   -1  'True
         Caption         =   "Longitude"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label lblInfoLat 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblInfoLatLabel 
         AutoSize        =   -1  'True
         Caption         =   "Latitude"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label lblInfoID 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblInfoIDLabel 
         AutoSize        =   -1  'True
         Caption         =   "Site ID"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblInfoName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblInfoNameLabel 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.Frame fraMap 
      Caption         =   "Site Locator"
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5655
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2790
         Left            =   120
         MousePointer    =   2  'Cross
         Picture         =   "SAMINFO.frx":030A
         ScaleHeight     =   2790
         ScaleWidth      =   5340
         TabIndex        =   3
         Top             =   240
         Width           =   5340
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Distance"
         Height          =   195
         Left            =   3720
         TabIndex        =   21
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   195
         Left            =   2880
         TabIndex        =   20
         Top             =   3120
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Site"
         Height          =   195
         Left            =   1080
         TabIndex        =   19
         Top             =   3120
         Width           =   270
      End
      Begin VB.Label lblSiteDistance 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   3720
         TabIndex        =   18
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label lblSiteID 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   17
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label lblSiteName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   16
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label lblSiteDistance 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   15
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblSiteID 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   14
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblSiteName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   13
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblSiteDistance 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   12
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label lblSiteID 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   11
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label lblSiteName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   10
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label lblSiteDistance 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   9
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label lblSiteID 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   8
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label lblSiteName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   7
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label lblSiteDistance 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   6
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblSiteID 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   5
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblSiteName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   3360
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      HelpContextID   =   1200
      Left            =   7800
      TabIndex        =   0
      Top             =   4320
      Width           =   855
   End
End
Attribute VB_Name = "frmSamsonInfo"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadSiteInfo(SiteName As String)
'Get information about a site from the
'database and display it. This routine may be
'called from outside the form.
  Dim DB As Database
  Dim DS As Recordset
  
  If Not LibOpenMAADB(DB) Then Exit Sub
  Set DS = DB.OpenRecordset("WindRose", dbOpenDynaset)
  
  DS.FindFirst "Name='" & SiteName & "'"
  If DS.NoMatch Then
    'ClearInfoControls
    DS.Close
    DB.Close
    Exit Sub
  End If
  lblInfoName.Caption = DS.Fields("Name")
  lblInfoID.Caption = DS.Fields("SamsonID")
  lblInfoLat.Caption = DS.Fields("Latitude")
  lblInfoLon.Caption = DS.Fields("Longitude")
  'Unknown elevations are listed as "99999"
  If Val(DS.Fields("Elevation")) < 99999 Then
    lblInfoElev.Caption = Format$(Int(UnitsDisplay(Val(DS.Fields("Elevation")), UN_LENGTH)))
  Else
    lblInfoElev.Caption = "Unknown"
  End If
  
  DS.Close
  DB.Close
End Sub

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  'Center the form on the screen
  CenterForm Me

  lblInfoElevUnits.Caption = UnitsName(UN_LENGTH)
End Sub

