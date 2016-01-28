VERSION 5.00
Begin VB.Form frmAircraftLibUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aircraft User Library"
   ClientHeight    =   4665
   ClientLeft      =   2955
   ClientTop       =   2040
   ClientWidth     =   5985
   ForeColor       =   &H80000008&
   HelpContextID   =   1462
   Icon            =   "frmAircraftLibUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4665
   ScaleWidth      =   5985
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Entry"
      Height          =   375
      HelpContextID   =   1462
      Left            =   120
      TabIndex        =   48
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1462
      Left            =   5040
      TabIndex        =   1
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1462
      Left            =   4080
      TabIndex        =   0
      Top             =   4200
      Width           =   855
   End
   Begin VB.Frame fraName 
      Caption         =   "Name"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   5775
      Begin VB.ComboBox cboName 
         Height          =   315
         HelpContextID   =   1023
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Properties"
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   5775
      Begin VB.Label lblBoomFwdlbl 
         AutoSize        =   -1  'True
         Caption         =   "Boom Fwd:"
         Height          =   195
         Left            =   3000
         TabIndex        =   47
         Top             =   2880
         Width           =   795
      End
      Begin VB.Label lblBoomFwd 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BoomFwd"
         Height          =   255
         Left            =   4320
         TabIndex        =   46
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblBoomFwdUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   5400
         TabIndex        =   45
         Top             =   2880
         Width           =   420
      End
      Begin VB.Label lblBoomVertlbl 
         AutoSize        =   -1  'True
         Caption         =   "Boom Vert:"
         Height          =   195
         Left            =   3000
         TabIndex        =   44
         Top             =   2520
         Width           =   780
      End
      Begin VB.Label lblBoomVert 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BoomVert"
         Height          =   255
         Left            =   4320
         TabIndex        =   43
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblBoomVertUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   5400
         TabIndex        =   42
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label lblWingVertlbl 
         AutoSize        =   -1  'True
         Caption         =   "Wing Vert:"
         Height          =   195
         Left            =   3000
         TabIndex        =   41
         Top             =   2160
         Width           =   750
      End
      Begin VB.Label lblWingVert 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WingVert"
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblWingVertUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   5400
         TabIndex        =   39
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label lblEngHoriz 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EngHoriz"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   38
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblEngHorizlbl 
         AutoSize        =   -1  'True
         Caption         =   "Engine Horiz.:"
         Height          =   195
         Left            =   3000
         TabIndex        =   37
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label lblEngHoriz 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EngHoriz"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   36
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblEngHorizUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   5400
         TabIndex        =   35
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label lblNumEnglbl 
         AutoSize        =   -1  'True
         Caption         =   "Engines.:"
         Height          =   195
         Left            =   3000
         TabIndex        =   34
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblNumEng 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NumEng"
         Height          =   255
         Left            =   4320
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblBiplSepUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2520
         TabIndex        =   25
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label lblPlanAreaUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2520
         TabIndex        =   32
         Top             =   2880
         Width           =   420
      End
      Begin VB.Label lblEngVertUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   5400
         TabIndex        =   31
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblEngFwdUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   5400
         TabIndex        =   30
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label lblPropRadUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2520
         TabIndex        =   29
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label lblTypSpeedUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2520
         TabIndex        =   28
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label lblWeightUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2520
         TabIndex        =   27
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label lblSemiSpanUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   2520
         TabIndex        =   26
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblEngFwd 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EngFwd"
         Height          =   255
         Left            =   4320
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblEngVert 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EngVert"
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblPropRad 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PropRad"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblPropRPM 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PropRPM"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblPlanArea 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PlanArea"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblWeight 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Weight"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblBiplSep 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BiplSep"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblTypSpeed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TypSpeed"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblSemiSpan 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SemiSpan"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblWingType 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WingType"
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblEngFwdlbl 
         AutoSize        =   -1  'True
         Caption         =   "Engine Fwd.:"
         Height          =   195
         Left            =   3000
         TabIndex        =   15
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label lblEngVertlbl 
         AutoSize        =   -1  'True
         Caption         =   "Engine Vert.:"
         Height          =   195
         Left            =   3000
         TabIndex        =   16
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label lblPropRadlbl 
         AutoSize        =   -1  'True
         Caption         =   "Prop Rad.:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label lblPropRPMlbl 
         AutoSize        =   -1  'True
         Caption         =   "Propeller RPM:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label lblPlanArealbl 
         AutoSize        =   -1  'True
         Caption         =   "Planform Area:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   1260
      End
      Begin VB.Label lblWeightlbl 
         AutoSize        =   -1  'True
         Caption         =   "Weight:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lblBiplSeplbl 
         AutoSize        =   -1  'True
         Caption         =   "Biplane Sep.:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label lblTypSpeedlbl 
         AutoSize        =   -1  'True
         Caption         =   "Typ. Speed:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label lblSemiSpanlbl 
         AutoSize        =   -1  'True
         Caption         =   "Semispan:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblWingTypelbl 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmAircraftLibUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: frmAircraftLibUser.frm,v 1.2 2001/05/24 20:16:20 tom Exp $

'This form interacts with the Aircraft User Library. When this form
'is loaded it fills a ComboBox with all of the entries found in the
'Aircraft Table in the User Library. If an entry exists which matches
'the current aircraft name, it is selected and its properties are
'displayed. Otherwise the first entry is displayed.

'Note that the following members of AircraftData are not
'affected by this form:
' BasicType
' PropEff
' DragCoeff

Public OK As Boolean       'return status: True if OK button pressed
Public ACName As String    'name of selected aircraft

Private AC As AircraftData 'local copy of aircraft data

Private colFixedControls As Collection 'Fixed-wing-only controls

Public Sub SelectEntry(EntryName As String)
'Try to select the supplied entry in the combo
  Dim i As Integer
  
  For i = 0 To cboName.ListCount - 1
    If Trim$(cboName.List(i)) = Trim$(EntryName) Then
      cboName.ListIndex = i
      Exit For
    End If
  Next
End Sub

Private Sub ClearPropertyControls()
  lblWingType.Caption = ""
  lblSemiSpan.Caption = ""
  lblTypSpeed.Caption = ""
  lblBiplSep.Caption = ""
  lblWeight.Caption = ""
  lblPlanArea.Caption = ""
  lblPropRPM.Caption = ""
  lblPropRad.Caption = ""
  lblEngVert.Caption = ""
  lblEngFwd.Caption = ""
  lblNumEng.Caption = ""
  lblEngHoriz(0).Caption = ""
  lblEngHoriz(1).Caption = ""
  lblWingVert.Caption = ""
  lblBoomVert.Caption = ""
  lblBoomFwd.Caption = ""
    
  lblSemiSpanlbl.Caption = "Semispan:"
  lblPropRPMlbl.Caption = "Propeller RPM:"
  For Each c In colFixedControls
    c.Visible = True
  Next
End Sub

Private Sub UpdatePropertyControls()
'Transfer the contents of the local data to the display controls
  Dim c As Control
  
  If Trim$(cboName.Text) = "" Then
    ClearPropertyControls
    Exit Sub
  End If
  
  Select Case AC.WingType
  Case 3 'fixed wing
    lblWingType.Caption = "Fixed-wing"
    lblSemiSpanlbl.Caption = "Semispan:"
    lblPropRPMlbl.Caption = "Propeller RPM:"
    For Each c In colFixedControls: c.Visible = True: Next
  Case 4 'helicopter
    lblWingType.Caption = "Helicopter"
    lblSemiSpanlbl.Caption = "Rotor Radius:"
    lblPropRPMlbl.Caption = "Rotor RPM:"
    For Each c In colFixedControls: c.Visible = False: Next
  End Select
  lblSemiSpan.Caption = AGFormat$(UnitsDisplay(AC.SemiSpan, UN_LENGTH))
  lblTypSpeed.Caption = AGFormat$(UnitsDisplay(AC.TypSpeed, UN_SPEED))
  lblBiplSep.Caption = AGFormat$(UnitsDisplay(AC.BiplSep, UN_LENGTH))
  lblWeight.Caption = AGFormat$(UnitsDisplay(AC.Weight, UN_MASS))
  lblPlanArea.Caption = AGFormat$(UnitsDisplay(AC.PlanArea, UN_AREA))
  lblPropRPM.Caption = AGFormat$(AC.PropRPM)
  lblPropRad.Caption = AGFormat$(UnitsDisplay(AC.PropRad, UN_LENGTH))
  lblEngVert.Caption = AGFormat$(UnitsDisplay(AC.EngVert, UN_LENGTH))
  lblEngFwd.Caption = AGFormat$(UnitsDisplay(AC.EngFwd, UN_LENGTH))
  lblNumEng.Caption = AGFormat$(AC.NumEng)
  lblEngHoriz(0).Caption = AGFormat$(UnitsDisplay(AC.EngHoriz(0), UN_LENGTH))
  lblEngHoriz(1).Caption = AGFormat$(UnitsDisplay(AC.EngHoriz(1), UN_LENGTH))
  lblWingVert.Caption = AGFormat$(UnitsDisplay(AC.WingVert, UN_LENGTH))
  lblBoomVert.Caption = AGFormat$(UnitsDisplay(AC.BoomVert, UN_LENGTH))
  lblBoomFwd.Caption = AGFormat$(UnitsDisplay(AC.BoomFwd, UN_LENGTH))
  
End Sub

Private Sub UpdateUnitsLabels()
  lblSemiSpanUnits.Caption = UnitsName(UN_LENGTH)
  lblWeightUnits.Caption = UnitsName(UN_MASS)
  lblTypSpeedUnits.Caption = UnitsName(UN_SPEED)
  lblPropRadUnits.Caption = UnitsName(UN_LENGTH)
  lblBiplSepUnits.Caption = UnitsName(UN_LENGTH)
  lblPlanAreaUnits.Caption = UnitsName(UN_AREA)
  lblEngVertUnits.Caption = UnitsName(UN_LENGTH)
  lblEngFwdUnits.Caption = UnitsName(UN_LENGTH)
  lblEngHorizUnits.Caption = UnitsName(UN_LENGTH)
  lblWingVertUnits.Caption = UnitsName(UN_LENGTH)
  lblBoomVertUnits.Caption = UnitsName(UN_LENGTH)
  lblBoomFwdUnits.Caption = UnitsName(UN_LENGTH)
End Sub

Private Sub cboName_Click()
  UserLibGetAircraftRecord cboName.Text, AC
  UpdatePropertyControls
End Sub

Private Sub cmdCancel_Click()
  OK = False 'cancelled
  Me.Hide
End Sub

Private Sub cmdDelete_Click()
  Dim i As Integer
  If Trim$(cboName.Text) <> "" Then
    If UserLibDeleteAircraftRecord(cboName.Text) Then
      i = cboName.ListIndex
      cboName.RemoveItem i
      'try to keep the same place in the list
      If cboName.ListCount - 1 >= i Then
        cboName.ListIndex = i
      ElseIf cboName.ListCount > 0 Then
        cboName.ListIndex = cboName.ListCount - 1
      Else
        ClearPropertyControls
      End If
    End If
  End If
End Sub

Private Sub cmdOk_Click()
  OK = True 'success! data is in AC
  ACName = cboName.Text 'return aircraft name
  Me.Hide
End Sub

Private Sub Form_Load()
'Initialize this form and its controls
  Dim DB As Database
  Dim RS As Recordset
  
  OK = False  'default form return value
  ACName = "" 'default name selection

  CenterForm Me

  'Initialize collections
  Set colFixedControls = New Collection
  colFixedControls.Add lblPropRadlbl
  colFixedControls.Add lblPropRad
  colFixedControls.Add lblPropRadUnits
  colFixedControls.Add lblBiplSeplbl
  colFixedControls.Add lblBiplSep
  colFixedControls.Add lblBiplSepUnits
  colFixedControls.Add lblPlanArealbl
  colFixedControls.Add lblPlanArea
  colFixedControls.Add lblPlanAreaUnits
  colFixedControls.Add lblEngVertlbl
  colFixedControls.Add lblEngVert
  colFixedControls.Add lblEngVertUnits
  colFixedControls.Add lblEngFwdlbl
  colFixedControls.Add lblEngFwd
  colFixedControls.Add lblEngFwdUnits
  colFixedControls.Add lblNumEnglbl
  colFixedControls.Add lblNumEng
  colFixedControls.Add lblEngHorizlbl
  colFixedControls.Add lblEngHoriz(0)
  colFixedControls.Add lblEngHoriz(1)
  colFixedControls.Add lblEngHorizUnits
  colFixedControls.Add lblWingVertlbl
  colFixedControls.Add lblWingVert
  colFixedControls.Add lblWingVertUnits
 
  ClearPropertyControls
  
  UpdateUnitsLabels

  'Load the name combo
  MatchingIndex = -1 'index for entry that matches current aircraft
  If UserLibOpen(DB, False) Then
    If UserLibOpenRS(DB, "Aircraft", RS) Then
      If Not (RS.BOF And RS.EOF) Then
        While Not RS.EOF
          cboName.AddItem RS("Name")
          RS.MoveNext
        Wend
      End If
      RS.Close
    End If
    DB.Close
  End If
  
  If cboName.ListCount > 0 Then cboName.ListIndex = 0
End Sub

