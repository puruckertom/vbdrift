VERSION 5.00
Begin VB.Form frmDropLib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Drop Size Distribution Library"
   ClientHeight    =   6795
   ClientLeft      =   1935
   ClientTop       =   2745
   ClientWidth     =   9480
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   HelpContextID   =   1120
   Icon            =   "DROPLIB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Begin VB.Data datNozzleTypMfg 
      Caption         =   "datNozzleTypMfg"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4680
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data datComp 
      Caption         =   "datComp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data datDrop 
      Caption         =   "datDrop"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1120
      Left            =   8520
      TabIndex        =   1
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1120
      Left            =   7560
      TabIndex        =   0
      Top             =   6360
      Width           =   855
   End
   Begin VB.Frame fraFilter 
      Caption         =   "Filter"
      Height          =   6135
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1477
         Index           =   6
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1442
         Index           =   5
         Left            =   1575
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2160
         Width           =   2175
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1073
         Index           =   4
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1259
         Index           =   0
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1187
         Index           =   1
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1183
         Index           =   2
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1310
         Index           =   3
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Nozzle RPM:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1845
         Width           =   1335
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nozzle Pressure:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   2205
         Width           =   1185
      End
      Begin VB.Label lblPressFilterUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3840
         TabIndex        =   38
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label lblAirSpeedFilterUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3840
         TabIndex        =   34
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label lblNozAngFilterUnits 
         AutoSize        =   -1  'True
         Caption         =   "deg"
         Height          =   195
         Left            =   3840
         TabIndex        =   33
         Top             =   1485
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Component:"
         Height          =   195
         Left            =   105
         TabIndex        =   30
         Top             =   405
         Width           =   1020
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Substance:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nozzle:"
         Height          =   195
         Left            =   105
         TabIndex        =   27
         Top             =   1125
         Width           =   525
      End
      Begin VB.Label Label8 
         Caption         =   "Nozzle Orientation:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1480
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Air Speed:"
         Height          =   195
         Left            =   105
         TabIndex        =   25
         Top             =   2565
         Width           =   900
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Browse Filtered Entries"
      Height          =   6135
      Left            =   4680
      TabIndex        =   20
      Top             =   120
      Width           =   4695
      Begin VB.ListBox lstComponents 
         Height          =   1620
         HelpContextID   =   1120
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   4455
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Prev"
         Height          =   375
         HelpContextID   =   1120
         Left            =   720
         TabIndex        =   11
         Top             =   5640
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   375
         HelpContextID   =   1120
         Left            =   1320
         TabIndex        =   12
         Top             =   5640
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "1st"
         Height          =   375
         HelpContextID   =   1120
         Left            =   120
         TabIndex        =   10
         Top             =   5640
         Width           =   615
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   375
         HelpContextID   =   1120
         Left            =   2040
         TabIndex        =   13
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Drop Size Classification:"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblSprayQuality 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblSprayQuality"
         Height          =   285
         Left            =   1920
         TabIndex        =   47
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label lblType 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblType"
         DataField       =   "Type"
         DataSource      =   "datNozzleTypMfg"
         Height          =   285
         Left            =   1920
         TabIndex        =   46
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblManufacturer 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblManufacturer"
         DataField       =   "Manufacturer"
         DataSource      =   "datNozzleTypMfg"
         Height          =   285
         Left            =   1920
         TabIndex        =   45
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nozzle Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Nozzle Manufacturer:"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nozzle RPM:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   2205
         Width           =   930
      End
      Begin VB.Label lblNozRPM 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNozRPM"
         Height          =   285
         Left            =   1920
         TabIndex        =   41
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nozzle Pressure:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   2580
         Width           =   1185
      End
      Begin VB.Label lblPress 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblPress"
         Height          =   285
         Left            =   1920
         TabIndex        =   36
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblPressUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   4200
         TabIndex        =   35
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label lblAirSpeedUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   4200
         TabIndex        =   32
         Top             =   2880
         Width           =   420
      End
      Begin VB.Label lblNozAngUnits 
         AutoSize        =   -1  'True
         Caption         =   "deg"
         Height          =   195
         Left            =   4200
         TabIndex        =   31
         Top             =   1845
         Width           =   330
      End
      Begin VB.Label lblWS 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblWS"
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblNozAng 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNozAng"
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblNoz 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNoz"
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblSubst 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblSubst"
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Components:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   3600
         Width           =   1110
      End
      Begin VB.Label lblPosition 
         AutoSize        =   -1  'True
         Caption         =   "0 of 0"
         Height          =   195
         Left            =   2760
         TabIndex        =   28
         Top             =   5760
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Air Speed:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2925
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nozzle Orientation:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1845
         Width           =   1635
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nozzle:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   765
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Substance:"
         Height          =   195
         Left            =   135
         TabIndex        =   21
         Top             =   405
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Edit Library"
      Height          =   735
      Left            =   120
      TabIndex        =   49
      Top             =   6120
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdCurrent 
         Caption         =   "Current"
         Height          =   375
         HelpContextID   =   1020
         Left            =   600
         TabIndex        =   53
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         HelpContextID   =   1020
         Left            =   1560
         TabIndex        =   52
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "Modify"
         Height          =   375
         HelpContextID   =   1020
         Left            =   2280
         TabIndex        =   51
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         HelpContextID   =   1020
         Left            =   3000
         TabIndex        =   50
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmDropLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: droplib.frm,v 1.7 2001/05/24 20:16:18 tom Exp $
'This form browses the Dropsize "library".
'
' The Tag property is used for i/o as follows:
'
' output:  a "key" that identifies the selected library
'          record. If blank, no record was selected
'
Dim PositionCount As Integer  'the current recordset position
Dim FSstr As String           'String to append to table names

Public Sub SelectTable(fstr As String)
'Initialize this form and its controls
  Dim DStmp As Recordset
  Dim c As Control
  Dim s As String
  Dim atmp As Single
  ReDim dslist(5) As String

  FSstr = fstr

  'set the database and initial query for the data controls
  datDrop.ReadOnly = True 'open database as read only
  datDrop.DatabaseName = UI.LibraryPath
  datDrop.RecordSource = "Dropsize" + FSstr
  datDrop.Refresh
  UpdateNumRecs

  If FSstr = "" Then
    datComp.ReadOnly = True 'open database as read only
    datComp.DatabaseName = UI.LibraryPath
    datComp.RecordSource = "Components"
    datComp.Refresh
  End If
  
  'clear filter combo boxes
  For i = 0 To 6
    cboFilter(i).Clear
  Next
  
  'Fill the combo box lists with the unique lists from
  'the database
  'Pad integer fields with zeros so that they sort properly
  
  Set DStmp = datDrop.Database.OpenRecordset("SubstanceList" + FSstr)
  DStmp.MoveFirst
  Do Until DStmp.EOF
    cboFilter(0).AddItem DStmp.Fields(0)
    DStmp.MoveNext
  Loop
  DStmp.Close
  
  Set DStmp = datDrop.Database.OpenRecordset("NozzleList" + FSstr)
  DStmp.MoveFirst
  Do Until DStmp.EOF
    cboFilter(1).AddItem DStmp.Fields(0)
    DStmp.MoveNext
  Loop
  DStmp.Close
  
  Set DStmp = datDrop.Database.OpenRecordset("NozzleAngleList" + FSstr)
  DStmp.MoveFirst
  Do Until DStmp.EOF
    cboFilter(2).AddItem Format$(DStmp.Fields(0), "@@@")
    DStmp.MoveNext
  Loop
  DStmp.Close
  'now clean up the list formatting without losing the sorting
  For i = 0 To cboFilter(2).ListCount - 1
    s = cboFilter(2).List(i)
    cboFilter(2).RemoveItem i
    cboFilter(2).AddItem Trim$(s), i
  Next
  
  Set DStmp = datDrop.Database.OpenRecordset("NozzleRPMList" + FSstr)
  DStmp.MoveFirst
  Do Until DStmp.EOF
    cboFilter(6).AddItem Format$(DStmp.Fields(0), "@@@@@")
    DStmp.MoveNext
  Loop
  DStmp.Close
  'now clean up the list formatting without losing the sorting
  For i = 0 To cboFilter(6).ListCount - 1
    s = cboFilter(6).List(i)
    cboFilter(6).RemoveItem i
    cboFilter(6).AddItem Trim$(s), i
  Next
  
  Set DStmp = datDrop.Database.OpenRecordset("WindSpeedList" + FSstr)
  DStmp.MoveFirst
  Do Until DStmp.EOF
    cboFilter(3).AddItem Format$(UnitsDisplay(DStmp.Fields(0), UN_SPEED), "0000.0000")
    DStmp.MoveNext
  Loop
  DStmp.Close
  'now clean up the list formatting without losing the sorting
  For i = 0 To cboFilter(3).ListCount - 1
    atmp = cboFilter(3).List(i)
    cboFilter(3).RemoveItem i
    cboFilter(3).AddItem AGFormat$(atmp), i
  Next
  
  'the fourth filter comes from a different place
  If FSstr = "" Then
    Set DStmp = datComp.Database.OpenRecordset("ComponentList" + FSstr)
    DStmp.MoveFirst
    Do Until DStmp.EOF
      cboFilter(4).AddItem DStmp.Fields(0)
      DStmp.MoveNext
    Loop
    DStmp.Close
  End If
  
  Set DStmp = datDrop.Database.OpenRecordset("PressureList" + FSstr)
  DStmp.MoveFirst
  Do Until DStmp.EOF
    cboFilter(5).AddItem Format$(UnitsDisplay(DStmp.Fields(0), UN_PRESSURE), "0000.0000")
    DStmp.MoveNext
  Loop
  DStmp.Close
  'now clean up the list formatting without losing the sorting
  For i = 0 To cboFilter(5).ListCount - 1
    atmp = cboFilter(5).List(i)
    cboFilter(5).RemoveItem i
    cboFilter(5).AddItem AGFormat$(atmp), i
  Next
  
  'Add "Any" as the first item, indicating that we will not
  'filter based on this field
  For i = 0 To 6
    cboFilter(i).AddItem "Any", 0
  Next
  
  'Set combo boxes to first item
  For i = 0 To 6
    cboFilter(i).ListIndex = 0
  Next

  'units labels
  lblPressUnits.Caption = UnitsName(UN_PRESSURE)
  lblPressFilterUnits.Caption = UnitsName(UN_PRESSURE)
  lblAirSpeedUnits.Caption = UnitsName(UN_SPEED)
  lblAirSpeedFilterUnits.Caption = UnitsName(UN_SPEED)

End Sub

Private Sub cboFilter_Click(Index As Integer)
  'Update the data controls recordset when the Filters change
  GetNewRecordset
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdFirst_Click()
  If datDrop.Recordset.RecordCount > 0 Then
    datDrop.Recordset.MoveFirst
    PositionCount = 1
  UpdatePositionCount
  UpdatePropertyControls
  End If
End Sub

Private Sub cmdLast_Click()
  If datDrop.Recordset.RecordCount > 0 Then
    datDrop.Recordset.MoveLast
    PositionCount = datDrop.Recordset.RecordCount
    UpdatePositionCount
    UpdatePropertyControls
  End If
End Sub

Private Sub cmdNext_Click()
  If Not datDrop.Recordset.EOF Then
    datDrop.Recordset.MoveNext
    'before we update the Position, see if
    'we've moved off the last record
    If Not datDrop.Recordset.EOF Then
      PositionCount = PositionCount + 1
      UpdatePositionCount
      UpdatePropertyControls
    Else
      datDrop.Recordset.MovePrevious 'go back to where we were
    End If
  End If
End Sub

Private Sub cmdOk_Click()
  Dim DS As Recordset
  Dim s As String
  
  Set DS = datDrop.Recordset
  If DS.BOF And DS.EOF Then
    s = "No library entry has been selected."
    t% = vbCritical + vbOKOnly
    MsgBox s$, t%
  Else
    'Transfer the library key
    If FSstr = "" Then
      Me.Tag = "0," 'regulatory
    Else
      Me.Tag = "1," 'FS
    End If
    Me.Tag = Me.Tag + Trim$(DS("Substance")) + ","
    Me.Tag = Me.Tag + Trim$(DS("Nozzle")) + ","
    Me.Tag = Me.Tag + Format$(DS("NozzleAngle")) + ","
    Me.Tag = Me.Tag + Format$(DS("NozzleRPM")) + ","
    Me.Tag = Me.Tag + Format$(DS("Pressure")) + ","
    Me.Tag = Me.Tag + Format$(DS("WindSpeed"))
    Me.Hide
  End If
End Sub

Private Sub cmdPrev_Click()
  If Not datDrop.Recordset.BOF Then
    datDrop.Recordset.MovePrevious
    'before we update the Position, see if
    'we've moved off the first record
    If Not datDrop.Recordset.BOF Then
      PositionCount = PositionCount - 1
      UpdatePositionCount
      UpdatePropertyControls
    Else
      datDrop.Recordset.MoveNext 'go back to where we were
    End If
  End If
End Sub

Private Sub ConditionalAddItem(s As String, c As Control)
'Add an item to a List control only if it is not
'already in the list
'
  For i = 0 To c.ListCount - 1
    If (s = c.List(i)) Then Exit Sub
  Next
  c.AddItem s
End Sub

Private Sub datDrop_Reposition()
  'When this data control repositions, update the
  'datNozzleTypMfg control to match
  If Not datDrop.Recordset.BOF And _
     Not datDrop.Recordset.EOF Then
    datNozzleTypMfg.RecordSource = _
      "select * from NozzlesTypMfg where Nozzle='" & _
      datDrop.Recordset.Fields("Nozzle") & "'"
    datNozzleTypMfg.Refresh
  End If
End Sub

Private Sub Form_Load()
  'This form uses the Tag property to return
  'a key suitable for retrieving a record from
  'the library. Set the key to blank as its default
  Me.Tag = ""

  'Center the form
  CenterForm Me

  datNozzleTypMfg.ReadOnly = True 'open database as read only
  datNozzleTypMfg.DatabaseName = UI.LibraryPath
  datNozzleTypMfg.RecordSource = "NozzlesTypMfg"
  datNozzleTypMfg.Refresh

End Sub

Private Sub GetNewRecordset()
'Build a new query from the values of the control array
'and retrieve a new recordset from
'a data control by resetting its RecordSource
'
'
  Dim s As String        'the new query
  Dim needsep As Integer 'flag
  Dim CompStr As String  'component value for search
  Dim QueryStr As String 'String to hold query
  Dim DS As Recordset      'Recordset of components for search
  Dim midrange As Single
  Dim minrange As Single
  Dim maxrange As Single

  'set the basic query that would return all records
  'from the Dropsize Table
  s = "SELECT * FROM Dropsize" + FSstr
  'preview the filer fields to see if we need to add "WHERE"
  For i = 0 To 6
    If cboFilter(i).ListIndex > 0 Then
      s = s + " WHERE"
      Exit For
    End If
  Next
  'set a flag for adding separators
  needsep = False
  If cboFilter(0).ListIndex > 0 Then
    s = s + " Substance= '" + cboFilter(0).List(cboFilter(0).ListIndex) + "'"
    needsep = True
  End If
  If cboFilter(1).ListIndex > 0 Then
    If needsep Then s = s + " AND"
    s = s + " Nozzle= '" + cboFilter(1).List(cboFilter(1).ListIndex) + "'"
    needsep = True
  End If
  If cboFilter(2).ListIndex > 0 Then
    If needsep Then s = s + " AND"
    s = s + " NozzleAngle= " + cboFilter(2).List(cboFilter(2).ListIndex)
    needsep = True
  End If
  If cboFilter(6).ListIndex > 0 Then
    If needsep Then s = s + " AND"
    s = s + " NozzleRPM= " + cboFilter(6).List(cboFilter(6).ListIndex)
    needsep = True
  End If
  If cboFilter(3).ListIndex > 0 Then
    'since windspeed is not an integer, we must look for a range
    If needsep Then s = s + " AND"
    midrange = Val(AGFormat$(UnitsInternal((Val(cboFilter(3).List(cboFilter(3).ListIndex))), UN_SPEED)))
    If midrange <> 0 Then
      minrange = midrange * 0.99  'val - 1%
      maxrange = midrange * 1.01  'val + 1%
      s = s + " WindSpeed >" + AGFormat$(minrange) + " AND WindSpeed <" + AGFormat$(maxrange)
    Else
      s = s + " WindSpeed = 0"
    End If
    needsep = True
  End If
  
  'If the Component filter is set to something other than
  '"Any", we must build an additional query section from
  'the components list. We must search the Components Table
  'for Substance ID's that contain the desired component,
  'then build the additional query based on those ID's
  If cboFilter(4).ListIndex > 0 Then
    'recover the component for the search
    CompStr = cboFilter(4).List(cboFilter(4).ListIndex)
    'pad the string with spaces to avoid substring matches
    CompStr = CompStr & Space$(32 - Len(CompStr))
    'build the SQL query
    QueryStr = "SELECT * FROM Components WHERE Component LIKE '*" & CompStr & "*'"
    'Create a Recordset of matching records
    Set DS = datComp.Database.OpenRecordset(QueryStr, dbReadOnly)
    'add a separator if need be
    If needsep Then s = s + " AND"
    needsep = True
    'start to build this portion of SQL query
    s = s & " Substance IN ("
    'process first record
    DS.MoveFirst
    SubstStr = DS.Fields("Substance").Value
    s = s & "'" & SubstStr & "'"
    DS.MoveNext
    'process other records
    While Not DS.EOF
      SubstStr = DS.Fields("Substance").Value
      s = s & ",'" & SubstStr & "'"
      DS.MoveNext
    Wend
    'complete this portion of the SQL query
    s = s & ")"
    'close out the Recordset
    DS.Close
  End If
  
  If cboFilter(5).ListIndex > 0 Then
    'since pressure is not an integer, we must look for a range
    If needsep Then s = s + " AND"
    midrange = Val(AGFormat$(UnitsInternal((Val(cboFilter(5).List(cboFilter(5).ListIndex))), UN_PRESSURE)))
    If midrange <> 0 Then
      minrange = midrange * 0.99  'val - 1%
      maxrange = midrange * 1.01  'val + 1%
      s = s + " Pressure >" + AGFormat$(minrange) + " AND Pressure <" + AGFormat$(maxrange)
    Else
      s = s + " Pressure = 0"
    End If
    needsep = True
  End If
  
  'Reset the RecordSource property and refresh to get the
  'new recordset
  datDrop.RecordSource = s
  datDrop.Refresh
  'Update the display for number of records in this set
  UpdateNumRecs
  'after moving to a new record, update the component list
  UpdatePropertyControls
End Sub

Private Sub UpdateComponentList()
'find the components for the current entry and
'fill the list box
'
  Dim tmpFieldC As Field
  Dim tmpFieldP As Field
  Dim tmpFieldN As Field
  Dim CompStr As String
  Dim FracStr As String
  Dim SubStr As String
  Dim s As String
  ReDim dat(6) As Single

  'clear the list control
  lstComponents.Clear

  If FSstr = "FS" Then Exit Sub
  
  'get the subtance code for the search from the datDrop data control
  If datDrop.Recordset.BOF And datDrop.Recordset.EOF Then Exit Sub
  SubStr = datDrop.Recordset.Fields("Substance").Value

  'find the record in the datComp data control with the matching Substance
  datComp.Recordset.FindFirst "[Substance] = '" & SubStr & "'"
  If datComp.Recordset.NoMatch Then Exit Sub
  
  'get field value (all strings are packed into one)
  Set tmpFieldC = datComp.Recordset.Fields("Component")
  Set tmpFieldP = datComp.Recordset.Fields("Percent")
  Set tmpFieldN = datComp.Recordset.Fields("NumComponents")
   
  CompStr = tmpFieldC.Value  'String comtaining all components
  FieldToArray tmpFieldP, dat() 'array containing all percents
  num = Val(tmpFieldN.Value) 'number of components

  'load up the form controls
  For i = 0 To num - 1
    s = Trim$(Mid$(CompStr, i * 32 + 1, 32))
    s = s & " (" & Format$(dat(i)) & "%)"
    lstComponents.AddItem s
  Next
  
End Sub

Private Sub UpdateNumRecs()
'Ensure the validity of RecordCount by moving to the
'last record. Also update the PositionCount variable
'that is global to this form

  'if BOF and EOF are true, there are no records
  If Not datDrop.Recordset.BOF And Not datDrop.Recordset.EOF Then
    datDrop.Recordset.MoveLast
    datDrop.Recordset.MoveFirst
    PositionCount = 1
  Else
    PositionCount = 0
  End If
  UpdatePositionCount
End Sub

Private Sub UpdatePositionCount()
'update the caption of the position label
  lblPosition.Caption = Format$(PositionCount) + " of " + Format$(datDrop.Recordset.RecordCount)
End Sub

Private Sub UpdatePropertyControls()
'Update the property display controls
  
  Dim DS As Recordset

  Set DS = datDrop.Recordset

  'If there is no current record, clear the labels
  If DS.EOF Then
    lblSubst.Caption = ""
    lblNoz.Caption = ""
    lblNozAng.Caption = ""
    lblNozRPM.Caption = ""
    lblPress.Caption = ""
    lblWS.Caption = ""
    lblSprayQuality = ""
  Else
    'transfer the current record data to the controls
    lblSubst.Caption = DS.Fields("Substance")
    lblNoz.Caption = DS.Fields("Nozzle")
    lblNozAng.Caption = AGFormat$(DS.Fields("NozzleAngle"))
    lblNozRPM.Caption = AGFormat$(DS.Fields("NozzleRPM"))
    lblPress.Caption = AGFormat$(UnitsDisplay(DS.Fields("Pressure"), UN_PRESSURE))
    'special treatment for WindSpeed, since it is a Single, but is displayed as an int
    lblWS.Caption = AGFormat$(UnitsDisplay(DS.Fields("WindSpeed"), UN_SPEED))
    lblSprayQuality.Caption = GetBasicNameDSD(DS.Fields("SprayQuality"))
  End If
  
  'always update the component list, even if just
  'to clear the control
  UpdateComponentList
End Sub

