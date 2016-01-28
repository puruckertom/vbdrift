VERSION 5.00
Begin VB.Form frmTrialLib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Field Trial Library"
   ClientHeight    =   6210
   ClientLeft      =   1650
   ClientTop       =   1785
   ClientWidth     =   6015
   ForeColor       =   &H80000008&
   Icon            =   "TRIALLIB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   6015
   Begin VB.Data datTrials 
      Caption         =   "datTrials"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1175
      Left            =   5040
      TabIndex        =   1
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1175
      Left            =   4080
      TabIndex        =   0
      Top             =   5760
      Width           =   855
   End
   Begin VB.Frame fraFilter 
      Caption         =   "Filter"
      Height          =   3015
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   5775
      Begin VB.ComboBox cboCombine 
         Height          =   315
         HelpContextID   =   1175
         Index           =   0
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox cboRange2 
         Height          =   315
         HelpContextID   =   1175
         Index           =   0
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtValue2 
         Height          =   315
         HelpContextID   =   1175
         Index           =   0
         Left            =   4800
         TabIndex        =   10
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtValue2 
         Height          =   315
         HelpContextID   =   1175
         Index           =   2
         Left            =   4800
         TabIndex        =   20
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtValue1 
         Height          =   315
         HelpContextID   =   1175
         Index           =   2
         Left            =   2280
         TabIndex        =   17
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtValue2 
         Height          =   315
         HelpContextID   =   1175
         Index           =   1
         Left            =   4800
         TabIndex        =   15
         Top             =   2160
         Width           =   855
      End
      Begin VB.ComboBox cboRange2 
         Height          =   315
         HelpContextID   =   1175
         Index           =   2
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2520
         Width           =   615
      End
      Begin VB.ComboBox cboRange1 
         Height          =   315
         HelpContextID   =   1175
         Index           =   2
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2520
         Width           =   615
      End
      Begin VB.ComboBox cboRange2 
         Height          =   315
         HelpContextID   =   1175
         Index           =   1
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2160
         Width           =   615
      End
      Begin VB.ComboBox cboCombine 
         Height          =   315
         HelpContextID   =   1175
         Index           =   2
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox cboCombine 
         Height          =   315
         HelpContextID   =   1175
         Index           =   1
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtValue1 
         Height          =   315
         HelpContextID   =   1175
         Index           =   1
         Left            =   2280
         TabIndex        =   12
         Top             =   2160
         Width           =   855
      End
      Begin VB.ComboBox cboRange1 
         Height          =   315
         HelpContextID   =   1175
         Index           =   1
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtValue1 
         Height          =   315
         HelpContextID   =   1175
         Index           =   0
         Left            =   2280
         TabIndex        =   7
         Top             =   1800
         Width           =   855
      End
      Begin VB.ComboBox cboRange1 
         Height          =   315
         HelpContextID   =   1175
         Index           =   0
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1175
         Index           =   3
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   4095
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1175
         Index           =   2
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1080
         Width           =   4095
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1175
         Index           =   1
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   4095
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1175
         Index           =   0
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label7 
         Caption         =   "Wind Speed (m/s):"
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Boom Height (m):"
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Percent < 141 µm:"
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Aircraft:"
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Test Number:"
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "Test Type:"
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1050
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Browse Filtered Entries"
      Height          =   2415
      Left            =   120
      TabIndex        =   30
      Top             =   3240
      Width           =   5775
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Prev"
         Height          =   375
         HelpContextID   =   1175
         Left            =   720
         TabIndex        =   22
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   375
         HelpContextID   =   1175
         Left            =   1320
         TabIndex        =   23
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "1st"
         Height          =   375
         HelpContextID   =   1175
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   375
         HelpContextID   =   1175
         Left            =   2040
         TabIndex        =   24
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   5280
         TabIndex        =   51
         Top             =   720
         Width           =   120
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "m/s"
         Height          =   195
         Left            =   5280
         TabIndex        =   50
         Top             =   1440
         Width           =   270
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "m"
         Height          =   195
         Left            =   5280
         TabIndex        =   49
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label lblWindSpeed 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4440
         TabIndex        =   48
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblHeight 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4440
         TabIndex        =   47
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblF141 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4440
         TabIndex        =   46
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblAircraft 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   45
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblNumber 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   44
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblMaterial 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   43
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Wind Speed:"
         Height          =   195
         Left            =   3000
         TabIndex        =   42
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Boom Height:"
         Height          =   195
         Left            =   3000
         TabIndex        =   41
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Percent < 141 µm:"
         Height          =   195
         Left            =   3000
         TabIndex        =   40
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Aircraft:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Test Number:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Test Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblPosition 
         AutoSize        =   -1  'True
         Caption         =   "0 of 0"
         Height          =   195
         Left            =   2760
         TabIndex        =   29
         Top             =   2040
         Width           =   510
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmTrialLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: triallib.frm,v 1.4 2001/04/26 16:22:07 tom Exp $
'frmTrialLib - Form for selecting Field Trial data
'This form returns a Tag that is the name of a Field
'Trial Title, which can be used to retrieve an entry
'If the Tag is empty, no entry was selected.
'
Dim PositionCount As Integer  'the current recordset position
Dim RangeField(2) As String 'Field names corresponding to range controls
'this flag is used to tell some controls not to
'take action on their new values. This is required
'to differentiate between programatic state changes
'and user actions
Dim PropTakeAction As Integer 'if true, execute automatic change-related code

Private Sub cboCombine_Click(Index As Integer)
  If PropTakeAction Then
    If cboCombine(Index).ListIndex = 0 Then
      cboRange2(Index).Visible = False
      txtValue2(Index).Visible = False
    Else
      cboRange2(Index).Visible = True
      txtValue2(Index).Visible = True
    End If
    DoEvents 'Let the other controls catch up
    GetNewRecordset
  End If
End Sub


Private Sub cboFilter_Click(Index As Integer)
  'Update the data controls recordset when the Filters change
  GetNewRecordset
End Sub

Private Sub cboRange1_Click(Index As Integer)
  If PropTakeAction Then
    If cboRange1(Index).ListIndex = 0 Then
      txtValue1(Index).Visible = False
      cboCombine(Index).Visible = False
      cboCombine(Index).ListIndex = 0
    Else
      txtValue1(Index).Visible = True
      cboCombine(Index).Visible = True
    End If
    GetNewRecordset
  End If
End Sub


Private Sub cboRange2_Click(Index As Integer)
  If PropTakeAction Then
    GetNewRecordset
  End If
End Sub


Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdFirst_Click()
  If datTrials.Recordset.RecordCount > 0 Then
    datTrials.Recordset.MoveFirst
    PositionCount = 1
    UpdatePositionCount
    UpdatePropertyControls
  End If
End Sub

Private Sub cmdLast_Click()
  If datTrials.Recordset.RecordCount > 0 Then
    datTrials.Recordset.MoveLast
    PositionCount = datTrials.Recordset.RecordCount
    UpdatePositionCount
    UpdatePropertyControls
  End If
End Sub

Private Sub cmdNext_Click()
  If Not datTrials.Recordset.EOF Then
    datTrials.Recordset.MoveNext
    'before we update the Position, see if
    'we've moved off the last record
    If Not datTrials.Recordset.EOF Then
      PositionCount = PositionCount + 1
      UpdatePositionCount
      UpdatePropertyControls
    Else
      datTrials.Recordset.MovePrevious 'go back to where we were
    End If
  End If
End Sub

Private Sub cmdOk_Click()
  Dim DS As Recordset
  Set DS = datTrials.Recordset

  'Return the current title in the Tag
  If DS.EOF Then
    MsgBox "No library entry has been selected.", _
           vbCritical + vbOKOnly
  Else
    Me.Tag = Trim$(DS.Fields("Title"))
    Me.Hide
  End If
End Sub

Private Sub cmdPrev_Click()
  If Not datTrials.Recordset.BOF Then
    datTrials.Recordset.MovePrevious
    'before we update the Position, see if
    'we've moved off the first record
    If Not datTrials.Recordset.BOF Then
      PositionCount = PositionCount - 1
      UpdatePositionCount
      UpdatePropertyControls
    Else
      datTrials.Recordset.MoveNext 'go back to where we were
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

Private Sub Form_Load()
  InitForm
End Sub

Private Sub GetNewRecordset()
'Build a new query from the values of the control array
'and retrieve a new recordset from
'a data control by resetting its RecordSource
'
'
  Dim s As String        'the new query
  Dim stmp As String     'a temporary string
  Dim needsep As Integer 'flag
  Dim needwhere As Integer 'flag
  Dim CompStr As String  'component value for search
  Dim QueryStr As String 'String to hold query
  Dim DS As Recordset      'Recordset of components for search

  'set the basic query that would return all records
  'from the Nozzles Table
  s = "SELECT * FROM FieldTrial"
  
  'preview the filter fields to see if we need to add "WHERE"
  needwhere = False
  For i = 0 To 3
    If cboFilter(i).ListIndex > 0 Then
      needwhere = True
      Exit For
    End If
  Next
  For i = 0 To 2
    If cboRange1(i).ListIndex > 0 Then
      needwhere = True
      Exit For
    End If
  Next
  If needwhere Then s = s + " WHERE"
  
  'set a flag for adding separators
  needsep = False
  'Title
  If cboFilter(0).ListIndex > 0 Then
    s = s + " Title= '" + cboFilter(0).List(cboFilter(0).ListIndex) + "'"
    needsep = True
  End If
  'Material
  If cboFilter(1).ListIndex > 0 Then
    If needsep Then s = s + " AND"
    s = s + " SprayMaterial= " + CStr(cboFilter(1).ListIndex - 1)
    needsep = True
  End If
  'Test number
  If cboFilter(2).ListIndex > 0 Then
    If needsep Then s = s + " AND"
    s = s + " TestNumber= " + CStr(cboFilter(2).List(cboFilter(2).ListIndex))
    needsep = True
  End If
  'Aircraft
  If cboFilter(3).ListIndex > 0 Then
    If needsep Then s = s + " AND"
    s = s + " Aircraft= " + CStr(cboFilter(3).ListIndex - 1)
    needsep = True
  End If
  
  'range-oriented filters
  For i = 0 To 2
    stmp = ""
    'primary range
    If cboRange1(i).ListIndex > 0 Then
      stmp = stmp + " " + RangeField(i) + _
             CStr(cboRange1(i).List(cboRange1(i).ListIndex)) + _
             CStr(Val(txtValue1(i).Text))
      If cboCombine(i).ListIndex > 0 Then
        stmp = stmp + " " + CStr(cboCombine(i).List(cboCombine(i).ListIndex))
        'secondary range
        stmp = stmp + " " + RangeField(i) + _
               CStr(cboRange2(i).List(cboRange2(i).ListIndex)) + _
               CStr(Val(txtValue2(i).Text))
      End If
    End If
    If stmp <> "" Then
      If needsep Then s = s + " AND"
      s = s + " (" + stmp + ")"
      needsep = True
    End If
  Next
  
  'Reset the RecordSource property and refresh to get the
  'new recordset
  datTrials.RecordSource = s
  datTrials.Refresh
  'Update the display for number of records in this set
  UpdateNumRecs
  UpdatePropertyControls
End Sub

Private Sub InitForm()
'Initialize this form and its controls
  Dim DStmp As Recordset
  Dim c As Control
  Dim s As String
  Dim dslist(3) As String

  'This form uses the Tag property to return
  'a database key. The Tag will be blank if no
  'key was chosen
  'set to blank by default
  Me.Tag = ""

  'Center the form
  CenterForm Me

  'set the database and initial query for the data controls
  datTrials.ReadOnly = True 'open database as read only
  datTrials.DatabaseName = UI.LibraryPath
  datTrials.RecordSource = "FieldTrial"
  datTrials.Refresh
  UpdateNumRecs

  'Disable some control reactions
  PropTakeAction = False
  
  'clear filter controls
  For i = 0 To 3
    cboFilter(i).Clear
  Next

  'Fill the filter combo boxes
  'Pad integer fields with zeros so that they sort properly
  Set DStmp = datTrials.Database.OpenRecordset("FieldTrial")
  DStmp.MoveFirst
  Do Until DStmp.EOF
    cboFilter(0).AddItem DStmp.Fields("Title")
    ConditionalAddItem Format$(DStmp.Fields("TestNumber"), "00#"), cboFilter(2)
    DStmp.MoveNext
  Loop
  DStmp.Close
  
  'Material
  cboFilter(1).AddItem "Standard"
  cboFilter(1).AddItem "Variable"
  
  'Aircraft
  cboFilter(3).AddItem GetBasicNameAC(0)
  cboFilter(3).AddItem GetBasicNameAC(1)
  cboFilter(3).AddItem GetBasicNameAC(2)
  
  'Add "Any" as the first item, indicating that we will not
  'filter based on this field, and select first list item
  For i = 0 To 3
    cboFilter(i).AddItem "Any", 0
    cboFilter(i).ListIndex = 0
  Next
  
  'Set up the range-oriented filter controls
  RangeField(0) = "F141"
  RangeField(1) = "Height"
  RangeField(2) = "WindSpeed"
  '
  For i = 0 To 2
    cboRange1(i).Clear
    cboRange1(i).AddItem "Any"
    cboRange1(i).AddItem "="
    cboRange1(i).AddItem "<>"
    cboRange1(i).AddItem "<"
    cboRange1(i).AddItem ">"
    cboRange1(i).AddItem "<="
    cboRange1(i).AddItem ">="
    cboRange1(i).ListIndex = 0
    '
    txtValue1(i).Text = ""
    txtValue1(i).Visible = False
    '
    cboCombine(i).Clear
    cboCombine(i).AddItem ""
    cboCombine(i).AddItem "And"
    cboCombine(i).AddItem "Or"
    cboCombine(i).ListIndex = 0
    cboCombine(i).Visible = False
    '
    cboRange2(i).Clear
    cboRange2(i).AddItem "="
    cboRange2(i).AddItem "<>"
    cboRange2(i).AddItem "<"
    cboRange2(i).AddItem ">"
    cboRange2(i).AddItem "<="
    cboRange2(i).AddItem ">="
    cboRange2(i).ListIndex = 0
    cboRange2(i).Visible = False
    '
    txtValue2(i).Text = ""
    txtValue2(i).Visible = False
  Next
  
  'Enable control reactions
  PropTakeAction = True
  
  UpdatePropertyControls
End Sub


Private Sub UpdateNumRecs()
'Ensure the validity of RecordCount by moving to the
'last record. Also update the PositionCount variable
'that is global to this form

  'if BOF and EOF are true, there are no records
  If Not datTrials.Recordset.BOF And Not datTrials.Recordset.EOF Then
    datTrials.Recordset.MoveLast
    datTrials.Recordset.MoveFirst
    PositionCount = 1
  Else
    PositionCount = 0
  End If
  UpdatePositionCount
End Sub

Private Sub UpdatePositionCount()
'update the caption of the position label
  lblPosition.Caption = Format$(PositionCount) + " of " + Format$(datTrials.Recordset.RecordCount)
End Sub

Private Sub UpdatePropertyControls()
'Update the property display controls
  
  Dim DS As Recordset

  Set DS = datTrials.Recordset

  'If there is no current record, turn on the labels and exit
  If DS.EOF Then
    lblName.Caption = ""
    lblMaterial.Caption = ""
    lblNumber.Caption = ""
    lblAircraft.Caption = ""
    lblF141.Caption = ""
    lblHeight.Caption = ""
    lblWindSpeed.Caption = ""
    Exit Sub
  End If

  'transfer the current record data to the controls
  lblName.Caption = DS.Fields("Title")
  Select Case DS.Fields("SprayMaterial")
    Case 0:
      lblMaterial.Caption = "Standard"
    Case 1:
      lblMaterial.Caption = "Variable"
  End Select
  lblNumber.Caption = Format$(DS.Fields("TestNumber"), "00#")
  lblAircraft.Caption = GetBasicNameAC(DS.Fields("Aircraft"))
  lblF141.Caption = DS.Fields("F141")
  lblHeight.Caption = DS.Fields("Height")
  lblWindSpeed.Caption = DS.Fields("WindSpeed")
  
  'update the field labels and properties
End Sub



Private Sub Timer1_Timer()
  Timer1.Interval = 0 'turn off the timer
  GetNewRecordset
End Sub


Private Sub txtValue1_Change(Index As Integer)
  'Restart the timer interval. When it times out
  'it will do a GetNewRecordSet
  Timer1.Interval = 1000   'milliseconds
End Sub


Private Sub txtValue2_Change(Index As Integer)
  'Restart the timer interval. When it times out
  'it will do a GetNewRecordSet
  Timer1.Interval = 1000   'milliseconds
End Sub


