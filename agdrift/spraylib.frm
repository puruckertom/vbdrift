VERSION 5.00
Begin VB.Form frmSprayLib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spray Material Library"
   ClientHeight    =   4560
   ClientLeft      =   1710
   ClientTop       =   1560
   ClientWidth     =   8760
   ForeColor       =   &H80000008&
   Icon            =   "SPRAYLIB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4560
   ScaleWidth      =   8760
   Begin VB.Data datMat 
      Caption         =   "datMat"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2400
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
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
      Height          =   300
      Left            =   4560
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data datEvap 
      Caption         =   "datEvap"
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
      Top             =   4080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1258
      Left            =   7800
      TabIndex        =   1
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1258
      Left            =   6840
      TabIndex        =   0
      Top             =   4080
      Width           =   855
   End
   Begin VB.Frame fraFilter 
      Caption         =   "Filter"
      Height          =   3855
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1073
         Index           =   1
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1259
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Component:"
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   405
         Width           =   1020
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Substance:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   765
         Width           =   975
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Browse Filtered Entries"
      Height          =   3855
      Left            =   4560
      TabIndex        =   14
      Top             =   120
      Width           =   4095
      Begin VB.ListBox lstComponents 
         Height          =   1230
         HelpContextID   =   1258
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   3855
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Prev"
         Height          =   375
         HelpContextID   =   1258
         Left            =   720
         TabIndex        =   5
         Top             =   3360
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   375
         HelpContextID   =   1258
         Left            =   1320
         TabIndex        =   6
         Top             =   3360
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "1st"
         Height          =   375
         HelpContextID   =   1258
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Width           =   615
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   375
         HelpContextID   =   1258
         Left            =   2040
         TabIndex        =   7
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Specific Gravity:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1485
         Width           =   1425
      End
      Begin VB.Label lblSpecGrav 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblSpecGrav"
         DataField       =   "Density"
         DataSource      =   "datMat"
         Height          =   285
         Left            =   2160
         TabIndex        =   21
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblNVFrac 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNVFrac"
         DataField       =   "NonvolFraction"
         DataSource      =   "datEvap"
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblEvapRate 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblEvapRate"
         DataField       =   "EvaporationRate"
         DataSource      =   "datEvap"
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblSubst 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblSubst"
         DataField       =   "Substance"
         DataSource      =   "datEvap"
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Components:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   1110
      End
      Begin VB.Label lblPosition 
         AutoSize        =   -1  'True
         Caption         =   "0 of 0"
         Height          =   195
         Left            =   2760
         TabIndex        =   20
         Top             =   3480
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nonvolatile Fraction:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1125
         Width           =   1785
      End
      Begin VB.Label Label4 
         Caption         =   "Evaporation Rate:   (µm²/deg C/sec)"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   675
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Substance:"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   405
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSprayLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Spray Library Form
' $Id: spraylib.frm,v 1.6 2006/11/08 15:18:11 tom Exp $
'
'Form Key:
' On Input:
'  not used
' On Return
'  True=Okay pressed, data transferred
'  False=operation cancelled

Dim PositionCount As Integer  'the current recordset position

'Private variables to complement form properties
Private mName As String
Private mNVFrac As Single
Private mACFrac As Single
Private mFlowRate As Single
Private mFlowRateUnits As Integer
Private mNonVGrav As Single
Private mEvapRate As Single

Private mCancelled As Boolean

Private Sub cboFilter_Click(Index As Integer)
  'Update the data controls recordset when the Filters change
  GetNewRecordset
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdFirst_Click()
  If datEvap.Recordset.RecordCount > 0 Then
    datEvap.Recordset.MoveFirst
    PositionCount = 1
    UpdatePositionCount
    SyncDisplayControls
  End If
End Sub

Private Sub cmdLast_Click()
  If datEvap.Recordset.RecordCount > 0 Then
    datEvap.Recordset.MoveLast
    PositionCount = datEvap.Recordset.RecordCount
    UpdatePositionCount
    SyncDisplayControls
  End If
End Sub

Private Sub cmdNext_Click()
  If Not datEvap.Recordset.EOF Then
    datEvap.Recordset.MoveNext
    'before we update the Position, see if
    'we've moved off the last record
    If Not datEvap.Recordset.EOF Then
      PositionCount = PositionCount + 1
      UpdatePositionCount
      SyncDisplayControls
    Else
      datEvap.Recordset.MovePrevious 'go back to where we were
    End If
  End If
End Sub

Private Sub cmdOk_Click()
  If TransferData() Then
    mCancelled = False
    Me.Hide
  Else
    s$ = "No library entry has been selected."
    t% = vbCritical + vbOKOnly
    MsgBox s$, t%
  End If
End Sub

Private Sub cmdPrev_Click()
  If Not datEvap.Recordset.BOF Then
    datEvap.Recordset.MovePrevious
    'before we update the Position, see if
    'we've moved off the first record
    If Not datEvap.Recordset.BOF Then
      PositionCount = PositionCount - 1
      UpdatePositionCount
      SyncDisplayControls
    Else
      datEvap.Recordset.MoveNext 'go back to where we were
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
  Dim needsep As Integer 'flag
  Dim CompStr As String  'component value for search
  Dim QueryStr As String 'String to hold query
  Dim DS As Recordset      'Recordset of components for search

  'set the basic query that would return all records
  'from the Dropsize Table
  s = "SELECT * FROM Evaporation"
  'preview the filer fields to see if we need to add "WHERE"
  For i = 0 To 1
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
  
  'If the Component filter is set to something other than
  '"Any", we must build an additional query section from
  'the components list. We must search the Components Table
  'for Substance ID's that contain the desired component,
  'then build the additional query based on those ID's
  If cboFilter(1).ListIndex > 0 Then
    'recover the component for the search
    CompStr = cboFilter(1).List(cboFilter(1).ListIndex)
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
  
  'Reset the RecordSource property and refresh to get the
  'new recordset
  datEvap.RecordSource = s
  datEvap.Refresh
  'Update the display for number of records in this set
  UpdateNumRecs
  'after moving to a new record, update the display controls
  SyncDisplayControls
End Sub

Private Sub InitForm()
'Initialize this form and its controls
  Dim DStmp As Recordset
  Dim c As Control
  Dim s As String

  mCancelled = True 'Set this to false only if OK is clicked

  'Center the form
  CenterForm Me

  'set the database and initial query for the data controls
  datEvap.ReadOnly = True 'open database as read only
  datEvap.DatabaseName = UI.LibraryPath
  datEvap.RecordSource = "Evaporation"
  datEvap.Refresh
  UpdateNumRecs

  datComp.ReadOnly = True 'open database as read only
  datComp.DatabaseName = UI.LibraryPath
  datComp.RecordSource = "Components"
  datComp.Refresh
  
  datMat.ReadOnly = True 'open database as read only
  datMat.DatabaseName = UI.LibraryPath
  datMat.RecordSource = "Materials"
  datMat.Refresh
  
  'clear filter combo boxes
  For i = 0 To 1
    cboFilter(i).Clear
  Next
  
  'Fill the combo box lists with the unique lists from
  'the database
  'Pad integer fields with zeros so that they sort properly
  Set DStmp = datEvap.Database.OpenRecordset("Evaporation")
  DStmp.MoveFirst
  Do Until DStmp.EOF
    cboFilter(0).AddItem DStmp.Fields(0)
    DStmp.MoveNext
  Loop
  DStmp.Close
  'the second filter comes from a different place
  Set DStmp = datComp.Database.OpenRecordset("ComponentList")
  DStmp.MoveFirst
  Do Until DStmp.EOF
    cboFilter(1).AddItem DStmp.Fields(0)
    DStmp.MoveNext
  Loop
  DStmp.Close
  
  'Add "Any" as the first item, indicating that we will not
  'filter based on this field
  For i = 0 To 1
    cboFilter(i).AddItem "Any", 0
  Next
  
  'Set combo boxes to first item
  For i = 0 To 1
    cboFilter(i).ListIndex = 0
  Next
End Sub

Private Sub SyncDisplayControls()
'Update the display controls on this form to display
'to reflect the current spray material.
  UpdateComponentList
  UpdateSpecificGravity
End Sub

Private Function TransferData() As Integer
'Transfer data from the form controls to the form properties
  Dim DsEvap As Recordset
  Set DsEvap = datEvap.Recordset  'select the current recordset

  'make sure there is a current record
  If DsEvap.BOF And DsEvap.EOF Then
    TransferData = False
    Exit Function
  End If
  
  'transfer the database record to the form properties
  SMName = Trim$(DsEvap.Fields("Substance"))
  If datMat.Recordset.NoMatch Then
    NonVGrav = 1
  Else
    NonVGrav = datMat.Recordset.Fields("Density")
  End If
  NVFrac = DsEvap.Fields("NonvolFraction")
  ACFrac = DsEvap.Fields("NonvolFraction")
  EvapRate = DsEvap.Fields("EvaporationRate")

  TransferData = True
End Function

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

  'get the subtance code for the search from the datEvap data control
  If datEvap.Recordset.BOF And datEvap.Recordset.EOF Then Exit Sub
  SubStr = datEvap.Recordset.Fields("Substance").Value

  'find the record in the datComp data control with the matching Substance
  datComp.Recordset.FindFirst "[Substance] = '" & SubStr & "'"
  If datComp.Recordset.NoMatch Then Exit Sub
  
  'get field value (all strings are packed into one)
  Set tmpFieldC = datComp.Recordset.Fields("Component")
  Set tmpFieldP = datComp.Recordset.Fields("Percent")
  Set tmpFieldN = datComp.Recordset.Fields("NumComponents")
   
  CompStr = tmpFieldC.Value  'String comtaining all components
  FieldToArray tmpFieldP, dat() 'recover array of percentages
  num = Val(tmpFieldN.Value) 'number of components

  'load up the form controls
  For i = 0 To num - 1
    s = Trim$(Mid$(CompStr, i * 32 + 1, 32))
    s = s & " (" & AGFormat$(dat(i)) & "%)"
    lstComponents.AddItem s
  Next
  
End Sub

Private Sub UpdateNumRecs()
'Ensure the validity of RecordCount by moving to the
'last record. Also update the PositionCount variable
'that is global to this form

  'if BOF and EOF are true, there are no records
  If Not datEvap.Recordset.BOF And Not datEvap.Recordset.EOF Then
    datEvap.Recordset.MoveLast
    datEvap.Recordset.MoveFirst
    PositionCount = 1
  Else
    PositionCount = 0
  End If
  UpdatePositionCount
End Sub

Private Sub UpdatePositionCount()
'update the caption of the position label
  lblPosition.Caption = Format$(PositionCount) + " of " + Format$(datEvap.Recordset.RecordCount)
End Sub

Private Sub UpdateSpecificGravity()
'find the Specific Gravity (Density) for the current
'entry (defined by the datEvap control) and fill the Label
'
  Dim Substance As String
  
  'clear the Label control
  lblSpecGrav.Caption = ""

  'get the subtance code from the datEvap data control
  If datEvap.Recordset.BOF And datEvap.Recordset.EOF Then Exit Sub
  Substance = datEvap.Recordset.Fields("Substance").Value

  'find the record in the datMat data control with the matching Substance
  datMat.Recordset.FindFirst "[Substance] = '" & Substance & "'"
  If datMat.Recordset.NoMatch Then
    lblSpecGrav.Caption = "Not available"
    Exit Sub
  End If
  
  'get field value and load form control
  lblSpecGrav.Caption = AGFormat$(datMat.Recordset.Fields("Density"))
End Sub

Public Property Get SMName() As String
  
  SMName = mName
End Property

Public Property Let SMName(ByVal vNewValue As String)

  mName = vNewValue
End Property

Public Property Get NVFrac() As Single

  NVFrac = mNVFrac
End Property

Public Property Let NVFrac(ByVal vNewValue As Single)

  mNVFrac = vNewValue
End Property

Public Property Get ACFrac() As Single

  ACFrac = mACFrac
End Property

Public Property Let ACFrac(ByVal vNewValue As Single)

  mACFrac = vNewValue
End Property

Public Property Get NonVGrav() As Single

  NonVGrav = mNonVGrav
End Property

Public Property Let NonVGrav(ByVal vNewValue As Single)

  mNonVGrav = vNewValue
End Property

Public Property Get EvapRate() As Single

  EvapRate = mEvapRate
End Property

Public Property Let EvapRate(ByVal vNewValue As Single)

  mEvapRate = vNewValue
End Property

Public Property Get Cancelled() As Boolean

  Cancelled = mCancelled
End Property

