VERSION 5.00
Begin VB.Form frmDKMatLib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DropKick Material Library"
   ClientHeight    =   4905
   ClientLeft      =   450
   ClientTop       =   1350
   ClientWidth     =   8760
   ForeColor       =   &H80000008&
   Icon            =   "DKMATLIB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4905
   ScaleWidth      =   8760
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
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data datMat 
      Caption         =   "datMat"
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
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1095
      Left            =   7800
      TabIndex        =   1
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1095
      Left            =   6840
      TabIndex        =   0
      Top             =   4440
      Width           =   855
   End
   Begin VB.Frame fraFilter 
      Caption         =   "Filter"
      Height          =   4215
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   3615
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1073
         Index           =   1
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   1259
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Component:"
         Height          =   195
         Left            =   105
         TabIndex        =   20
         Top             =   405
         Width           =   1020
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Substance:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   765
         Width           =   975
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Browse Filtered Entries"
      Height          =   4215
      Left            =   3840
      TabIndex        =   15
      Top             =   120
      Width           =   4815
      Begin VB.ListBox lstComponents 
         Height          =   1230
         HelpContextID   =   1095
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   4575
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Prev"
         Height          =   375
         HelpContextID   =   1095
         Left            =   720
         TabIndex        =   5
         Top             =   3720
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   375
         HelpContextID   =   1095
         Left            =   1320
         TabIndex        =   6
         Top             =   3720
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "1st"
         Height          =   375
         HelpContextID   =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   3720
         Width           =   615
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   375
         HelpContextID   =   1095
         Left            =   2040
         TabIndex        =   7
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblElongVisc 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblElongVisc"
         DataField       =   "ElongVisc"
         DataSource      =   "datMat"
         Height          =   285
         Left            =   2520
         TabIndex        =   21
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblShearVisc 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblShearVisc"
         DataField       =   "ShearVisc"
         DataSource      =   "datMat"
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblDynSurfTens 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDynSurfTens"
         DataField       =   "DynSurfTens"
         DataSource      =   "datMat"
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblDensity 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDensity"
         DataField       =   "Density"
         DataSource      =   "datMat"
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblSubst 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblSubst"
         DataField       =   "Substance"
         DataSource      =   "datMat"
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Elongational Viscosity (cp):"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1485
         Width           =   2325
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Components:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1110
      End
      Begin VB.Label lblPosition 
         AutoSize        =   -1  'True
         Caption         =   "0 of 0"
         Height          =   195
         Left            =   2760
         TabIndex        =   23
         Top             =   3840
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Shear Viscosity (cp):"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1125
         Width           =   1770
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Dynamic Surface Tension: (dynes/cm)"
         Height          =   435
         Left            =   120
         TabIndex        =   18
         Top             =   660
         Width           =   1905
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Specific Gravity:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1845
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Substance:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   405
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmDKMatLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: dkmatlib.frm,v 1.4 2001/04/26 16:21:48 tom Exp $
Dim PositionCount As Integer  'the current recordset position

Private Sub cboFilter_Click(Index As Integer)
  'Update the data controls recordset when the Filters change
  GetNewRecordset
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdFirst_Click()
  If datMat.Recordset.RecordCount > 0 Then
    datMat.Recordset.MoveFirst
    PositionCount = 1
  UpdatePositionCount
  UpdateComponentList
  End If
End Sub

Private Sub cmdLast_Click()
  If datMat.Recordset.RecordCount > 0 Then
    datMat.Recordset.MoveLast
    PositionCount = datMat.Recordset.RecordCount
    UpdatePositionCount
    UpdateComponentList
  End If
End Sub

Private Sub cmdNext_Click()
  If Not datMat.Recordset.EOF Then
    datMat.Recordset.MoveNext
    'before we update the Position, see if
    'we've moved off the last record
    If Not datMat.Recordset.EOF Then
      PositionCount = PositionCount + 1
      UpdatePositionCount
      UpdateComponentList
    Else
      datMat.Recordset.MovePrevious 'go back to where we were
    End If
  End If
End Sub

Private Sub cmdOk_Click()
  If TransferData() Then
    Me.Tag = "True"  'indicate success
    Me.Hide
  Else
    s$ = "No library entry has been selected."
    t% = vbCritical + vbOKOnly
    MsgBox s$, t%
  End If
End Sub

Private Sub cmdPrev_Click()
  If Not datMat.Recordset.BOF Then
    datMat.Recordset.MovePrevious
    'before we update the Position, see if
    'we've moved off the first record
    If Not datMat.Recordset.BOF Then
      PositionCount = PositionCount - 1
      UpdatePositionCount
      UpdateComponentList
    Else
      datMat.Recordset.MoveNext 'go back to where we were
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
  s = "SELECT * FROM Materials"
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
  '"Any" or "xxx Min/Max", we must build an additional query section from
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
  datMat.RecordSource = s
  datMat.Refresh
  'Update the display for number of records in this set
  UpdateNumRecs
  'after moving to a new record, update the component list
  UpdateComponentList
End Sub

Private Sub InitForm()
'Initialize this form and its controls
  Dim DStmp As Recordset
  Dim c As Control
  Dim s As String
  ReDim dslist(1) As String

  'This form uses the Tag property to return status
  'true for success, false for failure
  'set to "false by default
  Me.Tag = "False"

  'Center the form
  CenterForm Me

  'set the database and initial query for the data controls
  datMat.ReadOnly = True 'open database as read only
  datMat.DatabaseName = UI.LibraryPath
  datMat.RecordSource = "Materials"
  datMat.Refresh
  UpdateNumRecs

  datComp.ReadOnly = True 'open database as read only
  datComp.DatabaseName = UI.LibraryPath
  datComp.RecordSource = "Components"
  datComp.Refresh
  
  'clear filter combo boxes
  For i = 0 To 1
    cboFilter(i).Clear
  Next
  
  'Fill the combo box lists with the unique lists from
  'the database
  'Pad integer fields with zeros so that they sort properly
  dslist(0) = "Materials"
  For ids = 0 To 0
    Set DStmp = datMat.Database.OpenRecordset(dslist(ids))
    DStmp.MoveFirst
    Do Until DStmp.EOF
      If DStmp.Fields(0).Type = dbInteger Then
        cboFilter(ids).AddItem Format$(DStmp.Fields(0), "@@@")
      Else
        cboFilter(ids).AddItem DStmp.Fields(0)
      End If
      DStmp.MoveNext
    Loop
    DStmp.Close
  Next
  'the second filter comes from a different place
  ids = 1
  dslist(ids) = "ComponentList"
  Set DStmp = datComp.Database.OpenRecordset(dslist(ids))
  DStmp.MoveFirst
  Do Until DStmp.EOF
    cboFilter(ids).AddItem DStmp.Fields(0)
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

Private Function TransferData() As Integer
'Transfer the selected data to the parent form
  Dim fm As Form
  Dim DS As Recordset
  Dim Subst As String
  Dim DynSurfTens As Single
  Dim ShearVisc As Single
  Dim Density As Single
  Dim ElongVisc As Single

  Set fm = frmDropKick 'select the parent form
  Set DS = datMat.Recordset  'select the current recordset
  
  'make sure there is a current record
  If DS.BOF And DS.EOF Then
    TransferData = False
    Exit Function
  End If
  
  'transfer the database record to temp storage
  Subst = Trim$(DS.Fields("Substance"))
  Density = DS.Fields("Density")
  DynSurfTens = DS.Fields("DynSurfTens")
  ShearVisc = DS.Fields("ShearVisc")
  ElongVisc = DS.Fields("ElongVisc")

  'transfer the distribution to the output controls
  fm.lblLibMaterial.Caption = Subst
  fm.txtDynSurfTens.Text = AGFormat$(DynSurfTens)
  fm.txtShearVisc.Text = AGFormat$(ShearVisc)
  fm.txtDensity.Text = AGFormat$(Density)
  fm.txtElongVisc.Text = AGFormat$(ElongVisc)
  
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

  'get the subtance code for the search from the datMat data control
  If datMat.Recordset.BOF And datMat.Recordset.EOF Then Exit Sub
  SubStr = datMat.Recordset.Fields("Substance").Value

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
  If Not datMat.Recordset.BOF And Not datMat.Recordset.EOF Then
    datMat.Recordset.MoveLast
    datMat.Recordset.MoveFirst
    PositionCount = 1
  Else
    PositionCount = 0
  End If
  UpdatePositionCount
End Sub

Private Sub UpdatePositionCount()
'update the caption of the position label
  lblPosition.Caption = Format$(PositionCount) + " of " + Format$(datMat.Recordset.RecordCount)
End Sub

