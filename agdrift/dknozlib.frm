VERSION 5.00
Begin VB.Form frmDKNozLib 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DropKick Nozzle Library"
   ClientHeight    =   4170
   ClientLeft      =   1425
   ClientTop       =   3480
   ClientWidth     =   8760
   ForeColor       =   &H80000008&
   Icon            =   "DKNOZLIB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4170
   ScaleWidth      =   8760
   Begin VB.Data datNozzleTypMfg 
      Caption         =   "datNozzleTypMfg"
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
      Top             =   3720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data datNozzle 
      Caption         =   "datNozzle"
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
      Top             =   3720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1093
      Left            =   7800
      TabIndex        =   1
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1093
      Left            =   6840
      TabIndex        =   0
      Top             =   3720
      Width           =   855
   End
   Begin VB.Frame fraFilter 
      Caption         =   "Filter"
      Height          =   3495
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox cboNozzle 
         Height          =   315
         HelpContextID   =   1187
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nozzle:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   405
         Width           =   525
      End
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "Browse Filtered Entries"
      Height          =   3495
      Left            =   4560
      TabIndex        =   11
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmdPrev 
         Caption         =   "Prev"
         Height          =   375
         HelpContextID   =   1093
         Left            =   720
         TabIndex        =   4
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   375
         HelpContextID   =   1093
         Left            =   1320
         TabIndex        =   5
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "1st"
         Height          =   375
         HelpContextID   =   1093
         Left            =   120
         TabIndex        =   3
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   375
         HelpContextID   =   1093
         Left            =   2040
         TabIndex        =   6
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblDiameterUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   3720
         TabIndex        =   26
         Top             =   2520
         Width           =   330
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Manufacturer:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   405
      End
      Begin VB.Label lblManufacturer 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblManufacturer"
         DataField       =   "Manufacturer"
         DataSource      =   "datNozzleTypMfg"
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblType 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblType"
         DataField       =   "Type"
         DataSource      =   "datNozzleTypMfg"
         Height          =   285
         Left            =   2280
         TabIndex        =   22
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Spray Angle (deg):"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblSprayAngle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblSprayAngle"
         DataField       =   "SprayAngle"
         DataSource      =   "datNozzle"
         Height          =   285
         Left            =   2265
         TabIndex        =   20
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "V0.5"
         Height          =   195
         Left            =   270
         TabIndex        =   19
         Top             =   1515
         Width           =   405
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "D          (µm):"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label lblDiameter 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDiameter"
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblVMD 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblVMD"
         DataField       =   "VMD"
         DataSource      =   "datNozzle"
         Height          =   285
         Left            =   2280
         TabIndex        =   15
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblRelSpan 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblRelSpan"
         DataField       =   "RelSpan"
         DataSource      =   "datNozzle"
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblNozzle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNozzle"
         DataField       =   "Nozzle"
         DataSource      =   "datNozzle"
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Effective Nozzle Diameter:"
         Height          =   285
         Left            =   135
         TabIndex        =   16
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label lblPosition 
         AutoSize        =   -1  'True
         Caption         =   "0 of 0"
         Height          =   195
         Left            =   2760
         TabIndex        =   17
         Top             =   3120
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Relative Span:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1845
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nozzle:"
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   405
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmDKNozLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' $Id: dknozlib.frm,v 1.7 2001/04/26 16:21:48 tom Exp $
Dim PositionCount As Integer  'the current recordset position

Private Sub cboNozzle_Click()
  'Update the data controls recordset when the Filters change
  GetNewRecordset
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdFirst_Click()
  If datNozzle.Recordset.RecordCount > 0 Then
    datNozzle.Recordset.MoveFirst
    PositionCount = 1
  UpdatePositionCount
  End If
End Sub

Private Sub cmdLast_Click()
  If datNozzle.Recordset.RecordCount > 0 Then
    datNozzle.Recordset.MoveLast
    PositionCount = datNozzle.Recordset.RecordCount
    UpdatePositionCount
  End If
End Sub

Private Sub cmdNext_Click()
  If Not datNozzle.Recordset.EOF Then
    datNozzle.Recordset.MoveNext
    'before we update the Position, see if
    'we've moved off the last record
    If Not datNozzle.Recordset.EOF Then
      PositionCount = PositionCount + 1
      UpdatePositionCount
    Else
      datNozzle.Recordset.MovePrevious 'go back to where we were
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
  If Not datNozzle.Recordset.BOF Then
    datNozzle.Recordset.MovePrevious
    'before we update the Position, see if
    'we've moved off the first record
    If Not datNozzle.Recordset.BOF Then
      PositionCount = PositionCount - 1
      UpdatePositionCount
    Else
      datNozzle.Recordset.MoveNext 'go back to where we were
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

Private Sub datNozzle_Reposition()
  'When this data control repositions, update the
  'datNozzleTypMfg control to match
  '*and* take care of the Effective Diameter control,
  'since units must be handled
  If Not datNozzle.Recordset.BOF And _
     Not datNozzle.Recordset.EOF Then
    lblDiameter.Caption = _
      AGFormat$(UnitsDisplay( _
        datNozzle.Recordset.Fields("Diameter"), UN_SMLENGTH2))
    datNozzleTypMfg.RecordSource = _
      "select * from NozzlesTypMfg where Nozzle='" & _
      datNozzle.Recordset.Fields("Nozzle") & "'"
    datNozzleTypMfg.Refresh
  End If
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
  'from the Nozzles Table
  s = "SELECT * FROM Nozzles"
  'preview the filer fields to see if we need to add "WHERE"
  If cboNozzle.ListIndex > 0 Then
    s = s + " WHERE"
  End If
  'set a flag for adding separators
  needsep = False
  If cboNozzle.ListIndex > 0 Then
    s = s + " Nozzle= '" + cboNozzle.List(cboNozzle.ListIndex) + "'"
    needsep = True
  End If
  
  'Reset the RecordSource property and refresh to get the
  'new recordset
  datNozzle.RecordSource = s
  datNozzle.Refresh
  'Update the display for number of records in this set
  UpdateNumRecs
End Sub

Private Sub InitForm()
'Initialize this form and its controls
  Dim DStmp As Recordset
  Dim c As Control
  Dim s As String
  Dim dslist As String

  'This form uses the Tag property to return status
  'true for success, false for failure
  'set to "false" by default
  Me.Tag = "False"

  'Center the form
  CenterForm Me

  'units
  lblDiameterUnits.Caption = UnitsName(UN_SMLENGTH2)
  
  'set the database and initial query for the data controls
  datNozzle.ReadOnly = True 'open database as read only
  datNozzle.DatabaseName = UI.LibraryPath
  datNozzle.RecordSource = "Nozzles"
  datNozzle.Refresh
  UpdateNumRecs

  datNozzleTypMfg.ReadOnly = True 'open database as read only
  datNozzleTypMfg.DatabaseName = UI.LibraryPath
  datNozzleTypMfg.RecordSource = "NozzlesTypMfg"
  datNozzleTypMfg.Refresh

  'clear filter combo boxes
  cboNozzle.Clear

  'Fill the combo box lists with the unique lists from
  'the database
  'Pad integer fields with zeros so that they sort properly
  dslist = "Nozzles"
  Set DStmp = datNozzle.Database.OpenRecordset(dslist)
  DStmp.MoveFirst
  Do Until DStmp.EOF
    cboNozzle.AddItem DStmp.Fields(0)
    DStmp.MoveNext
  Loop
  DStmp.Close
  
  'Add "Any" as the first item, indicating that we will not
  'filter based on this field
  cboNozzle.AddItem "Any", 0
  
  'Set combo boxes to first item
  cboNozzle.ListIndex = 0
End Sub

Private Function TransferData() As Integer
'Transfer the selected data to the parent form
  Dim fm As Form
  Dim g As Control
  Dim DS As Recordset

  Set fm = frmDropKick 'select the parent form
  Set DS = datNozzle.Recordset  'select the current recordset
  
  'make sure there is a current record
  If DS.BOF And DS.EOF Then
    TransferData = False
    Exit Function
  End If
  
  'transfer the distribution from the database to the output controls
  fm.lblLibNozzle.Caption = Trim$(DS.Fields("Nozzle"))
  fm.txtVMD.Text = AGFormat$(DS.Fields("VMD"))
  fm.txtRelSpan.Text = AGFormat$(DS.Fields("RelSpan"))
  fm.txtDiam.Text = AGFormat$(UnitsDisplay(DS.Fields("Diameter"), UN_SMLENGTH2))
  fm.txtSprayAngle.Text = AGFormat$(DS.Fields("SprayAngle"))
  
  TransferData = True
End Function

Private Sub UpdateNumRecs()
'Ensure the validity of RecordCount by moving to the
'last record. Also update the PositionCount variable
'that is global to this form

  'if BOF and EOF are true, there are no records
  If Not datNozzle.Recordset.BOF And Not datNozzle.Recordset.EOF Then
    datNozzle.Recordset.MoveLast
    datNozzle.Recordset.MoveFirst
    PositionCount = 1
  Else
    PositionCount = 0
  End If
  UpdatePositionCount
End Sub

Private Sub UpdatePositionCount()
'update the caption of the position label
  lblPosition.Caption = Format$(PositionCount) + " of " + Format$(datNozzle.Recordset.RecordCount)
End Sub

