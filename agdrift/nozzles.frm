VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmNozzles 
   Caption         =   "Nozzles"
   ClientHeight    =   6675
   ClientLeft      =   1035
   ClientTop       =   1545
   ClientWidth     =   9480
   HelpContextID   =   1185
   Icon            =   "NOZZLES.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6675
   ScaleWidth      =   9480
   Begin VB.Frame fraProp 
      Caption         =   "Nozzle Installation Properties"
      Height          =   975
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   9375
      Begin VB.Label lblACName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aircraft Name"
         Height          =   255
         Left            =   720
         TabIndex        =   60
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblACNameLbl 
         Caption         =   "Aircraft:"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblLimitLabel 
         AutoSize        =   -1  'True
         Caption         =   "Right"
         Height          =   195
         Index           =   2
         Left            =   3840
         TabIndex        =   38
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblLimitLabel 
         AutoSize        =   -1  'True
         Caption         =   "Left"
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   37
         Top             =   600
         Width           =   270
      End
      Begin VB.Label lblLimitUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   1
         Left            =   5280
         TabIndex        =   36
         Top             =   600
         Width           =   120
      End
      Begin VB.Label lblLimit 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   35
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblLimit 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   34
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblLimitLabel 
         Caption         =   "Nozzle Distribution Extent"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblLimitUnits 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   32
         Top             =   600
         Width           =   120
      End
      Begin VB.Label lblSemiSpanUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   5280
         TabIndex        =   31
         Top             =   240
         Width           =   330
      End
      Begin VB.Label lblSemiSpan 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SemiSpan"
         Height          =   255
         Left            =   4320
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblSemiSpanLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Semispan"
         Height          =   195
         Left            =   3480
         TabIndex        =   29
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      HelpContextID   =   1185
      Left            =   7800
      TabIndex        =   0
      Top             =   6120
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   1185
      Left            =   8520
      TabIndex        =   1
      Top             =   6120
      Width           =   855
   End
   Begin VB.Frame fraView 
      Caption         =   "Nozzles"
      Height          =   4335
      Left            =   0
      TabIndex        =   27
      Top             =   960
      Width           =   9375
      Begin VB.CommandButton cmdGenRegDist 
         Caption         =   "Generate Regular Distribution"
         Height          =   255
         HelpContextID   =   1480
         Left            =   3240
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "Graphical View"
         Height          =   255
         HelpContextID   =   1185
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optDisplay 
         Caption         =   "Tabular View"
         Height          =   255
         HelpContextID   =   1185
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.PictureBox picTabular 
         Height          =   3615
         HelpContextID   =   1185
         Left            =   5880
         ScaleHeight     =   3555
         ScaleWidth      =   3315
         TabIndex        =   46
         Top             =   600
         Width           =   3375
         Begin MSFlexGridLib.MSFlexGrid grdNozzle 
            Height          =   1815
            Left            =   1680
            TabIndex        =   65
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   3201
            _Version        =   393216
            Cols            =   5
            WordWrap        =   -1  'True
            Appearance      =   0
         End
         Begin VB.TextBox txtEdit 
            BorderStyle     =   0  'None
            Height          =   255
            HelpContextID   =   1185
            Left            =   1200
            TabIndex        =   17
            Text            =   "grid edit text box"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmdDeleteTable 
            Caption         =   "&Delete"
            Height          =   255
            HelpContextID   =   1185
            Left            =   840
            TabIndex        =   14
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdAddTable 
            Caption         =   "Add"
            Height          =   255
            HelpContextID   =   1185
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "Im&port"
            Height          =   255
            HelpContextID   =   1185
            Left            =   480
            TabIndex        =   16
            Top             =   1080
            Width           =   615
         End
         Begin VB.CommandButton cmdSort 
            Caption         =   "Sort"
            Height          =   255
            HelpContextID   =   1185
            Left            =   480
            TabIndex        =   15
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblNumNozLabel 
            AutoSize        =   -1  'True
            Caption         =   "Nozzles:"
            Height          =   195
            Left            =   1440
            TabIndex        =   48
            Top             =   2160
            Width           =   600
         End
         Begin VB.Label lblNumNoz 
            AutoSize        =   -1  'True
            Caption         =   "NumNoz"
            Height          =   195
            Left            =   2160
            TabIndex        =   47
            Top             =   2160
            Width           =   615
         End
      End
      Begin VB.PictureBox picGraphical 
         Height          =   3615
         Left            =   120
         ScaleHeight     =   3555
         ScaleWidth      =   3315
         TabIndex        =   39
         Top             =   600
         Width           =   3375
         Begin VB.CommandButton cmdDeletePicture 
            Caption         =   "Del"
            Height          =   315
            HelpContextID   =   1185
            Left            =   1680
            TabIndex        =   12
            Top             =   3120
            Width           =   495
         End
         Begin VB.CommandButton cmdAddPicture 
            Caption         =   "Add"
            Height          =   315
            HelpContextID   =   1185
            Left            =   1080
            TabIndex        =   11
            Top             =   3120
            Width           =   495
         End
         Begin VB.CommandButton cmdZoomFit 
            Caption         =   "Fit"
            Height          =   315
            HelpContextID   =   1185
            Left            =   1800
            TabIndex        =   8
            Top             =   2640
            Width           =   315
         End
         Begin VB.CommandButton cmdZoomOut 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            HelpContextID   =   1185
            Left            =   1440
            TabIndex        =   7
            Top             =   2640
            Width           =   315
         End
         Begin VB.CommandButton cmdZoomIn 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            HelpContextID   =   1185
            Left            =   1080
            TabIndex        =   6
            Top             =   2640
            Width           =   315
         End
         Begin VB.PictureBox picViewPort 
            Height          =   2295
            HelpContextID   =   1185
            Left            =   120
            ScaleHeight     =   2235
            ScaleWidth      =   2595
            TabIndex        =   40
            Top             =   240
            Width           =   2655
            Begin VB.PictureBox picViewArea 
               Height          =   1935
               HelpContextID   =   1185
               Left            =   120
               ScaleHeight     =   1875
               ScaleWidth      =   2355
               TabIndex        =   41
               Top             =   120
               Width           =   2415
               Begin VB.PictureBox picNozzle 
                  Appearance      =   0  'Flat
                  BackColor       =   &H000000FF&
                  BorderStyle     =   0  'None
                  DrawWidth       =   40
                  ForeColor       =   &H80000008&
                  Height          =   100
                  HelpContextID   =   1185
                  Index           =   0
                  Left            =   120
                  ScaleHeight     =   105
                  ScaleWidth      =   105
                  TabIndex        =   42
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   100
               End
               Begin VB.Line linCenterLineH 
                  BorderStyle     =   3  'Dot
                  X1              =   960
                  X2              =   1440
                  Y1              =   240
                  Y2              =   240
               End
               Begin VB.Shape shpSelect 
                  BorderStyle     =   3  'Dot
                  Height          =   495
                  Left            =   840
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.Label lblDown 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Down"
                  Height          =   195
                  Left            =   1320
                  TabIndex        =   51
                  Top             =   480
                  Width           =   420
               End
               Begin VB.Label lblUp 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Up"
                  Height          =   195
                  Left            =   1320
                  TabIndex        =   50
                  Top             =   0
                  Width           =   210
               End
               Begin VB.Label lblHilite 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hilite info"
                  Height          =   195
                  Left            =   720
                  TabIndex        =   49
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   645
               End
               Begin VB.Label lblCenter 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Center"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   720
                  TabIndex        =   45
                  Top             =   240
                  Width           =   570
               End
               Begin VB.Label lblRight 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Right"
                  Height          =   195
                  Left            =   1920
                  TabIndex        =   44
                  Top             =   240
                  Width           =   375
               End
               Begin VB.Label lblLeft 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Left"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   43
                  Top             =   240
                  Width           =   270
               End
               Begin VB.Line linCenterLineV 
                  BorderStyle     =   3  'Dot
                  X1              =   1200
                  X2              =   1200
                  Y1              =   120
                  Y2              =   720
               End
               Begin VB.Line linBoom 
                  BorderWidth     =   2
                  X1              =   240
                  X2              =   2160
                  Y1              =   480
                  Y2              =   480
               End
               Begin VB.Shape shpWing 
                  FillColor       =   &H00FFFFFF&
                  FillStyle       =   0  'Solid
                  Height          =   375
                  Left            =   1680
                  Top             =   960
                  Width           =   375
               End
            End
         End
         Begin VB.ComboBox cboView 
            Height          =   315
            HelpContextID   =   1185
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2640
            Width           =   975
         End
         Begin VB.HScrollBar hscViewPort 
            Height          =   315
            HelpContextID   =   1185
            LargeChange     =   20
            Left            =   2280
            Max             =   99
            TabIndex        =   9
            Top             =   2640
            Width           =   975
         End
         Begin VB.VScrollBar vscViewPort 
            Height          =   1335
            HelpContextID   =   1185
            LargeChange     =   20
            Left            =   2880
            Max             =   99
            TabIndex        =   10
            Top             =   840
            Width           =   315
         End
      End
   End
   Begin VB.Frame fraNozLoc 
      Caption         =   "Nozzle Location"
      Height          =   1335
      Left            =   0
      TabIndex        =   52
      Top             =   5280
      Width           =   2415
      Begin VB.TextBox txtNozParamH 
         DataField       =   "TypSpeed"
         DataSource      =   "Data1"
         Height          =   285
         HelpContextID   =   1530
         Left            =   960
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtNozParamV 
         DataField       =   "TypSpeed"
         DataSource      =   "Data1"
         Height          =   285
         HelpContextID   =   1531
         Left            =   960
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtNozParamF 
         DataField       =   "TypSpeed"
         DataSource      =   "Data1"
         Height          =   285
         HelpContextID   =   1532
         Left            =   960
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblNozParamHUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   1800
         TabIndex        =   58
         Top             =   240
         Width           =   330
      End
      Begin VB.Label lblNozParamH 
         Alignment       =   2  'Center
         Caption         =   "Horizontal"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblNozParamVUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   1800
         TabIndex        =   56
         Top             =   600
         Width           =   330
      End
      Begin VB.Label lblNozParamV 
         Alignment       =   2  'Center
         Caption         =   "Vertical"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblNozParamFUnits 
         AutoSize        =   -1  'True
         Caption         =   "units"
         Height          =   195
         Left            =   1800
         TabIndex        =   54
         Top             =   960
         Width           =   330
      End
      Begin VB.Label lblNozParamF 
         Alignment       =   2  'Center
         Caption         =   "Forward"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame fraNozDSD 
      Caption         =   "Nozzle Drop Size Distribution"
      Height          =   1335
      Left            =   2520
      TabIndex        =   61
      Top             =   5280
      Width           =   5175
      Begin VB.CommandButton cmdEditDrop 
         Caption         =   "Edit"
         Height          =   255
         HelpContextID   =   1100
         Index           =   2
         Left            =   1080
         TabIndex        =   26
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdEditDrop 
         Caption         =   "Edit"
         Height          =   255
         HelpContextID   =   1100
         Index           =   1
         Left            =   1080
         TabIndex        =   24
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdEditDrop 
         Caption         =   "Edit"
         Height          =   255
         HelpContextID   =   1100
         Index           =   0
         Left            =   1080
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optDSD 
         BackColor       =   &H000000FF&
         Caption         =   "DSD 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optDSD 
         BackColor       =   &H0000FF00&
         Caption         =   "DSD 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optDSD 
         BackColor       =   &H00FF0000&
         Caption         =   "DSD 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblDSDdesc 
         AutoSize        =   -1  'True
         Caption         =   "DSD description"
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   64
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblDSDdesc 
         AutoSize        =   -1  'True
         Caption         =   "DSD description"
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   63
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lblDSDdesc 
         AutoSize        =   -1  'True
         Caption         =   "DSD description"
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   62
         Top             =   360
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmNozzles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'$Id: nozzles.frm,v 1.12 2003/04/02 19:08:09 tom Exp $
Option Explicit

'this flag is used to tell the option buttons not to
'take action on their new values. This is required
'to differentiate between programatic state changes
'and user actions
Dim OptTakeAction As Integer  'if true, execute automatic change-related code
                              'for option button
Dim PropTakeAction As Integer 'if true, execute automatic change-related code
                              'for Property text boxes

Dim SaveType As Integer       'place to save type

'Current Nozzles - selected nozles for highlighting,
'moving, removing, etc.
Dim NumCurrentNozzles As Integer  'number of selected nozzles
Dim CurrentNozzles(MAX_NOZZLES - 1) As Integer 'indices of selected nozzles

'NumNozzleControls keeps track of the number
'of controls that are defined to simulate
'nozzles
Dim NumNozzleControls As Integer

'The following variables are used during
'dragging operations and correct for the pointer's
'position within the dragged control
Private Const DRAGSTARTDIST2 As Single = 250 'Square of min dist for starting a drag
Private Dragging As Boolean
Private DragOffsetX As Single
Private DragOffsetY As Single

'Selection box vars
Dim SelStartX
Dim SelStartY

'grid editing vars
'gRow/gCol provide a means of remembering where
'focus was in the grid, even when the focus
'changes suddenly, in response to clicks on other
'cells.
'CancelCellEdit provides a means to cancel an
'edit operation in progress for the grid. Setting
'this value to true during an edit will prevent
'any changes from becoming current.
'CellEdited provides a means of knowing if the
'contents of the current cell have been changed.
'This allows txtEdit.LostFocus to update the
'DataChangedFlag only when appropriate
Dim gRow As Integer
Dim gCol As Integer
Dim CancelCellEdit As Integer
Dim CellEdited As Integer

'An escape key means Cancel. Press it and the form goes
'away. Normally, we would set the Cancel property to True
'for the Cancel button and we would be done, but this form
'contains EditGrids. EditGrids rely on the Escape key to
'cancel an edit. If the Cancel property of the Cancel
'button is true, this behavior doesn't work. The desired
'behavior is for the Escape key to cancel an edit operation
'in an EditGrid, and to dismiss the form in all other cases.
'To that end, we employ this method:
'- Set the Cancel property to False for the Cancel button
'- Set the KeyPreview property to True for the form
'- Examine KeyPress events at the form level and pass Escapes
'  through to EditGrid text boxes, and dismiss the form for
'  all other cases.
'Here we define a collection to hold all EditGrid text boxes
'for this form. If, when an escape key is pressed, one of
'the controls in this collection is the ActiveControl, the
'program continues normally and the Text control receives a
'KeyPress event. If the ActiveControl is not one in the
'collection, the cmdCancel_Click event routine is invoked.
'See Form_KeyPress below.
Private ControlsThatMayReceiveEscape As New Collection

Private Sub UpdateNozzleCoords()
'Convert the coordinates of the currently selected nozzle(s)
'picture controls to real-world coordinates and store
'them in UserData
  
  Dim MetersPerTwip As Single
  Dim RefHorizTwip As Single
  Dim RefVertTwip As Single
  Dim PosHorizTwip As Single
  Dim PosVertTwip As Single
  Dim i As Integer
  Dim ic As Integer

  If NumCurrentNozzles > 0 Then
    'conversions
    MetersPerTwip = (UD.AC.SemiSpan * 2) / (linBoom.X2 - linBoom.X1)
    RefHorizTwip = (linBoom.X1 + linBoom.X2) * 0.5
    RefVertTwip = linBoom.Y1 + linBoom.BorderWidth
    
    For i = 0 To NumCurrentNozzles - 1
      ic = CurrentNozzles(i) 'corresponding control index
    
      'convert position to image units
      PosHorizTwip = picNozzle(ic).Left + (0.5 * picNozzle(ic).Width) - RefHorizTwip
      PosVertTwip = RefVertTwip - (picNozzle(ic).Top + (0.5 * picNozzle(ic).Height))
      'position the nozzle control
      NZ2.PosHoriz(ic) = PosHorizTwip * MetersPerTwip
      If cboView.ListIndex = 0 Then 'Rear view
        NZ2.PosVert(ic) = PosVertTwip * MetersPerTwip
      Else 'top view
        NZ2.PosFwd(ic) = PosVertTwip * MetersPerTwip
      End If
    Next
  End If
End Sub

Private Sub HighlightNozzles()
  Dim i As Integer
  Dim ic As Integer
  
  'determine selection status of all nozzles
  ReDim NZsel(NZ2.NumNoz - 1) 'redim and set to false (0)
  For i = 0 To NumCurrentNozzles - 1
    NZsel(CurrentNozzles(i)) = True
  Next
  
  'loop through all nozzles
  For ic = 0 To NZ2.NumNoz - 1
    'update the highlight condition
    HighlightNozzle ic, NZsel(ic)
  Next
End Sub

Private Sub HighlightNozzle(ic, hilite)
'Add/remove highlight features to/for specified nozzle
  Dim i As Integer
  
  Select Case NZ2.NozType(ic)
  Case 0
    If Not hilite Then
      picNozzle(ic).BackColor = RGB(255, 0, 0)
    Else
      picNozzle(ic).BackColor = RGB(255, 127, 127)
    End If
  Case 1
    If Not hilite Then
      picNozzle(ic).BackColor = RGB(0, 255, 0)
    Else
      picNozzle(ic).BackColor = RGB(127, 255, 127)
    End If
  Case 2
    If Not hilite Then
      picNozzle(ic).BackColor = RGB(0, 0, 255)
    Else
      picNozzle(ic).BackColor = RGB(127, 127, 255)
    End If
  End Select
  
End Sub

Private Sub UpdateNozzleParamControls()
'Adjust the Nozzle Parameter Controls to conform to the current settings
  Dim DispHoriz As Boolean
  Dim DispVert As Boolean
  Dim DispFwd As Boolean
  Dim DispDSD As Boolean
  Dim i As Integer
  
  'init display flags
  DispHoriz = True
  DispVert = True
  DispFwd = True
  DispDSD = True
  
  Select Case NumCurrentNozzles
  Case 0 'No Nozzles selected
    DispHoriz = False
    DispVert = False
    DispFwd = False
    DispDSD = False
  Case 1 'One Nozzle Selected
  Case Else 'More than one nozzle selected
    For i = 1 To NumCurrentNozzles - 1
      If NZ2.PosHoriz(CurrentNozzles(0)) <> _
         NZ2.PosHoriz(CurrentNozzles(i)) Then DispHoriz = False
      If NZ2.PosVert(CurrentNozzles(0)) <> _
         NZ2.PosVert(CurrentNozzles(i)) Then DispVert = False
      If NZ2.PosFwd(CurrentNozzles(0)) <> _
         NZ2.PosFwd(CurrentNozzles(i)) Then DispFwd = False
      If NZ2.NozType(CurrentNozzles(0)) <> _
         NZ2.NozType(CurrentNozzles(i)) Then DispDSD = False
    Next
  End Select
  
  If DispHoriz Then
    txtNozParamH.Text = AGFormat$(UnitsDisplay(NZ2.PosHoriz(CurrentNozzles(0)), UN_LENGTH))
  Else
    txtNozParamH.Text = ""
  End If
  If DispVert Then
    txtNozParamV.Text = AGFormat$(UnitsDisplay(NZ2.PosVert(CurrentNozzles(0)), UN_LENGTH))
  Else
    txtNozParamV.Text = ""
  End If
  If DispFwd Then
    txtNozParamF.Text = AGFormat$(UnitsDisplay(NZ2.PosFwd(CurrentNozzles(0)), UN_LENGTH))
  Else
    txtNozParamF.Text = ""
  End If
  If DispDSD Then
    optDSD(NZ2.NozType(CurrentNozzles(0))).Value = True
  Else
    For i = 0 To 2
      optDSD(i).Value = False
    Next
  End If
End Sub

Private Sub DataToForm()
'transfer local data to form controls for editing
  Dim temp As Integer
  Dim i As Integer
  Dim tmpVisible As Boolean
  Dim ic As Integer
  Dim n As Integer
  
  UpdateDistributionExtent
  
  'Drop size
  For i = 0 To 2: UpdateDSDTypeLabel (i): Next
  
  'Nozzles
  If optDisplay(0) Then 'Graphical
    tmpVisible = picViewPort.Visible
    picViewPort.Visible = False
    'Boom Rear View
    For i = 0 To NZ2.NumNoz - 1
      If i + 1 > NumNozzleControls Then
        Load picNozzle(i)
        NumNozzleControls = NumNozzleControls + 1
      End If
      picNozzle(i).Visible = True
    Next
    'Turn off any extra nozzle controls.
    'We can't unload them, because this routine may be called from
    'within an inappropriate context, such as the Click method of
    'an option menu.
    For i = NZ2.NumNoz To NumNozzleControls - 1
      picNozzle(i).Visible = False
    Next
    
    'update the Selection list
    ic = 0
    n = NumCurrentNozzles
    While ic < n
      If CurrentNozzles(ic) > NZ2.NumNoz - 1 Then
        For i = ic To NumCurrentNozzles - 2
          CurrentNozzles(i) = CurrentNozzles(i + 1)
        Next
        NumCurrentNozzles = NumCurrentNozzles - 1
      End If
      ic = ic + 1
    Wend
    
    'Draw the nozzle picture controls
    DrawNozzlePicture
    
    picViewPort.Visible = tmpVisible
  
  ElseIf optDisplay(1) Then 'tabular
    ArrayToGrid NZ2.NumNoz, NZ2.NozType(), NZ2.PosHoriz(), NZ2.PosVert(), NZ2.PosFwd()
    lblNumNoz.Caption = Format$(NZ2.NumNoz)
  End If
End Sub

Private Sub UpdateNozzlePictureCurrent()
'Position Nozzle Picture controls of currently selected
'nozzles to match the locations specified in UserData

  Dim i As Integer
  Dim ic As Integer
  Dim TwipsPerMeter As Single
  Dim PosTwip As Single
  Dim RefHorizTwip As Single
  Dim RefVertTwip As Single
  Dim PosHorizTwip As Single
  Dim PosVertTwip As Single
  
  If NumCurrentNozzles > 0 Then
    'conversion
    TwipsPerMeter = (linBoom.X2 - linBoom.X1) / (UD.AC.SemiSpan * 2)
    RefHorizTwip = (linBoom.X1 + linBoom.X2) * 0.5
    RefVertTwip = linBoom.Y1
    For i = 0 To NumCurrentNozzles - 1
      'recall current nozzle
      ic = CurrentNozzles(i)
      'convert position to image units
      PosHorizTwip = RefHorizTwip + (NZ2.PosHoriz(ic) * TwipsPerMeter)
      If cboView.ListIndex = 0 Then 'rear view
        PosVertTwip = RefVertTwip - (NZ2.PosVert(ic) * TwipsPerMeter)
      ElseIf cboView.ListIndex = 1 Then 'top view
        PosVertTwip = RefVertTwip - (NZ2.PosFwd(ic) * TwipsPerMeter)
      End If
      'position the nozzle control
      picNozzle(ic).Left = PosHorizTwip - (0.5 * picNozzle(ic).Width)
      picNozzle(ic).Top = PosVertTwip - (0.5 * picNozzle(ic).Height)
      'update the highlight condition
      HighlightNozzle ic, True
    Next
  End If
End Sub

Private Sub UpdateNozzlePictureAll()
'Position Nozzle Picture controls of all
'nozzles to match the locations specified in UserData

  Dim i As Integer
  Dim ic As Integer
  Dim PosTwip As Single
  Dim TwipsPerMeter As Single
  Dim RefHorizTwip As Single
  Dim RefVertTwip As Single
  Dim PosHorizTwip As Single
  Dim PosVertTwip As Single
  
  'conversion
  TwipsPerMeter = (linBoom.X2 - linBoom.X1) / (UD.AC.SemiSpan * 2)
  RefHorizTwip = (linBoom.X1 + linBoom.X2) * 0.5
  RefVertTwip = linBoom.Y1
  
  'determine selection status of all nozzles
  ReDim NZsel(NZ2.NumNoz - 1) 'redim and set to false (0)
  For i = 0 To NumCurrentNozzles - 1
    NZsel(CurrentNozzles(i)) = True
  Next
  
  'loop through all nozzles
  For ic = 0 To NZ2.NumNoz - 1
    'convert position to image units
    PosHorizTwip = RefHorizTwip + (NZ2.PosHoriz(ic) * TwipsPerMeter)
    If cboView.ListIndex = 0 Then 'rear view
      PosVertTwip = RefVertTwip - (NZ2.PosVert(ic) * TwipsPerMeter)
    ElseIf cboView.ListIndex = 1 Then 'top view
      PosVertTwip = RefVertTwip - (NZ2.PosFwd(ic) * TwipsPerMeter)
    End If
    'position the nozzle control
    picNozzle(ic).Left = PosHorizTwip - (0.5 * picNozzle(ic).Width)
    picNozzle(ic).Top = PosVertTwip - (0.5 * picNozzle(ic).Height)
    'update the highlight condition
    HighlightNozzle ic, NZsel(ic)
  Next
End Sub

Private Sub ChangeType(NewType As Integer)
'Select a new Distribution Type and do what is necessary to
'get new data
  NZ2.Type = NewType
  'if returning to basic, reread the distribution
  If NewType = 0 Then
    'recover Basic parameters
    GetBasicDataNZ NZ2.BasicType, NZ2
    AdjustBasicNozzles NZ2.BoomWidth, UD.AC.SemiSpan, NZ2
    DataToForm
  End If
End Sub

Private Sub FormToData()
'Place the form data in user data storage
  Dim lnum As Long
  ReDim tmph(59) As Single
  Dim tmpv As Single
  Dim tmpf As Single
  Dim i As Integer

  Dim HorizMax As Single
  Dim dh As Single
  Dim inz As Integer
  
  'sort the nozzles, just in case
  SortNozzles
  
  'copy local storage into user data
  UD.NZ.Type = NZ2.Type
  UD.NZ.BasicType = NZ2.BasicType
  UD.NZ.Name = NZ2.Name
  UD.NZ.LName = NZ2.LName
  UD.NZ.NumNoz = NZ2.NumNoz
  For i = 0 To NZ2.NumNoz - 1
    UD.NZ.NozType(i) = NZ2.NozType(i)
    UD.NZ.PosHoriz(i) = NZ2.PosHoriz(i)
    UD.NZ.PosVert(i) = NZ2.PosVert(i)
    UD.NZ.PosFwd(i) = NZ2.PosFwd(i)
  Next
  UD.NZ.PosHorizLimit = NZ2.PosHorizLimit
  UD.NZ.BoomWidth = NZ2.BoomWidth
  
  UpdateDataChangedFlag True 'Data was changed
  UC.Valid = False 'Calcs need to be redone
End Sub

Private Sub RenumberGrid()
'Redo the row numbering on the grid
  Dim saverow As Integer
  Dim savecol As Integer
  Dim i As Integer
  
  Dim g As Control
  Set g = grdNozzle
  saverow = g.Row
  savecol = g.Col

  g.Col = 0
  For i = 1 To g.Rows - 1
    g.Row = i
    g.Text = Format$(i)
  Next

  g.Row = saverow
  g.Col = savecol
End Sub

Private Sub cboView_Click()
  If PropTakeAction Then
    DataToForm  'Refresh the view
  End If
End Sub

Private Sub cmdAddTable_Click()
  AddNozzle
End Sub

Private Sub cmdAddPicture_Click()
  AddNozzle
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub ClearSelectedCells()
'clear the selected cells in a grid
  Dim saverow As Integer
  Dim savecol As Integer
  Dim ir As Integer, ir1 As Integer, ir2 As Integer
  Dim ic As Integer, ic1 As Integer, ic2 As Integer
  
  With grdNozzle
    'Ensure .Row is before .RowSel
    If .RowSel >= .Row Then
      saverow = .Row
    Else
      saverow = .RowSel
      .RowSel = .Row
      .Row = saverow
    End If
    'Ensure .Col is before .ColSel
    If .ColSel >= .Col Then
      savecol = .Col
    Else
      savecol = .ColSel
      .ColSel = .Col
      .Col = savecol
    End If
    'Clear the cell contents in the selected area
    ir1 = .Row: ir2 = .RowSel
    ic1 = .Col: ic2 = .ColSel
    For ic = ic1 To ic2
      .Col = ic
      For ir = ir1 To ir2
        .Row = ir
        .Text = ""
      Next
    Next
    .Row = saverow
    .Col = savecol
  End With
  'set dist type to "user-defined"
  If NZ2.Type <> 1 Then ChangeType 1 'user-def
End Sub

Private Sub cmdDeleteTable_Click()
  DeleteNozzlesTable
End Sub

Private Sub cmdDeletePicture_Click()
  DeleteNozzlesPicture
End Sub

Private Sub cmdEditDrop_Click(Index As Integer)
  Me.MousePointer = vbHourglass
  Load frmDropDist
  frmDropDist.lblDSDselection = Index 'Send the DSD index to the form
  frmDropDist.Show vbModal
  UpdateDSDTypeLabel Index 'Update it, it might have changed
  Me.MousePointer = vbDefault
End Sub

Private Sub cmdGenRegDist_Click()
cmdGenRegDist_Click_Top:
  frmNozzleRD.Show vbModal
  If Not frmNozzleRD.Cancelled Then
    If GenRegNozDist(NZ2, UD.AC.SemiSpan, _
         frmNozzleRD.txtNozzles.Text, _
         frmNozzleRD.txtExtent.Text, _
         frmNozzleRD.txtSpacing.Text) Then
      DataToForm
    Else
      GoTo cmdGenRegDist_Click_Top
    End If
  End If
  Unload frmNozzleRD
End Sub

Private Sub cmdImport_Click()
  ImportNozzles
End Sub

Private Sub cmdOk_Click()
  FormToData
  Unload Me
End Sub

Private Sub cmdSort_Click()
  SortNozzles
End Sub

Private Sub cmdZoomFit_Click()
  ZoomView 0
End Sub

Private Sub cmdZoomIn_Click()
  ZoomView 1.25
End Sub

Private Sub cmdZoomOut_Click()
  ZoomView 0.75
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim c As Control
  If KeyAscii = 27 Then
    For Each c In ControlsThatMayReceiveEscape
      If c Is Me.ActiveControl Then
        Exit Sub
      End If
    Next
    cmdCancel_Click
  End If
End Sub

Private Sub Form_Load()
'Initialize the controls on this form
  Dim g As Control
  Dim i As Integer
  
  'center the form
  CenterForm Me

  'Initialize the collection of controls that may receive
  'an escape character. This allows Escape to dismiss the
  'form OR abort an EditGrid edit.
  With ControlsThatMayReceiveEscape
    .Add txtEdit
  End With
  
  'Copy user data to local storage for editing
  NZ2 = UD.NZ
  
  'Prevent form controls from responding to events
  OptTakeAction = False
  PropTakeAction = False
  
  'Graphical Editing controls
  NumNozzleControls = 1 'the form starts with one control in the array
  NumCurrentNozzles = 0 'no selection yet
  
  'View control
  cboView.Clear
  cboView.AddItem "Rear"
  cboView.AddItem "Top"
  cboView.ListIndex = 0
  
  'Properties
  lblACName.Caption = UD.AC.Name
  Select Case UD.AC.WingType
  Case 3: 'fixed
    lblSemiSpanLabel.Caption = "Semispan:"
  Case 4: 'helicopter
    lblSemiSpanLabel.Caption = "Rotor Radius:"
  End Select
  lblSemiSpan.Caption = AGFormat$(UnitsDisplay(UD.AC.SemiSpan, UN_LENGTH))
  lblSemiSpanUnits.Caption = UnitsName(UN_LENGTH)
  
  'Nozzle Parameters
  lblNozParamHUnits.Caption = UnitsName(UN_LENGTH)
  lblNozParamVUnits.Caption = UnitsName(UN_LENGTH)
  lblNozParamFUnits.Caption = UnitsName(UN_LENGTH)

  'Nozzle Position Grid
  Set g = grdNozzle
  g.Row = 0
  g.Col = 0
  g.Text = "Nozzle"
  g.Col = 1
  g.Text = "DSD"
  g.Col = 2
  g.Text = "Horizontal (" + UnitsName(UN_LENGTH) + ")"
  g.Col = 3
  g.Text = "Vertical (" + UnitsName(UN_LENGTH) + ")"
  g.Col = 4
  g.Text = "Forward (" + UnitsName(UN_LENGTH) + ")"
  g.FixedAlignment(0) = flexAlignCenterCenter
  g.FixedAlignment(1) = flexAlignCenterCenter
  g.FixedAlignment(2) = flexAlignCenterCenter
  g.FixedAlignment(3) = flexAlignCenterCenter
  g.FixedAlignment(4) = flexAlignCenterCenter
  g.ColAlignment(1) = flexAlignCenterCenter
  g.ColAlignment(2) = flexAlignCenterCenter
  g.ColAlignment(3) = flexAlignCenterCenter
  g.ColAlignment(4) = flexAlignCenterCenter
  
  'allow option button changes to take action
  '(see declarations section)
  OptTakeAction = True
  PropTakeAction = True

  'select the default display
  'this triggers the click event, which does a DataToForm
  optDisplay(0).Value = True
  
  'initialize the graphical view
  ZoomView 0  'fit
End Sub

Private Sub Form_Resize()
  'Adjust control positions and sizes to fit the form
  'These controls don't change size or location:
    'Distribution Type frame
    'Display Frame
    'Regular Diatribution frame
  
  Dim minwidth As Single
  Dim newwidth As Single
  Dim mintop As Single
  Dim newtop As Single
  
  Const MRGN = 60
  
  'Distribution Type
  'does not move
  
  'Distribution Properties frame
  'Right edge attached to form
  minwidth = fraNozLoc.Width + fraNozDSD.Width + _
             cmdOK.Width + cmdCancel.Width + _
             4 * MRGN
  newwidth = Me.ScaleWidth - MRGN
  If newwidth > minwidth Then
    fraProp.Width = newwidth
  Else
    fraProp.Width = minwidth
  End If
  
  'Nozzle Location frame, Nozzle DSD frame
  'fixed to bottom of form
  mintop = fraProp.Height + MRGN + 2000
  newtop = Me.ScaleHeight - fraNozLoc.Height - MRGN
  If newtop > mintop Then
    fraNozLoc.Top = newtop
    fraNozDSD.Top = newtop
  Else
    fraNozLoc.Top = mintop
    fraNozDSD.Top = mintop
  End If

  'View frame
  'fixed to Distribution Properties and Nozzle Location
  fraView.Height = fraNozLoc.Top - MRGN - fraProp.Height
  fraView.Width = fraProp.Left + fraProp.Width
  
  'Ok/Cancel buttons
  'fixed to View frame and Nozzle DSD
  cmdCancel.Top = fraNozDSD.Top + fraNozDSD.Height - cmdCancel.Height
  cmdCancel.Left = fraView.Left + fraView.Width - cmdCancel.Width
  cmdOK.Top = cmdCancel.Top
  cmdOK.Left = cmdCancel.Left - cmdOK.Width - MRGN
  
  If optDisplay(0) Then
    ResizeGraphicalView
  ElseIf optDisplay(1) Then
    ResizeTabularView
  End If
End Sub

Private Sub grdNozzle_DblClick()
  EditGridCell 0
End Sub

Private Sub grdNozzle_KeyDown(KeyCode As Integer, Shift As Integer)
  'PgUp and PgDn mess up the grid control
  If KeyCode = 33 Or KeyCode = 34 Then
    KeyCode = 0
  End If
End Sub

Private Sub grdNozzle_KeyPress(KeyAscii As Integer)
  EditGridCell KeyAscii
End Sub

Private Sub EditGridCell(KeyAscii As Integer)
'Start editing the current grid cell
  ' Move the text box to the current grid cell:
  PositionTextBox

  ' Save the position of the grids Row and Col for later:
  gRow = grdNozzle.Row
  gCol = grdNozzle.Col

  ' Make text box same size as current grid cell:
  txtEdit.Width = grdNozzle.ColWidth(grdNozzle.Col) - 2 * Screen.TwipsPerPixelX
  txtEdit.Height = grdNozzle.RowHeight(grdNozzle.Row) - 2 * Screen.TwipsPerPixelY

  ' Transfer the grid cell text:
  txtEdit.Text = grdNozzle.Text
  txtEdit.SelStart = Len(grdNozzle.Text)
  
  ' Show the text box:
  txtEdit.Visible = True
  txtEdit.ZOrder 0
  txtEdit.SetFocus

  ' Init flags
  CancelCellEdit = False
  CellEdited = False
  
  ' Redirect this KeyPress event to the text box:
  If KeyAscii <> 13 Then 'Enter
     SendKeys Chr$(KeyAscii)
  End If
End Sub

Private Sub PositionTextBox()
'Move the txtEdit Text to cover the current grid cell
  With grdNozzle
    If .RowIsVisible(.Row) And .ColIsVisible(.Col) Then
    txtEdit.Left = .Left + .CellLeft
    txtEdit.Top = .Top + .CellTop
    txtEdit.Height = .CellHeight
    txtEdit.Width = .CellWidth
    Else
      .SetFocus
    End If
  End With
End Sub

Private Sub InsertCellRow()
'Insert a blank row of cells in a grid above the current row
  Dim g As Control
  
  Set g = grdNozzle
  
  If g.Rows - 1 < MAX_NOZZLES Then
    g.AddItem "0", g.Row    'add a new row
    RenumberGrid
    If NZ2.Type <> 1 Then ChangeType 1 'user-def
  Else
    Beep
  End If
End Sub

Private Sub DeleteNozzlesTable()
'Delete the selected nozzles(s) in the grid control
  Dim saverow As Integer
  Dim savecol As Integer
  Dim ist As Integer
  Dim ien As Integer
  Dim nd As Integer
  Dim ne As Integer
  Dim i As Integer

  Dim g As Control
  
  Set g = grdNozzle
  
  With g
    'Ensure .Row is before .RowSel
    If .RowSel >= .Row Then
      saverow = .Row
    Else
      saverow = .RowSel
      .RowSel = .Row
      .Row = saverow
    End If
  End With
  
  ist = g.Row - 1 'starting element
  ien = g.RowSel - 1   'ending element
  nd = ien - ist + 1      'number of deleted elements
  ne = NZ2.NumNoz         'original number of elements
  
  'make sure the header is not selected, it's not the last
  'element, and that all the elements are not to be deleted
  If ist < 0 Or ne = 1 Or ne = nd Then
    Beep
    Exit Sub
  End If
  
  'remove array elements by shifting remaining items down
  For i = 0 To ne - ien - 2
    NZ2.NozType(ist + i) = NZ2.NozType(ien + i + 1)
    NZ2.PosHoriz(ist + i) = NZ2.PosHoriz(ien + i + 1)
    NZ2.PosVert(ist + i) = NZ2.PosVert(ien + i + 1)
    NZ2.PosFwd(ist + i) = NZ2.PosFwd(ien + i + 1)
  Next
  NZ2.NumNoz = NZ2.NumNoz - nd
  lblNumNoz.Caption = Format$(NZ2.NumNoz)
  
  'remove grid cells
  For i = ist To ien
    g.RemoveItem ist + 1
  Next
  If g.Row = 0 Then g.Row = g.Rows - 1
  g.RowSel = g.Row
  RenumberGrid                'renumber the grid
  
  UpdateDistributionExtent
  If NZ2.Type <> 1 Then ChangeType 1 'user-def
End Sub

Private Sub DeleteNozzlesPicture()
'Delete the selected nozzles(s) in the picture
  Dim NZdel() As Boolean
  Dim idel As Integer
  Dim inoz As Integer
  Dim nnoz As Integer

  If NumCurrentNozzles = 0 Then Exit Sub
  
  'nozzle selection status
  ReDim NZdel(NZ2.NumNoz - 1)
  For inoz = 0 To NZ2.NumNoz - 1
    NZdel(inoz) = False
  Next
  For idel = 0 To NumCurrentNozzles - 1
    NZdel(CurrentNozzles(idel)) = True
  Next
  'all current nozzles will be deleted, so reset current arrays
  NumCurrentNozzles = 0
  
  'delete nozzles by shifting down the unselected ones
  idel = -1
  nnoz = 0
  For inoz = 0 To NZ2.NumNoz - 1
    If NZdel(inoz) Then
      If idel < 0 Then idel = inoz
    Else
      If idel >= 0 Then
        NZ2.NozType(idel) = NZ2.NozType(inoz)
        NZ2.PosHoriz(idel) = NZ2.PosHoriz(inoz)
        NZ2.PosVert(idel) = NZ2.PosVert(inoz)
        NZ2.PosFwd(idel) = NZ2.PosFwd(inoz)
        idel = idel + 1
      End If
      nnoz = nnoz + 1 'count remaining nozzles
    End If
  Next
  'delete unused nozzle controls
'tbc   'hide the last control
  For inoz = nnoz To NZ2.NumNoz - 1
    Unload picNozzle(inoz)
  Next
  NumNozzleControls = nnoz
  'Update number of nozzles
  NZ2.NumNoz = nnoz
  
  UpdateNozzlePictureAll
  UpdateNozzleParamControls
  UpdateDistributionExtent
  If NZ2.Type <> 1 Then ChangeType 1 'user-def
End Sub

Private Sub hscViewPort_Change()
  If PropTakeAction Then
    picViewArea.Left = hscViewPort.Value * _
      (picViewPort.ScaleWidth - picViewArea.Width) / _
      (hscViewPort.Max - hscViewPort.Min)
  End If
End Sub

Private Sub optDisplay_Click(Index As Integer)
  Select Case Index
  Case 0: 'Graphical
    picTabular.Visible = False
    DataToForm
    ResizeGraphicalView
    ZoomView 1
    picGraphical.Visible = True
  Case 1: 'Table
    picGraphical.Visible = False
    DataToForm
    ResizeTabularView
    picTabular.Visible = True
  End Select
End Sub

Private Sub optDSD_Click(Index As Integer)
'Set the DSD of the currently selected nozzles
  Dim i As Integer
    
  If NumCurrentNozzles = 0 Then
    optDSD(Index).Value = False
  Else
    For i = 0 To NumCurrentNozzles - 1
      NZ2.NozType(CurrentNozzles(i)) = Index
    Next
    UpdateNozzlePictureCurrent
  End If
End Sub

Private Sub picNozzle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim nzselect As Integer
  
  If (Button And 1) Then 'left button
    'init drag coords
    DragOffsetX = X
    DragOffsetY = Y
    'select nozzle
    nzselect = AlreadySelected(Index)
    If Shift = 0 And Not nzselect Then
      NumCurrentNozzles = 0 'clear current selections
    End If
    'add to selection list if not already selected
    If Not nzselect Then
      NumCurrentNozzles = NumCurrentNozzles + 1
      CurrentNozzles(NumCurrentNozzles - 1) = Index
    End If
    HighlightNozzles
  End If
End Sub

Private Sub picNozzle_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  Dim ic As Integer
  Dim dx As Single
  Dim dy As Single
  If (Button And 1) > 0 Then 'Left button
    dx = -DragOffsetX + X
    dy = -DragOffsetY + Y
    If Not Dragging Then
      If (dx * dx + dy * dy) >= DRAGSTARTDIST2 Then
        Dragging = True
      End If
    End If
    If Dragging Then
      If NZ2.Type <> 1 Then ChangeType 1 'user-def
      'reposition the controls
      For i = 0 To NumCurrentNozzles - 1
        ic = CurrentNozzles(i)
        picNozzle(ic).Left = picNozzle(ic).Left + dx
        picNozzle(ic).Top = picNozzle(ic).Top + dy
      Next
    End If
  End If
End Sub

Private Sub picNozzle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If (Button And 1) > 0 Then 'Left button
    If Dragging Then
      Dragging = False 'turn off the dragging flag
      UpdateNozzleCoords
      UpdateNozzlePictureCurrent
      UpdateDistributionExtent
    End If
    UpdateNozzleParamControls
  End If
End Sub

Private Sub picViewArea_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If (Button And 1) > 0 Then 'Left button
    'Begin selection process
    SelStartX = X
    SelStartY = Y
    shpSelect.Left = SelStartX
    shpSelect.Top = SelStartY
    shpSelect.Width = 0
    shpSelect.Height = 0
    shpSelect.Visible = True
  End If
End Sub

Private Sub picViewArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If (Button And 1) > 0 Then 'Left button
    If X >= SelStartX Then
      shpSelect.Width = X - SelStartX
    Else
      shpSelect.Left = X
      shpSelect.Width = SelStartX - X
    End If
    If Y >= SelStartY Then
      shpSelect.Height = Y - SelStartY
    Else
      shpSelect.Top = Y
      shpSelect.Height = SelStartY - Y
    End If
  End If
End Sub

Private Sub picViewArea_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim ic As Integer
  
  If (Button And 1) > 0 Then 'Left button
    'select all nozzles within shpSelect
    'If Shift is down, add the new selected nozzles to the list, if not
    'replace the selected list with the new selections
    If Shift = 0 Then NumCurrentNozzles = 0
    
    For ic = 0 To NumNozzleControls - 1
      'see if the nozzle is within the box
      If picNozzle(ic).Visible Then
        If (picNozzle(ic).Left + picNozzle(ic).Width > shpSelect.Left) And _
           (picNozzle(ic).Left < shpSelect.Left + shpSelect.Width) And _
           (picNozzle(ic).Top + picNozzle(ic).Height > shpSelect.Top) And _
           (picNozzle(ic).Top < shpSelect.Top + shpSelect.Height) Then
          'Add the nozzle to the list
          If Not AlreadySelected(ic) Then
            NumCurrentNozzles = NumCurrentNozzles + 1
            CurrentNozzles(NumCurrentNozzles - 1) = ic
          End If
        End If
      End If
    Next
    shpSelect.Visible = False
    UpdateNozzlePictureAll
    UpdateNozzleParamControls
  End If
End Sub

Private Sub txtEdit_Change()
  CellEdited = True 'remember that changes have been made
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'Change the function of some special keys
  Select Case KeyCode
    Case vbKeyDown
      grdNozzle.SetFocus ' Set focus back to grid, see txtEdit_LostFocus.
      SendKeys "{DOWN}" 'send a downarrow to the grid
    Case vbKeyUp
      grdNozzle.SetFocus ' Set focus back to grid, see txtEdit_LostFocus.
      SendKeys "{UP}"   'send an uparrow to the grid
  End Select
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
  Case 13  'Enter
    grdNozzle.SetFocus ' Set focus back to grid, see txtEdit_LostFocus.
    KeyAscii = 0      ' Remove this KeyPress.
  Case 27  'Escape - abort edit
    CancelCellEdit = True 'prevent LostFocus from making edit permanent
    grdNozzle.SetFocus ' Set focus back to grid, see txtEdit_LostFocus.
    KeyAscii = 0      ' Remove this KeyPress.
  End Select
End Sub

Private Sub txtEdit_LostFocus()
  Dim tmpRow As Integer
  Dim tmpCol As Integer
  Dim tmpType As Integer

  ' Save current settings of Grid Row and col. This is needed only if
  ' the focus is set somewhere else in the Grid.
  tmpRow = grdNozzle.Row
  tmpCol = grdNozzle.Col

  ' Set Row and Col back to what they were before txtEdit_LostFocus:
  grdNozzle.Row = gRow
  grdNozzle.Col = gCol

  If CellEdited And Not CancelCellEdit Then
    If NZ2.Type <> 1 Then ChangeType 1 'user-def
    grdNozzle.Text = txtEdit.Text  ' Transfer text back to grid.
    Select Case gCol
    Case 1 'NozType
      tmpType = Val(txtEdit.Text) - 1 'Display types are 1-3
      If tmpType < 0 Then tmpType = 0
      If tmpType > 2 Then tmpType = 2
      NZ2.NozType(gRow - 1) = tmpType
      grdNozzle.Text = Format$(tmpType + 1) 'Replace grid contents in case of mods
    Case 2 'Horiz
      NZ2.PosHoriz(gRow - 1) = UnitsInternal(Val(txtEdit.Text), UN_LENGTH)
      UpdateDistributionExtent
    Case 3 'Vert
      NZ2.PosVert(gRow - 1) = UnitsInternal(Val(txtEdit.Text), UN_LENGTH)
    Case 4 'Fwd
      NZ2.PosFwd(gRow - 1) = UnitsInternal(Val(txtEdit.Text), UN_LENGTH)
    End Select
    If NZ2.Type <> 1 Then ChangeType 1 'user-def
    CellEdited = False 'reset flag in case this method is repeated
  End If
  txtEdit.SelStart = 0       ' Return caret to beginning.
  txtEdit.Visible = False    ' Disable text box.

  ' Return row and Col contents:
  grdNozzle.Row = tmpRow
  grdNozzle.Col = tmpCol
End Sub

Private Sub ArrayToGrid(NumNoz As Integer, NozType() As Integer, _
                        PosHoriz() As Single, _
                        PosVert() As Single, PosFwd() As Single)
'Place the values contained in the given array into
'the grid control
  Dim i As Integer
  Dim g As Control
  
  Set g = grdNozzle

  'set the number of rows
  g.Rows = NumNoz + 1
  
  'copy the data
  For i = 1 To g.Rows - 1
    g.Row = i
    g.Col = 1
    g.Text = Format$(NozType(i - 1) + 1) 'Displayed types are 1-3
    g.Col = 2
    g.Text = AGFormat$(UnitsDisplay(PosHoriz(i - 1), UN_LENGTH))
    g.Col = 3
    g.Text = AGFormat$(UnitsDisplay(PosVert(i - 1), UN_LENGTH))
    g.Col = 4
    g.Text = AGFormat$(UnitsDisplay(PosFwd(i - 1), UN_LENGTH))
  Next
  
  'number the rows
  RenumberGrid
End Sub

Private Sub AddNozzle()
'Add a new nozzle and adjust all form controls
  Dim newpos As Integer
  
  Dim g As Control
  
  Set g = grdNozzle
  
  'first, see if there is room for a new nozzle
  If NZ2.NumNoz = MAX_NOZZLES Then
    Beep
    Exit Sub
  End If
  
  'set the distribution type to user-defined
  If NZ2.Type <> 1 Then ChangeType 1 'user-def
  
  'add a new element into the local array
  NZ2.NumNoz = NZ2.NumNoz + 1
  newpos = NZ2.NumNoz - 1 'the new position is at the end
  NZ2.PosHoriz(newpos) = 0 'initial value
  NZ2.PosVert(newpos) = 0 'initial value
  NZ2.PosFwd(newpos) = 0 'initial value
  NZ2.NozType(newpos) = 0 'initial value
  
  If optDisplay(0) Then 'graphical view
    'add a new nozzle control
    If newpos > NumNozzleControls - 1 Then
      NumNozzleControls = NumNozzleControls + 1
      Load picNozzle(newpos)
    End If
    NumCurrentNozzles = 1
    CurrentNozzles(NumCurrentNozzles - 1) = newpos
    picNozzle(newpos).Visible = True
    UpdateNozzlePictureAll
    UpdateNozzleParamControls
    UpdateDistributionExtent
  ElseIf optDisplay(1) Then 'tabular view
    'add a new row to the grid
    g.AddItem Format$(newpos + 1) + Chr$(9) + "1", newpos + 1
    g.Row = newpos + 1
    g.RowSel = newpos + 1
    lblNumNoz.Caption = Format$(NZ2.NumNoz)
  End If

End Sub

Private Sub SortNozzles()
'Sort the existing Nozzle array and update form controls
  Dim tmpHoriz() As Single
  Dim tmpNum As Integer
  Dim i As Integer
  
  If NZ2.NumNoz <= 1 Then Exit Sub
  
  'sort the array into a temporary array
  ReDim tmpHoriz(NZ2.NumNoz - 1)
  tmpNum = 0
  For i = 0 To NZ2.NumNoz - 1
    AddToArray NZ2.PosHoriz(i), tmpNum, tmpHoriz()
  Next
  'load the sorted array back into the original array
  For i = 0 To NZ2.NumNoz - 1
    NZ2.PosHoriz(i) = tmpHoriz(i)
  Next
  
  If optDisplay(0) Then 'graphical
    'update the nozzle controls
    UpdateNozzlePictureAll
  
  ElseIf optDisplay(1) Then
    'update the grid control
    ArrayToGrid NZ2.NumNoz, NZ2.NozType(), NZ2.PosHoriz(), NZ2.PosVert(), NZ2.PosFwd()
  End If
  
End Sub

Private Sub UpdateDistributionExtent()
'calculate and display Nozzle Distribution Extent
  Dim farleft As Single
  Dim farright As Single
  Dim PosHoriz As Single
  Dim i As Integer
  
  'find the farthest-out nozzle
  farleft = 0
  farright = 0
  For i = 0 To NZ2.NumNoz - 1
    PosHoriz = NZ2.PosHoriz(i)
    If PosHoriz > farright Then farright = PosHoriz
    If PosHoriz < farleft Then farleft = PosHoriz
  Next
  lblLimit(0).Caption = AGFormat$(-farleft / UD.AC.SemiSpan * 100)
  lblLimit(1).Caption = AGFormat$(farright / UD.AC.SemiSpan * 100)
End Sub

Private Sub ImportNozzles()
'Read a set of nozle horizontal positions from a text file
  Dim fn As String
  
  If FileDialog(FD_OPEN, FD_TYPE_TEXT, fn) Then  'get a filename
    'Open the file
    On Error GoTo ErrHandImportNozzles
    OpenFileAndSkipComments fn, 1
    
    'read nozzle data
    With NZ2
      .NumNoz = 0
      While Not EOF(1)
        .NozType(.NumNoz) = 0
        Input #1, .PosHoriz(.NumNoz), .PosVert(.NumNoz), .PosFwd(.NumNoz)
        .NumNoz = .NumNoz + 1
      Wend
    End With
    Close #1
    If NZ2.Type <> 1 Then ChangeType 1 'user-def

    DataToForm
  End If
  Exit Sub

ErrHandImportNozzles:
  Close #1
  MsgBox "Error importing file: " + fn + vbCrLf + Error$(Err), _
         vbCritical + vbOKOnly
  Exit Sub
End Sub

Private Sub UpdateDSDTypeLabel(Index)
'update the state of the DropDist Type and
'description labels

  Select Case UD.DSD(Index).Type
  Case 0  'basic
    lblDSDdesc(Index).Caption = "Basic"
  Case 1  'dropkick
    lblDSDdesc(Index).Caption = "DropKick"
  Case 2  'user-defined
    lblDSDdesc(Index).Caption = "User-defined"
  Case 3  'library (SDTF)
    lblDSDdesc(Index).Caption = "Library (SDTF)"
  Case 4  'library (FS)
    lblDSDdesc(Index).Caption = "Library (FS)"
  Case 5  'dropkirk
    lblDSDdesc(Index).Caption = "USDA ARS Nozzle Models"
  End Select
  lblDSDdesc(Index).Caption = lblDSDdesc(Index).Caption + " (" + Trim$(UD.DSD(Index).Name) + ")"
End Sub

Private Sub ResizeGraphicalView()
'Resize the Graphical View Controls, but not the graphical view itself
'This routine only handles sizing the viewport, scollbars, etc. The view
'area is taken care of by ZoomView
  
  Const MRGN = 100
  
  'Fit the container picturebox to the view frame
  picGraphical.Top = optDisplay(0).Top + optDisplay(0).Height + MRGN
  picGraphical.Left = MRGN
  picGraphical.Width = fraView.Width - MRGN - MRGN
  picGraphical.Height = fraView.Height - optDisplay(0).Top - optDisplay(0).Height - MRGN - MRGN
  
  'Postition the view controls along the bottom edge
  cboView.Left = 0
  cboView.Top = picGraphical.ScaleHeight - cboView.Height
  
  cmdZoomIn.Left = cboView.Left + cboView.Width
  cmdZoomIn.Top = picGraphical.ScaleHeight - cmdZoomIn.Height
  
  cmdZoomOut.Left = cmdZoomIn.Left + cmdZoomIn.Width
  cmdZoomOut.Top = picGraphical.ScaleHeight - cmdZoomOut.Height
  
  cmdZoomFit.Left = cmdZoomOut.Left + cmdZoomOut.Width
  cmdZoomFit.Top = picGraphical.ScaleHeight - cmdZoomFit.Height
  
  cmdAddPicture.Left = cmdZoomFit.Left + cmdZoomFit.Width + MRGN
  cmdAddPicture.Top = picGraphical.ScaleHeight - cmdAddPicture.Height
  
  cmdDeletePicture.Left = cmdAddPicture.Left + cmdAddPicture.Width
  cmdDeletePicture.Top = picGraphical.ScaleHeight - cmdDeletePicture.Height
  
  hscViewPort.Top = picGraphical.ScaleHeight - hscViewPort.Height
  hscViewPort.Left = cmdDeletePicture.Left + cmdDeletePicture.Width + MRGN
  hscViewPort.Width = picGraphical.ScaleWidth - hscViewPort.Left - vscViewPort.Width
  
  'position the vertical scroll bar along the right edge
  vscViewPort.Top = 0
  vscViewPort.Left = picGraphical.ScaleWidth - vscViewPort.Width
  vscViewPort.Height = picGraphical.ScaleHeight - hscViewPort.Height
  
  'position the viewport picturebox
  picViewPort.Top = 0
  picViewPort.Left = 0
  picViewPort.Width = picGraphical.ScaleWidth - vscViewPort.Width
  picViewPort.Height = picGraphical.ScaleHeight - hscViewPort.Height

End Sub

Private Sub ResizeTabularView()
'Resize the Graphical View Controls
  Dim wid As Single
  Dim i As Integer
  Const MRGN = 100
  
  With picTabular
    .Top = optDisplay(0).Top + optDisplay(0).Height + MRGN
    .Left = MRGN
    .Width = fraView.Width - MRGN - MRGN
    .Height = fraView.Height - optDisplay(0).Top - optDisplay(0).Height - MRGN - MRGN
  End With
  
  With lblNumNozLabel
    .Top = picTabular.Height - .Height - MRGN
    .Left = cmdDeleteTable.Left + cmdDeleteTable.Width + MRGN
  End With
  
  With lblNumNoz
    .Top = picTabular.Height - .Height - MRGN
    .Left = lblNumNozLabel.Left + lblNumNozLabel.Width + MRGN
  End With
  
  With grdNozzle
    .Width = 5000
    .Height = lblNumNozLabel.Top - .Top - MRGN
    wid = CSng(.Width - .ColWidth(0) - .ColWidth(1) - 325) / 3!
    For i = 2 To .cols - 1
      .ColWidth(i) = wid
    Next
  End With
End Sub

Private Sub ZoomView(zoomfactor)
'Zoom the graphical view up, down, or to fit
'
'zoomfactor  i   the factor by which to scale the scene.
'                0 = fit scene to the current viewport size
'                1 = refresh scene without resizing
'
'
'picViewArea holds the entire viewed scene. We resize this
'control and all that it holds to zoom the view.
  Dim pvph As Single
  Dim pvpw As Single
  Dim pvah As Single
  Dim pvaw As Single
  Dim pval As Single
  Dim pvat As Single
  Dim newwidth As Single
  Dim newheight As Single
  Dim newleft As Single
  Dim newtop As Single
  Dim newsize As Single
  Dim maxsize As Single
  Dim hmn As Single
  Dim hmx As Single
  Dim vmn As Single
  Dim vmx As Single
  Dim fmn As Single
  Dim fmx As Single
  Dim i As Integer

  'Hiding the viewport makes the redraw operation faster
  picViewPort.Visible = False
  
  'gather various control information
  pvph = picViewPort.ScaleHeight 'Inside height
  pvpw = picViewPort.ScaleWidth  'Inside width
  
  pvah = picViewArea.Height
  pvaw = picViewArea.Width
  pval = picViewArea.Left
  pvat = picViewArea.Top
  
  'Find the max dimension of the viewport
  If pvph < pvpw Then
    maxsize = pvpw
  Else
    maxsize = pvph
  End If
  
  'calculate the new view area size and position
  'and limit the zoomfactor if required
  If zoomfactor = 0 Then  'fit
    'find the extents of the wing and nozzles
    hmn = -UD.AC.SemiSpan
    hmx = UD.AC.SemiSpan
    vmn = -UD.AC.BoomVert
    vmx = -UD.AC.BoomVert
    If UD.AC.WingType = 3 Then 'fixed
      fmn = 0
      fmx = 0.1 * UD.AC.SemiSpan 'the wing is drawn this way
    ElseIf UD.AC.WingType = 4 Then 'heli
      fmn = -UD.AC.SemiSpan
      fmx = UD.AC.SemiSpan
    End If
    For i = 0 To NZ2.NumNoz - 1
      If NZ2.PosHoriz(i) < hmn Then hmn = NZ2.PosHoriz(i)
      If NZ2.PosHoriz(i) > hmx Then hmx = NZ2.PosHoriz(i)
      If NZ2.PosVert(i) > vmx Then vmx = NZ2.PosVert(i)
      If NZ2.PosVert(i) < vmn Then vmn = NZ2.PosVert(i)
      If NZ2.PosFwd(i) > fmx Then fmx = NZ2.PosFwd(i)
      If NZ2.PosFwd(i) < fmn Then fmn = NZ2.PosFwd(i)
    Next
    newwidth = ((2 * UD.AC.SemiSpan) / (hmx - hmn)) * (1.9 * pvpw) '2.0 is a perfect fit
    If cboView.ListIndex = 0 Then 'rear view
      If (vmx - vmn) <> 0 Then
        newheight = ((2 * UD.AC.SemiSpan) / (vmx - vmn)) * (1.9 * pvph) '2.0 is a perfect fit
      Else
        newheight = newwidth + 1 'make sure newwidth wins below
      End If
    ElseIf cboView.ListIndex = 1 Then 'top view
      If (fmx - fmn) <> 0 Then
        newheight = ((2 * UD.AC.SemiSpan) / (fmx - fmn)) * (1.9 * pvph) '2.0 is a perfect fit
      Else
        newheight = newwidth + 1 'make sure newwidth wins below
      End If
    End If
    If newwidth < newheight Then
      newsize = newwidth
    Else
      newsize = newheight
    End If
    If newsize < maxsize Then newsize = maxsize
    'position the view area to center the extents
    newleft = (0.5 * (pvpw - newsize)) - (newsize * (hmn + hmx) / (8 * UD.AC.SemiSpan))
    If cboView.ListIndex = 0 Then
      newtop = (0.5 * (pvph - newsize)) + (newsize * (vmn + vmx) / (8 * UD.AC.SemiSpan))
    ElseIf cboView.ListIndex = 1 Then
      newtop = (0.5 * (pvph - newsize)) + (newsize * (fmn + fmx) / (8 * UD.AC.SemiSpan))
    End If
  Else
    newsize = pvaw * zoomfactor
    If newsize < maxsize Then
      newsize = maxsize
      zoomfactor = newsize / pvaw
    End If
    'position the view area to have the same relative center as before
    newleft = (0.5 * pvpw) + (pval - (0.5 * pvpw)) * zoomfactor
    newtop = (0.5 * pvph) + (pvat - (0.5 * pvph)) * zoomfactor
  End If
  
  'Keep the view area within the viewport
  If newleft > 0 Then newleft = 0
  If newleft < pvpw - newsize Then newleft = pvpw - newsize
  If newtop > 0 Then newtop = 0
  If newtop < pvph - newsize Then newtop = pvph - newsize
  
  'resize/reposition the view area
  picViewArea.Height = newsize
  picViewArea.Width = newsize
  picViewArea.Left = newleft
  picViewArea.Top = newtop
  
  'Adjust the scroll bars
  If newsize - pvpw > 0 Then
    hscViewPort.Enabled = True
    PropTakeAction = False 'prevent value change from taking action
    hscViewPort.Value = newleft * (hscViewPort.Max - hscViewPort.Min) / (pvpw - newsize)
    PropTakeAction = True
  Else
    hscViewPort.Enabled = False
  End If
 
  If newsize - pvph > 0 Then
    vscViewPort.Enabled = True
    PropTakeAction = False 'prevent value change from taking action
    vscViewPort.Value = newtop * (vscViewPort.Max - vscViewPort.Min) / (pvph - newsize)
    PropTakeAction = True
  Else
    vscViewPort.Enabled = False
  End If
  
  'Redraw the elements in the view area
  DrawNozzlePicture
  
  picViewPort.Visible = True
End Sub

Private Sub txtNozParamH_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i As Integer
  Dim pos As Single
  If KeyCode = 13 Then
    If NumCurrentNozzles > 0 Then
      If NZ2.Type <> 1 Then ChangeType 1 'user-def
      pos = UnitsInternal(Val(txtNozParamH.Text), UN_LENGTH)
      For i = 0 To NumCurrentNozzles - 1
        NZ2.PosHoriz(CurrentNozzles(i)) = pos
      Next
      UpdateNozzlePictureCurrent
    End If
    KeyCode = 0
  End If
End Sub

Private Sub txtNozParamV_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i As Integer
  Dim pos As Single
  If KeyCode = 13 Then
    If NumCurrentNozzles > 0 Then
      If NZ2.Type <> 1 Then ChangeType 1 'user-def
      pos = UnitsInternal(Val(txtNozParamV.Text), UN_LENGTH)
      For i = 0 To NumCurrentNozzles - 1
        NZ2.PosVert(CurrentNozzles(i)) = pos
      Next
      UpdateNozzlePictureCurrent
    End If
    KeyCode = 0
  End If
End Sub

Private Sub txtNozParamF_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim i As Integer
  Dim pos As Single
  If KeyCode = 13 Then
    If NumCurrentNozzles > 0 Then
      If NZ2.Type <> 1 Then ChangeType 1 'user-def
      pos = UnitsInternal(Val(txtNozParamF.Text), UN_LENGTH)
      For i = 0 To NumCurrentNozzles - 1
        NZ2.PosFwd(CurrentNozzles(i)) = pos
      Next
      UpdateNozzlePictureCurrent
    End If
    KeyCode = 0
  End If
End Sub

Private Sub vscViewPort_Change()
  If PropTakeAction Then
    picViewArea.Top = vscViewPort.Value * _
      (picViewPort.ScaleHeight - picViewArea.Height) / _
      (vscViewPort.Max - vscViewPort.Min)
  End If
End Sub

Private Sub DrawNozzlePicture()
'Draw the nozzle picture in the View Area
'
'

  'scale the View Area so that we can place the boom
  With picViewArea
    .ScaleMode = 0 'user-scaling
    .ScaleWidth = 4 * UD.AC.SemiSpan
    .ScaleLeft = -2 * UD.AC.SemiSpan
    .ScaleHeight = -4 * UD.AC.SemiSpan
    .ScaleTop = 2 * UD.AC.SemiSpan
  End With
  
  'Place the boom, about which all other elements are placed.
  With linBoom
    .X1 = -UD.AC.SemiSpan
    .X2 = UD.AC.SemiSpan
    .Y1 = 0
    .Y2 = 0
  End With
  
  'add a reference centerline
  With linCenterLineV
    .X1 = 0
    .X2 = 0
    If UD.AC.WingType = 4 And cboView.ListIndex = 1 Then 'heli, top view
      .Y1 = -UD.AC.SemiSpan
      .Y2 = UD.AC.SemiSpan
    Else
      .Y1 = -0.125 * UD.AC.SemiSpan
      .Y2 = 0.125 * UD.AC.SemiSpan
    End If
  End With
  With linCenterLineH
    If UD.AC.WingType = 4 And cboView.ListIndex = 1 Then 'heli, top view
      .X1 = -UD.AC.SemiSpan
      .X2 = UD.AC.SemiSpan
    Else
      .X1 = -0.125 * UD.AC.SemiSpan
      .X2 = 0.125 * UD.AC.SemiSpan
    End If
    If cboView.ListIndex = 1 Then 'top view
      .Y1 = 0 '-UD.AC.BoomFwd
      .Y2 = 0 '-UD.AC.BoomFwd
    Else
      .Y1 = 0 '-UD.AC.BoomVert
      .Y2 = 0 '-UD.AC.BoomVert
    End If
  End With

  'draw the wing or rotor for the top view
  With shpWing
    If cboView.ListIndex = 1 Then 'Top View
      If UD.AC.WingType = 3 Then 'fixed-wing
        .Shape = 4 'rounded rectangle
        .Height = 0.2 * UD.AC.SemiSpan
        .Top = -UD.AC.BoomFwd + .Height
      ElseIf UD.AC.WingType = 4 Then 'helicopter
        .Shape = 3 'circle
        .Height = UD.AC.SemiSpan * 2
        .Top = -UD.AC.BoomFwd + UD.AC.SemiSpan
      End If
    Else  'rear view
      .Shape = 0 'rectangle
      .Height = 0.03 * UD.AC.SemiSpan
      .Top = -UD.AC.BoomVert + .Height
    End If
    .Left = -UD.AC.SemiSpan
    .Width = UD.AC.SemiSpan * 2
  End With
  
  'return to twip scaling (all the controls stay put)
  picViewArea.ScaleMode = 1 'twips
  
  'labels
  lblLeft.Left = linBoom.X1 + 120
  lblLeft.Top = linBoom.Y1 - lblLeft.Height - 120
  lblRight.Left = linBoom.X2 - lblRight.Width - 120
  lblRight.Top = linBoom.Y2 - lblRight.Height - 120
  If UD.AC.WingType = 3 Then 'fixed wing
    lblCenter.Caption = "Wing"
  ElseIf UD.AC.WingType = 4 Then 'helicopter
    lblCenter.Caption = "Rotor"
  End If
'tbc  lblcenter.Left = (linBoom.X1 + linBoom.X2 - lblcenter.Width) * 0.5
  lblCenter.Left = (linBoom.X1 + linBoom.X2) * 0.5 + lblUp.Width
  lblCenter.Top = shpWing.Top - lblCenter.Height - 120
  If cboView.ListIndex = 0 Then 'rear view
    lblUp.Caption = "Up"
    lblUp.Left = (linBoom.X1 + linBoom.X2 - lblUp.Width) * 0.5
    lblUp.Top = linBoom.Y1 - lblCenter.Height - 500
  ElseIf cboView.ListIndex = 1 Then 'top view
    lblUp.Caption = "Forward"
    lblUp.Left = (linBoom.X1 + linBoom.X2 - lblUp.Width) * 0.5
    If UD.AC.WingType = 3 Then 'fixed
      lblUp.Top = shpWing.Top - lblUp.Height - 120
    ElseIf UD.AC.WingType = 4 Then
      lblUp.Top = shpWing.Top + 120
    End If
  End If
  
  'nozzles
  UpdateNozzlePictureAll
End Sub

Private Function AlreadySelected(newnoz)
'Add a nozzle to the Selected list
  Dim i As Integer
  
  'see if the nozzle is already in the list
  AlreadySelected = False
  For i = 0 To NumCurrentNozzles - 1
    If CurrentNozzles(i) = newnoz Then
      AlreadySelected = True
      Exit For
    End If
  Next
End Function
