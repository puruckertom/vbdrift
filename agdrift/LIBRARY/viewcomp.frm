VERSION 4.00
Begin VB.Form frmViewComponent 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Component"
   ClientHeight    =   4320
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   4395
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   4725
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   4395
   Top             =   1140
   Width           =   4515
   Begin VB.ListBox lstComponent 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtNcomp 
      Appearance      =   0  'Flat
      DataField       =   "NumComponents"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "number"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtSubst 
      Appearance      =   0  'Flat
      DataField       =   "Substance"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "Substance"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Component"
      Connect         =   ""
      DatabaseName    =   ""
      Exclusive       =   0   'False
      Height          =   270
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Components"
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3840
      Width           =   855
   End
End
Attribute VB_Name = "frmViewComponent"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Data1_Reposition()
  GetComponents
End Sub

Private Sub Form_Load()
  CenterForm Me
  Data1.DatabaseName = GD.DBDirPath & GD.DBFileName
End Sub

Private Sub GetComponents()
'extract the percentage data from the current record
'and stuff it into the label controls
   Dim tmpFieldC As Field
   Dim tmpFieldP As Field
   Dim tmpFieldN As Field
   Dim CompStr As String
   Dim FracStr As String
   Dim s As String
   ReDim dat(6) As Single

   'get field value (all strings are packed into one)
   Set tmpFieldC = Data1.Recordset.Fields("Component")
   Set tmpFieldP = Data1.Recordset.Fields("Percent")
   Set tmpFieldN = Data1.Recordset.Fields("NumComponents")
   
   CompStr = tmpFieldC.Value
   FracStr = tmpFieldP.Value
   num = Val(tmpFieldN.Value)

   'convert to array
   StringToArray dat(), FracStr
   
   'load up the form controls
   lstComponent.Clear
   For i = 0 To num - 1
     s = Trim$(Mid$(CompStr, i * 32 + 1, 32))
     s = s & " (" & Format$(dat(i)) & "%)"
     lstComponent.AddItem s
   Next
End Sub

