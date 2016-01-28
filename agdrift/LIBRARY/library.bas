Attribute VB_Name = "basLibrary"
'DBcompile - a database compiler for AGDRIFT
'
'Creates a database containing Agdrift's "libraries" from
'a series of ascii source files
'
'source files:
'  atomize.out
'    substance ID, nozzle ID, nozzle angle, wind speed, nozzle type, mass frac,... (32 mass fracs, some zero)
'  subst.out
'    substance ID, num components, %, comp, %, comp,... (up to 7 %/comp)
'  aircraft.out
'
'  evap.out
'    substance ID, evaporation rate, nonvolatile fraction
'  nd.out
'    substance ID, dynamic surface tension, shear viscosity, density, ElongVisc
'  nozzle.out
'    nozzle,f50,f141,f220,vmd,diameter
'  ars.out
'
'
'Database Structure:
'   Table: Dropsize        drop size distribution data with 32 mass fracs
'     Field: Substance     substance ID
'     Field: Nozzle        nozzle ID
'     Field: NozzleAngle   nozzle angle in degrees
'     Field: WindSpeed     wind speed in mph
'     Field: DSLflag       flag for AGDSL
'     Field: MassFrac      mass fractions
'   Table: SubstanceList   unique list of Substances
'     Field: Substance
'   Table: NozzleList      unique list of Nozzles
'     Field: Nozzle
'   Table: NozzleAngleList unique list of nozzle angles
'     Field: NozzleAngle
'   Table: WindSpeedList   unique list of wind speeds
'     Field: WindSpeed
'
'   Table: Components      list of components for each substance
'     Field: Substance
'     Field: NumComponents
'     Field: Component
'     Field: Percent
'   Table: ComponentList   list of all components
'     Field Component
'
'   Table: Evaporation
'     Field: Substance
'     Field: EvaporationRate
'     Field: NonvolFraction
'
'   Table: Nozzles
'     Field: Nozzle
'     Field: F50
'     Field: F141
'     Field: F220
'     Field: VMD
'     Field: Diameter
'
'   Table: Materials
'     Field: Subst
'     Field: DynSurfTens
'     Field: ShearVisc
'     Field: Density
'     Field: ElongVisc
'
'   Table: Aircraft
'     Field: Name
'     Field: Type
'     Field: SemiSpan
'     Field: TypSpeed
'     Field: BiplSep
'     Field: Weight
'     Field: DragCoef
'     Field: PlanArea
'     Field: PropEff
'     Field: PropRPM
'     Field: PropRad
'     Field: EngVert
'     Field: EngFwd
'     Field: EngHoriz
'
'   Table: Canopy
'     Field: Name
'     Field: B
'     Field: C

Public Const LIBRARYVERSION = 7  'Current library file version

'
' Global variables
'
Type SourceFiles
  airbasic As String  'Basic Aircraft descriptions
  aircraft As String  'All Aircraft Descriptions
  atomize As String   '
  atomizeFS As String '
  asaedep As String   'ASAE Deposition
  basicdrp As String  'Basic Dropsize distributions
  dropsize As String  'All Dropsize distributions
  evap As String      '
  fieldrun As String  'Field Trial Data
  nd As String        '
  nozbasic As String  'Basic Nozzle Data
  nozzle As String    'All Nozzle Data
  nozzletm As String  'Nozzle type and manufacturer
  arsnoz As String    'ARS Nozzles
  sngldep As String   'ASAE Single-Swath Deposition
  subst As String     'Substance component data
  canopy As String    'Canopy data (optical)
  infbasic As String 'Tier 1 information
End Type

Type GlobalData
  LibType As Integer     'Output Library Type 0=World 1=SDTF 2=MAA
  DBDirPath As String    'Path to directory that will contain Database
  DBFileName As String   'Database file name for output
  SrcPath As String      'Full path to directory containing source files
  Src As SourceFiles     'Source file names
End Type

Global GD As GlobalData
'
' external funtion declaration
'
#If Win32 Then
  Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, _
    hpvSource As Any, ByVal cbCopy As Long)
#Else
  Declare Sub CopyMemory Lib "Kernel" Alias "hmemcpy" (hpvDest As Any, _
    hpvSource As Any, ByVal cbCopy As Long)
#End If

Sub AddAircraftRecords(TableName As String, SrcFile As String)
'read the Evaporation source file and stuff the data into tables
  
  Dim fn As String
  Dim DB As Database

  'local variables for record data storage
  Dim FD1 As String   'name
  Dim FD2 As Integer  'type
  Dim FD3 As Single   'semi-span (m)
  Dim FD4 As Single   'typ. Speed (m/s)
  Dim FD5 As Single   'biplane distance between wings (m)
  Dim FD6 As Single   'weight (kg)
  Dim FD7 As Single   'planform area (m2)
  Dim FD8 As Single   'prop/rotor rpm
  Dim FD9 As Single   'prop/rotor radius (m)
  Dim FD10 As Single  'engine vertical location (m)
  Dim FD11 As Single  'engine forward location (m)
  Dim FD12 As Integer 'number of engines
  Dim FD13(1) As Single  'engine horizontal locations (m)
'  Dim FD14 As Single
  Dim FD15 As Single  'dist from wingtip vortex to trailing edge (m)
  Dim FD16 As Single  'boom vertical dist to trailing edge (m)
  Dim FD17 As Single  'boom forward dist to trailing edge (m)
  Dim FD18 As Integer 'Source: 1=old blue book 2=new blue book 3=both

  Dim DS As Dynaset
  
  'Open the Database (Library) file
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset(TableName)

  'open the source file
  fn = GD.SrcPath & SrcFile
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Input #1, FD1, FD2, FD3, FD4, FD5, FD6, FD7, FD8, FD9, FD10, _
              FD11, FD12, FD13(0), FD13(1), FD15, FD16, FD17, F18

    'Add the new record to the dropsize table
    DS.AddNew
    DS.Fields("Name") = FD1
    DS.Fields("Type") = FD2
    DS.Fields("SemiSpan") = FD3
    DS.Fields("TypSpeed") = FD4
    DS.Fields("BiplSep") = FD5
    DS.Fields("Weight") = FD6
    DS.Fields("PlanArea") = FD7
    DS.Fields("PropRPM") = FD8
    DS.Fields("PropRad") = FD9
    DS.Fields("EngVert") = FD10
    DS.Fields("EngFwd") = FD11
    DS.Fields("NumEng") = FD12
    ArrayToField DS.Fields("EngHoriz"), FD13(), 2
    DS.Fields("WingVert") = FD15
    DS.Fields("BoomVert") = FD16
    DS.Fields("BoomFwd") = FD17

    DS.Update

  Wend
  Close #1
  
  'close the dynaset
  DS.Close
  
  AddToLog Format$(numrecs) + " aircraft records added"

  DB.Close
End Sub

Sub AddAircraftTables(TableName As String)
'Add tables for Evaporation info to database
'
'*********************************************************
' Database:
'   Table: Aircraft
'     Field: Name
'     Field: Type
'     Field: SemiSpan
'     Field: TypSpeed
'     Field: BiplSep
'     Field: Weight
'     Field: PlanArea
'     Field: PropRPM
'     Field: PropRad
'     Field: EngVert
'     Field: EngFwd
'     Field: NumEng
'     Field: EngHoriz
'     Field: WingVert
'     Field: BoomVert
'     Field: BoomFwd
'
  Dim DB As Database
  Dim TD As TableDef
  Dim FD As Field

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Dropsize table ****************************************
  Set TD = New TableDef
  TD.Name = TableName        'name the new table
  
  Set FD = New Field
  FD.Name = "Name"        'name the field
  FD.Type = dbText       'type the field
  FD.Size = 40            'size the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "Type"        'name the field
  FD.Type = dbInteger    'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "SemiSpan"    'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "TypSpeed"    'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "BiplSep"     'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "Weight"      'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "PlanArea"    'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "PropRPM"     'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "PropRad"     'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "EngVert"     'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "EngFwd"      'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "NumEng"      'name the field
  FD.Type = dbInteger    'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "EngHoriz"    'name the field
  FD.Type = dbLongBinary 'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "WingVert"    'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "BoomVert"    'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "BoomFwd"     'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  DB.TableDefs.Append TD  'append the table to the database

  'close the database
  DB.Close
End Sub

Sub AddARSRecords()
'read the ARS source file and stuff the data into tables
  
  Dim fn As String
  Dim DB As Database

  'local variables for record data storage
  Dim aname As String * 48
  Dim iindex As Integer
  Dim ofcunits As Integer
  Dim ofclabel As String * 40
  Dim ofcpreflag As Integer
  Dim ofcprefix As String * 2
  Dim numofc As Integer
  Dim ofcval(20) As Single
  Dim modunits As Integer
  Dim modlabel As String * 40
  Dim modpreflag As Integer
  Dim modprefix As String * 2
  Dim nummod As Integer
  Dim modval(20) As Single

  Dim DS As Dynaset
  
  'Open the Database (Library) file
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset("ARSNozzle")

  'open the source file
  fn = GD.SrcPath & GD.Src.arsnoz
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Input #1, iindex, aname          'sequential index (not used), Nozzle name
    Input #1, ofcunits, ofclabel     'orifice units (0=number 1=inches)
    Input #1, ofcpreflag, ofcprefix  'orifice prefix flag (0=none, 1=prefix), orifice prefix label
    Input #1, numofc                 'Number of discrete orifice values (0=range)
    If numofc = 0 Then 'range
      Input #1, ofcval(0), ofcval(1) 'min, max of range
    Else               'discrete values
      For i = 0 To numofc - 1
        Input #1, ofcval(i)          'discrete value
      Next
    End If
    Input #1, modunits, modlabel     'modifier units (0=number 1=degrees)
    Input #1, modpreflag, modprefix  'modifier prefix flag (0=none, 1=prefix), modifier prefix label
    Input #1, nummod                 'Number of discrete modifier values (0=range)
    If nummod = 0 Then
      Input #1, modval(0), modval(1) 'min, max range values
    Else
      For i = 0 To nummod - 1
        Input #1, modval(i)          'discrete value
      Next
    End If

    'Add the new record to the table
    DS.AddNew
    DS.Fields("Index") = iindex
    DS.Fields("Name") = aname
    DS.Fields("OfcUnits") = ofcunits
    DS.Fields("OfcLabel") = ofclabel
    DS.Fields("OfcPreflag") = ofcpreflag
    DS.Fields("OfcPrefix") = ofcprefix
    DS.Fields("NumOfc") = numofc
    If numofc = 0 Then
      ArrayToField DS.Fields("OfcVal"), ofcval(), 2
    Else
      ArrayToField DS.Fields("OfcVal"), ofcval(), numofc
    End If
    DS.Fields("ModUnits") = modunits
    DS.Fields("ModLabel") = modlabel
    DS.Fields("ModPreflag") = modpreflag
    DS.Fields("ModPrefix") = modprefix
    DS.Fields("NumMod") = nummod
    If nummod = 0 Then
      ArrayToField DS.Fields("ModVal"), modval(), 2
    Else
      ArrayToField DS.Fields("ModVal"), modval(), nummod
    End If

    DS.Update

  Wend
  Close #1
  
  'close the dynaset
  DS.Close
  
  AddToLog Format$(numrecs) + " ARS Nozzle records added"

  DB.Close
End Sub

Sub AddARSTables()
'Add tables for ARS info to database
'
'*********************************************************
' Database:
'   Table: ARSNozzle
'     Field: Index      'ARS Nozzle Index
'     Field: Name       'ARS Nozzle name
'     Field: OfcUnits   'Orifice units flag 0=none 1=cm/in
'     Field: OfcLabel   'Orifice value label "Orifice Size", etc.
'     Field: OfcPreflag 'Orifice prefix flag 0=none 1=prefix
'     Field: OfcPrefix  'Orifice prefix
'     Field: NumOfc     'Number of orifice values
'     Field: OfcVal     'Array of orifice values
'     Field: ModUnits   'Modifier units flag 0=none 1=deg
'     Field: ModLabel   'Modifier value label "Restrictor Number", etc.
'     Field: ModPreflag 'Orifice prefix flag 0=none 1=prefix
'     Field: ModPrefix  'Orifice prefix
'     Field: NumMod     'Number of modifier values
'     Field: ModVal     'Array of modifier values
'
  Dim DB As Database
  Dim TD As TableDef
  Dim FD As Field

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Dropsize table ****************************************
  Set TD = New TableDef
  TD.Name = "ARSNozzle"   'name the new table
  
  Set FD = New Field
  FD.Name = "Index"       'name the field
  FD.Type = dbInteger    'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "Name"        'name the field
  FD.Type = dbText       'type the field
  FD.Size = 48            'size the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "OfcUnits"    'name the field
  FD.Type = dbInteger    'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "OfcLabel"    'name the field
  FD.Type = dbText       'type the field
  FD.Size = 40            'size the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "OfcPreflag"  'name the field
  FD.Type = dbInteger     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "OfcPrefix"   'name the field
  FD.Type = dbText        'type the field
  FD.Size = 2             'size the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "NumOfc"      'name the field
  FD.Type = dbInteger     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "OfcVal"      'name the field
  FD.Type = dbLongBinary 'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "ModUnits"    'name the field
  FD.Type = dbInteger    'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "ModLabel"    'name the field
  FD.Type = dbText       'type the field
  FD.Size = 40            'size the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "ModPreflag"  'name the field
  FD.Type = dbInteger     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "ModPrefix"   'name the field
  FD.Type = dbText        'type the field
  FD.Size = 2             'size the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "NumMod"      'name the field
  FD.Type = dbInteger    'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "ModVal"      'name the field
  FD.Type = dbLongBinary 'type the field
  TD.Fields.Append FD     'append the field to the table
  
  DB.TableDefs.Append TD  'append the table to the database

  'close the database
  DB.Close
End Sub

Sub AddBasicRecords()
'read the Basic Data source file and stuff the data into tables
  
  Dim DB As Database
  Dim DS As Dynaset
  Dim fn As String

  'local variables for record data storage
  Dim aname As String * 36
  Dim itype As Integer
  Dim num As Integer
  Dim swathdispagpub As Single
  Dim swathdispagreg As Single
  Dim swathdispfs As Single
  Dim numel As Integer
  Dim diam() As Single
  Dim frac() As Single
  Dim semispan As Single
  Dim speed As Single
  Dim biplsep As Single
  Dim weight As Single
  Dim planarea As Single
  Dim proprpm As Single
  Dim proprad As Single
  Dim vert As Single
  Dim fwd As Single
  Dim neng As Integer
  Dim wingv As Single
  Dim boomv As Single
  Dim boomf As Single
  Dim dist() As Single
  Dim dep() As Single
  Dim horiz() As Single
  Dim imeth As Integer
  Dim sinfo As String
  Dim s As String

  'Open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  
' BasicDSD *************************************
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset("BasicDSD")

  'open the source file
  fn = GD.SrcPath & GD.Src.basicdrp
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  numel = 32
  ReDim diam(numel - 1)
  ReDim frac(numel - 1)
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Input #1, itype, num, swathdispagpub, swathdispagreg, swathdispfs, aname
    For i = 1 To numel
      Input #1, diam(i - 1)
    Next
    For i = 1 To numel
      Input #1, frac(i - 1)
    Next

    'Add the new record to the table
    DS.AddNew
    DS.Fields("Type") = itype
    DS.Fields("Name") = aname
    DS.Fields("SwathDispAgPub") = swathdispagpub
    DS.Fields("SwathDispAgReg") = swathdispagreg
    DS.Fields("SwathDispFS") = swathdispfs
    DS.Fields("NumDrop") = num
    ArrayToField DS.Fields("Diam"), diam(), numel
    ArrayToField DS.Fields("Frac"), frac(), numel
    DS.Update

  Wend
  Close #1
  AddToLog Format$(numrecs) + " Basic Tier2,3 DSD records added"
  
  'close the dynaset
  DS.Close
  
' BasicAC *********************************************
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset("BasicAC")

  'open the source file
  fn = GD.SrcPath & GD.Src.airbasic
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  ReDim horiz(1)
  numrecs = 0
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Input #1, aname, itype, semispan, speed, biplsep, weight, _
              planarea, proprpm, proprad, vert, fwd, _
              neng, horiz(0), horiz(1), wingv, boomv, boomf

    'Add the new record to the dropsize table
    DS.AddNew
    DS.Fields("Name") = aname
    DS.Fields("Type") = itype
    DS.Fields("SemiSpan") = semispan
    DS.Fields("TypSpeed") = speed
    DS.Fields("BiplSep") = biplsep
    DS.Fields("Weight") = weight
    DS.Fields("PlanArea") = planarea
    DS.Fields("PropRPM") = proprpm
    DS.Fields("PropRad") = proprad
    DS.Fields("EngVert") = vert
    DS.Fields("EngFwd") = fwd
    DS.Fields("NumEng") = neng
    ArrayToField DS.Fields("EngHoriz"), horiz(), 2
    DS.Fields("WingVert") = wingv
    DS.Fields("BoomVert") = boomv
    DS.Fields("BoomFwd") = boomf
    DS.Update

  Wend
  Close #1
  AddToLog Format$(numrecs) + " Basic AC records added"
  
  'close the dynaset
  DS.Close
  
' BasicDep *********************************************
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset("BasicDep")

  'open the source file
  fn = GD.SrcPath & GD.Src.asaedep
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  numel = 398 'was 840 'was 430
  ReDim dep(numel - 1)
  ReDim dist(numel - 1)
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Input #1, itype
    For i = 1 To numel
      Input #1, dep(i - 1)
      dist(i - 1) = (i - 1) * 2 'calc dist: every 2m
    Next

    'Add the new record to the deposition table
    DS.AddNew
    DS.Fields("Type") = itype
    DS.Fields("NumDep") = numel
    ArrayToField DS.Fields("Distance"), dist(), numel
    ArrayToField DS.Fields("Deposition"), dep(), numel
    DS.Update

  Wend
  Close #1
  AddToLog Format$(numrecs) + " Basic Dep records added"
  
  'close the dynaset
  DS.Close
  
' BasicSgl *********************************************
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset("BasicSgl")

  'open the source file
  fn = GD.SrcPath & GD.Src.sngldep
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  numel = 398 'was 375
  ReDim dep(numel - 1)
  ReDim dist(numel - 1)
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Input #1, itype
    For i = 1 To numel
      Input #1, dep(i - 1)
      dist(i - 1) = (i - 1) * 2 'calc dist: every 2m
    Next

    'Add the new record to the deposition table
    DS.AddNew
    DS.Fields("Type") = itype
    DS.Fields("NumDep") = numel
    ArrayToField DS.Fields("Distance"), dist(), numel
    ArrayToField DS.Fields("Deposition"), dep(), numel
    DS.Update

  Wend
  Close #1
  AddToLog Format$(numrecs) + " Basic Sgl records added"
  
  'close the dynaset
  DS.Close
  
' BasicNZ *********************************************
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset("BasicNZ")

  'open the source file
  fn = GD.SrcPath & GD.Src.nozbasic
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  numel = 60
  ReDim horiz(numel - 1)
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Input #1, itype, num
    For i = 1 To numel
      Input #1, horiz(i - 1)
    Next

    'Add the new record to the dropsize table
    DS.AddNew
    DS.Fields("Type") = itype
    DS.Fields("NumNoz") = num
    ArrayToField DS.Fields("PosHoriz"), horiz(), numel
    DS.Update

  Wend
  Close #1
  AddToLog Format$(numrecs) + " Basic NZ records added"
  
  'close the dynaset
  DS.Close
  
' BasicInfo *********************************************
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset("BasicInfo")

  'open the source file
  fn = GD.SrcPath & GD.Src.infbasic
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    sinfo = ""
    Input #1, imeth, itype, num
    For i = 1 To num
      Line Input #1, s
      sinfo = sinfo + s
      If i < num Then sinfo = sinfo + vbCrLf 'add a CRLF between lines
    Next

    'Add the new record to the table
    DS.AddNew
    DS.Fields("ApplMeth") = imeth
    DS.Fields("Type") = itype
    DS.Fields("Info") = sinfo
    DS.Update

  Wend
  Close #1
  AddToLog Format$(numrecs) + " Basic Info records added"
  
  'close the dynaset
  DS.Close
  
  'close the database
  DB.Close
End Sub

Sub AddBasicTables()
'Add tables for Basic Data to database
'
'*********************************************************
' Database:
'   Table: BasicDSD
'     Field: Type           Basic category
'     Field: Name           category name
'     Field: SwathDispAgPub Swath Displacement, SDTF, Public
'     Field: SwathDispAGReg Swath Displacement, SDTF, Regulatory
'     Field: SwathDispFS    Swath Displacement, FS
'     Field: NumDrop        number of drops
'     Field: Diam           diameters
'     Field: Frac           fractions

'   Table: BasicAC
'     Field: Type
'     Field: Name
'     Field: SemiSpan
'     Field: TypSpeed
'     Field: BiplSep
'     Field: Weight
'     Field: PlanArea
'     Field: PropRPM
'     Field: PropRad
'     Field: EngVert
'     Field: EngFwd
'     Field: NumEng
'     Field: EngHoriz
'     Field: WingVert
'     Field: BoomVert
'     Field: BoomFwd
'
'   Table: BasicDep
'     Field: Type       ASAE category
'     Field: NumDep     number of points
'     Field: Distance   downwind distance
'     Field: Deposition deposition
'
'   Table: BasicSgl
'     Field: Type       ASAE category
'     Field: NumDep     number of points
'     Field: Distance   downwind distance
'     Field: Deposition single-swath deposition
'
'   Table: BasicNZ
'     Field: Type
'     Field: NumNoz
'     Field: PosHoriz
'
'   Table: BasicInfo
'     Field: ApplMeth
'     Field: Type
'     Field: Info
'
  Dim DB As Database
  Dim TD As TableDef
  Dim FD As Field

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Basic DSD table **************************************
  Set TD = New TableDef      'create a new table
  TD.Name = "BasicDSD"       'name the new table
  
  Set FD = New Field         'create a new field
  FD.Name = "Type"           'name the field
  FD.Type = dbInteger       'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "Name"           'name the field
  FD.Type = dbText          'type the field
  FD.Size = 36               'set the field length
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "SwathDispAgPub" 'name the field
  FD.Type = dbSingle        'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "SwathDispAgReg" 'name the field
  FD.Type = dbSingle        'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "SwathDispFS"    'name the field
  FD.Type = dbSingle        'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "NumDrop"        'name the field
  FD.Type = dbInteger       'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "Diam"           'name the field
  FD.Type = dbLongBinary    'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "Frac"           'name the field
  FD.Type = dbLongBinary    'type the field
  TD.Fields.Append FD        'append the field to the table
  
  DB.TableDefs.Append TD     'append the table to the database

  'create BasicAC table ****************************************
  Set TD = New TableDef      'create a new table
  TD.Name = "BasicAC"        'name the new table
  
  Set FD = New Field         'create a new field
  FD.Name = "Type"           'name the field
  FD.Type = dbInteger       'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "Name"           'name the field
  FD.Type = dbText          'type the field
  FD.Size = 36               'set the field length
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "SemiSpan"       'name the field
  FD.Type = dbSingle        'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "TypSpeed"       'name the field
  FD.Type = dbSingle        'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "BiplSep"        'name the field
  FD.Type = dbSingle        'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "Weight"         'name the field
  FD.Type = dbSingle        'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "PlanArea"       'name the field
  FD.Type = dbSingle        'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "PropRPM"        'name the field
  FD.Type = dbSingle        'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "PropRad"        'name the field
  FD.Type = dbSingle        'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "EngVert"        'name the field
  FD.Type = dbSingle        'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "EngFwd"         'name the field
  FD.Type = dbSingle        'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field
  FD.Name = "NumEng"      'name the field
  FD.Type = dbInteger    'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "EngHoriz"    'name the field
  FD.Type = dbLongBinary 'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "WingVert"    'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "BoomVert"    'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "BoomFwd"     'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  DB.TableDefs.Append TD     'append the table to the database

  'create BasicDep table ****************************************
  Set TD = New TableDef      'create a new table
  TD.Name = "BasicDep"       'name the new table
  
  Set FD = New Field         'create a new field
  FD.Name = "Type"           'name the field
  FD.Type = dbInteger       'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "NumDep"         'name the field
  FD.Type = dbInteger       'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "Distance"       'name the field
  FD.Type = dbLongBinary    'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "Deposition"     'name the field
  FD.Type = dbLongBinary    'type the field
  TD.Fields.Append FD        'append the field to the table
  
  DB.TableDefs.Append TD     'append the table to the database

  'create BasicSgl table ****************************************
  Set TD = New TableDef      'create a new table
  TD.Name = "BasicSgl"       'name the new table
  
  Set FD = New Field         'create a new field
  FD.Name = "Type"           'name the field
  FD.Type = dbInteger       'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "NumDep"         'name the field
  FD.Type = dbInteger       'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "Distance"       'name the field
  FD.Type = dbLongBinary    'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "Deposition"     'name the field
  FD.Type = dbLongBinary    'type the field
  TD.Fields.Append FD        'append the field to the table
  
  DB.TableDefs.Append TD     'append the table to the database

  'create BasicNZ table ****************************************
  Set TD = New TableDef      'create a new table
  TD.Name = "BasicNZ"        'name the new table
  
  Set FD = New Field         'create a new field
  FD.Name = "Type"           'name the field
  FD.Type = dbInteger       'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "NumNoz"         'name the field
  FD.Type = dbInteger       'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "PosHoriz"       'name the field
  FD.Type = dbLongBinary    'type the field
  TD.Fields.Append FD        'append the field to the table
  
  DB.TableDefs.Append TD     'append the table to the database

  'create BasicInfo table ****************************************
  Set TD = New TableDef      'create a new table
  TD.Name = "BasicInfo"      'name the new table
  
  Set FD = New Field         'create a new field
  FD.Name = "ApplMeth"       'name the field
  FD.Type = dbInteger       'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "Type"           'name the field
  FD.Type = dbInteger       'type the field
  TD.Fields.Append FD        'append the field to the table
  
  Set FD = New Field         'create a new field
  FD.Name = "Info"           'name the field
  FD.Type = dbLongBinary    'type the field
  TD.Fields.Append FD        'append the field to the table
  
  DB.TableDefs.Append TD     'append the table to the database

  'close the database
  DB.Close
End Sub

Sub AddComponentRecords()
'read the Component source file and stuff the data into tables
  
  Dim DB As Database
  Dim fn As String

  'local variables for record data storage
  Dim subst As String      'substance
  Dim ncomp As Integer     'number of components
  ReDim comps(6) As String   'component list
  ReDim percents(6) As Single 'percentage list
  Dim bigcomp As String    'packed storage for components
  Dim ncomplist As Integer 'number of elements in comp_list
  ReDim comp_list(500) As String 'list of components
  
  Dim dssubst As Dynaset
  Dim dscomp As Dynaset
  
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  
  'create dynasets of records to work on
  Set dssubst = DB.CreateDynaset("Components")
  Set dscomp = DB.CreateDynaset("ComponentList")

  'open the source file
  fn = GD.SrcPath & GD.Src.subst
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  ncomlist = 0
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    bigcomp = ""
    Input #1, subst, ncomp
    For i = 0 To ncomp - 1
      Input #1, percents(i), comps(i)
      bigcomp = bigcomp & comps(i)
    Next
    
    bigcomp = ""
    For i = 0 To ncomp - 1
      bigcomp = bigcomp & comps(i) & String$(32 - Len(comps(i)), " ")
      Call ConditionalAddItem(comps(i), ncomplist, comp_list())
    Next

    'Add the new record to the dropsize table
    dssubst.AddNew
    dssubst.Fields("Substance") = subst
    dssubst.Fields("NumComponents") = ncomp
    dssubst.Fields("Component") = bigcomp
    ArrayToField dssubst.Fields("Percent"), percents(), ncomp
    dssubst.Update

  Wend
  Close #1
  
  'add the fieldlists to the tables
  For i = 0 To ncomplist - 1
    dscomp.AddNew
    dscomp.Fields(0) = comp_list(i)
    dscomp.Update
  Next
  
  'close the dynasets
  dssubst.Close
  dscomp.Close
  
  AddToLog Format$(numrecs) + " component records added"

  DB.Close
End Sub

Sub AddComponentTables()
'Add tables for Components to database
'
'*********************************************************
' Database:
'   Table: Components      list of components for each substance
'     Field: Substance
'     Field: NumComponents
'     Field: Component
'     Field: Percent
'   Table: ComponentList   list of all components
'     Field Component
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Components table *******************************
  Dim TDcom As New TableDef    'create a new table
  TDcom.Name = "Components"
  Dim FDsubs As New Field       'create a new field
  FDsubs.Name = "Substance"
  FDsubs.Type = dbText
  FDsubs.Size = 32
  TDcom.Fields.Append FDsubs
  Dim FDnc As New Field       'create a new field
  FDnc.Name = "NumComponents"
  FDnc.Type = dbInteger
  TDcom.Fields.Append FDnc
  Dim FDcoms As New Field       'create a new field
  FDcoms.Name = "Component"
  FDcoms.Type = dbText
  FDcoms.Size = 32 * 7
  TDcom.Fields.Append FDcoms
  Dim FDpers As New Field       'create a new field
  FDpers.Name = "Percent"
  FDpers.Type = dbLongBinary
  TDcom.Fields.Append FDpers
  DB.TableDefs.Append TDcom   'append the table to the database
  
  'create ComponentList table ************************************
  Dim TDsub As New TableDef    'create a new table
  TDsub.Name = "ComponentList"
  Dim FDsub1 As New Field       'create a new field
  FDsub1.Name = "Component"
  FDsub1.Type = dbText
  FDsub1.Size = 32
  TDsub.Fields.Append FDsub1
  DB.TableDefs.Append TDsub   'append the table to the database
  
  
  'close the completed database
  DB.Close
End Sub

Sub AddDropsizeRecords(FSflag As Integer)
'read the Dropsize source file and stuff the data into tables
  
  Dim DB As Database
  Dim fn As String

  'local variables for record data storage
  Dim subst As String      'substance
  Dim rpmflag As Integer   'angle/rpm flag
  Dim nozzle As String     'nozzle
  Dim nozang As Integer    'nozzle angle
  Dim nozrpm As Integer    'nozzle rpm
  Dim press As Single      'nozzle pressure
  Dim ws As Single         'wind speed
  Dim dslflag As Integer   'AGDSL flag
  ReDim mf(31) As Single   'mass fractions
  Dim spqual As Integer    'spray quality

  Dim ndiam As Integer     'number of diameters
  ReDim diam(32) As Single 'diameters
  
  Dim nsubst As Integer
  Dim nnozzle As Integer
  Dim nnozang As Integer
  Dim nnozrpm As Integer
  Dim npress As Integer
  Dim nws As Integer
  ReDim subst_list(500) As String
  ReDim nozzle_list(500) As String
  ReDim nozang_list(500) As String
  ReDim nozrpm_list(500) As String
  ReDim press_list(500) As String
  ReDim ws_list(500) As String

  Dim dsdrop As Dynaset
  Dim dsdiam As Dynaset
  Dim dssubst As Dynaset
  Dim dsnozzle As Dynaset
  Dim dsnozang As Dynaset
  Dim dsnozrpm As Dynaset
  Dim dspress As Dynaset
  Dim dsws As Dynaset
  
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create dynasets of records to work on
  If FSflag = 0 Then
    Set dsdrop = DB.CreateDynaset("Dropsize")
    Set dssubst = DB.CreateDynaset("SubstanceList")
    Set dsnozzle = DB.CreateDynaset("NozzleList")
    Set dsnozang = DB.CreateDynaset("NozzleAngleList")
    Set dsnozrpm = DB.CreateDynaset("NozzleRPMList")
    Set dspress = DB.CreateDynaset("PressureList")
    Set dsws = DB.CreateDynaset("WindSpeedList")
    fn = GD.SrcPath & GD.Src.atomize
  Else
    Set dsdrop = DB.CreateDynaset("DropsizeFS")
    Set dssubst = DB.CreateDynaset("SubstanceListFS")
    Set dsnozzle = DB.CreateDynaset("NozzleListFS")
    Set dsnozang = DB.CreateDynaset("NozzleAngleListFS")
    Set dsnozrpm = DB.CreateDynaset("NozzleRPMListFS")
    Set dspress = DB.CreateDynaset("PressureListFS")
    Set dsws = DB.CreateDynaset("WindSpeedListFS")
    fn = GD.SrcPath & GD.Src.atomizeFS
  End If
  'open the source file
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  nsubst = 0
  ndesc = 0
  nnozzle = 0
  nnozang = 0
  nnozrpm = 0
  npress = 0
  nws = 0
  numrecs = 0
  
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    If FSflag = 0 Then
      Input #1, subst, nozzle, nozang, press, ws, dslflag
      nozrpm = 0
    Else
      Input #1, subst, rpmflag, nozzle, nozang, press, ws, dslflag
      If rpmflag = 1 Then
        nozrpm = nozang
        nozang = 0
      Else
        nozrpm = 0
      End If
    End If
    For i = 0 To 31
      Input #1, mf(i)
    Next
    
    Input #1, spqual

    'Add the new record to the dropsize table
    dsdrop.AddNew
    dsdrop.Fields("Substance") = subst
    dsdrop.Fields("Nozzle") = nozzle
    dsdrop.Fields("NozzleAngle") = nozang
    dsdrop.Fields("NozzleRPM") = nozrpm
    dsdrop.Fields("Pressure") = press
    dsdrop.Fields("WindSpeed") = ws
    dsdrop.Fields("DSLflag") = dslflag
    ArrayToField dsdrop.Fields("MassFrac"), mf(), 32
    dsdrop.Fields("SprayQuality") = spqual
    dsdrop.Update

    'add the fields to the fieldlists
    ConditionalAddItem subst, nsubst, subst_list()
    ConditionalAddItem nozzle, nnozzle, nozzle_list()
    ConditionalAddItem str$(nozang), nnozang, nozang_list()
    ConditionalAddItem str$(nozrpm), nnozrpm, nozrpm_list()
    ConditionalAddItem str$(press), npress, press_list()
    ConditionalAddItem str$(ws), nws, ws_list()
  Wend
  Close #1
  
  'add the fieldlists to the tables
  For i = 0 To nsubst - 1
    dssubst.AddNew
    dssubst.Fields(0) = subst_list(i)
    dssubst.Update
  Next
  For i = 0 To nnozzle - 1
    dsnozzle.AddNew
    dsnozzle.Fields(0) = nozzle_list(i)
    dsnozzle.Update
  Next
  For i = 0 To nnozang - 1
    dsnozang.AddNew
    dsnozang.Fields(0) = nozang_list(i)
    dsnozang.Update
  Next
  For i = 0 To nnozrpm - 1
    dsnozrpm.AddNew
    dsnozrpm.Fields(0) = nozrpm_list(i)
    dsnozrpm.Update
  Next
  For i = 0 To npress - 1
    dspress.AddNew
    dspress.Fields(0) = press_list(i)
    dspress.Update
  Next
  For i = 0 To nws - 1
    dsws.AddNew
    dsws.Fields(0) = ws_list(i)
    dsws.Update
  Next
  
  'close the dynasets
  dsdrop.Close
  dssubst.Close
  dsnozzle.Close
  dsnozang.Close
  dsnozrpm.Close
  dspress.Close
  dsws.Close

  If FSflag = 0 Then
    AddToLog Format$(numrecs) + " dropsize records added"
  Else
    AddToLog Format$(numrecs) + " dropsize records added (FS)"
  End If

  'Now do diameter data
  If FSflag = 0 Then
    Set dsdiam = DB.CreateDynaset("Dropdiam")
  Else
    Set dsdiam = DB.CreateDynaset("DropdiamFS")
  End If
  
  fn = GD.SrcPath & GD.Src.dropsize
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  numrecs = 0
  While Not EOF(1)
    numrecs = numrecs + 1
    Input #1, dslflag, ndiam   'key matching diam to massfrac
    For i = 0 To ndiam - 1
      Input #1, diam(i)
    Next
    
    'Add the new record to the dropsize table
    dsdiam.AddNew
    dsdiam.Fields("DSLflag") = dslflag
    ArrayToField dsdiam.Fields("Diameter"), diam(), ndiam
    dsdiam.Update
  Wend
  Close #1
  dsdiam.Close
  If FSflag = 0 Then
    AddToLog Format$(numrecs) + " dropdiam records added"
  Else
    AddToLog Format$(numrecs) + " dropdiam records added (FS)"
  End If
  
  DB.Close
End Sub

Sub AddDropsizeTables(FSflag As Integer)
'Add tabes for Dropsize info to database
'
'*********************************************************
' Database:
'   Table: Dropsize        all the data with 32 mass fracs
'     Field: Substance
'     Field: Nozzle
'     Field: NozzleAngle   nozzle angle (deg)
'     Field: NozzleRPM     nozzle RPM
'     Field: Pressure      nozzle pressure (bar)
'     Field: WindSpeed     wind speed (m/s)
'     Field: DSLflag       flag for AGDSL
'     Field: MassFrac      mass fractions
'     Field: SprayQuality  spray quality
'   Table: Dropdiam        drop diameters for above
'     Field: DSLflag       key to match above
'     Field: Diameter      drop diameters
'   Table: SubstanceList   unique list of Substances
'     Field: Substance
'   Table: NozzleList      unique list of Nozzles
'     Field: Nozzle
'   Table: NozzleAngleList unique list of nozzle angles
'     Field: NozzleAngle
'   Table: NozzleRPMList   unique list of nozzle RPMs
'     Field: NozzleRPM
'   Table: PressureList    unique list of nozzle pressures
'     Field: Pressure
'   Table: WindSpeedList   unique list of wind speeds
'     Field: WindSpeed
'
'
  Dim DB As Database
  Dim TD As TableDef
  Dim FD As Field
  Dim FSstr As String
  

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Dropsize table ****************************************
  If FSflag = 0 Then
    FSstr = ""
  Else
    FSstr = "FS"
  End If
  Set TD = New TableDef    'create a new table
  TD.Name = "Dropsize" + FSstr   'name the new table
  
  Set FD = New Field       'create a new field
  FD.Name = "Substance"    'name the field
  FD.Type = dbText        'type the field
  FD.Size = 30             'size the field
  TD.Fields.Append FD      'append the field to the table
  
  Set FD = New Field       'create a new field
  FD.Name = "Nozzle"       'name the field
  FD.Type = dbText        'type the field
  FD.Size = 20             'size the field
  TD.Fields.Append FD      'append the field to the table
  
  Set FD = New Field       'create a new field
  FD.Name = "NozzleAngle"  'name the field
  FD.Type = dbInteger     'type the field
  TD.Fields.Append FD      'append the field to the table
  
  Set FD = New Field       'create a new field
  FD.Name = "NozzleRPM"    'name the field
  FD.Type = dbInteger     'type the field
  TD.Fields.Append FD      'append the field to the table
  
  Set FD = New Field       'create a new field
  FD.Name = "Pressure"     'name the field
  FD.Type = dbSingle      'type the field
  TD.Fields.Append FD      'append the field to the table
  
  Set FD = New Field       'create a new field
  FD.Name = "WindSpeed"    'name the field
  FD.Type = dbSingle      'type the field
  TD.Fields.Append FD      'append the field to the table
  
  Set FD = New Field       'create a new field
  FD.Name = "DSLflag"      'name the field
  FD.Type = dbInteger     'type the field
  TD.Fields.Append FD      'append the field to the table
  
  Set FD = New Field       'create a new field
  FD.Name = "MassFrac"     'name the field
  FD.Type = dbLongBinary  'type the field
  TD.Fields.Append FD      'append the field to the table

  Set FD = New Field       'create a new field
  FD.Name = "SprayQuality" 'name the field
  FD.Type = dbInteger     'type the field
  TD.Fields.Append FD      'append the field to the table
  
  DB.TableDefs.Append TD   'append the table to the database
  
  'create Dropdiam table ****************************************
  Set TD = New TableDef    'create a new table
  TD.Name = "Dropdiam" + FSstr     'name the new table
  
  Set FD = New Field       'create a new field
  FD.Name = "DSLflag"      'name the field
  FD.Type = dbInteger     'type the field
  TD.Fields.Append FD      'append the field to the table
  
  Set FD = New Field       'create a new field
  FD.Name = "Diameter"     'name the field
  FD.Type = dbLongBinary  'type the field
  TD.Fields.Append FD      'append the field to the table

  DB.TableDefs.Append TD   'append the table to the database
  
  'create SubstanceList table ************************************
  Set TD = New TableDef    'create a new table
  TD.Name = "SubstanceList" + FSstr
  Set FD = New Field       'create a new field
  FD.Name = "Substance"
  FD.Type = dbText
  FD.Size = 30
  TD.Fields.Append FD
  DB.TableDefs.Append TD   'append the table to the database
  
  'create NozzleList table ***************************************
  Set TD = New TableDef    'create a new table
  TD.Name = "NozzleList" + FSstr
  Set FD = New Field       'create a new field
  FD.Name = "Nozzle"
  FD.Type = dbText
  FD.Size = 20
  TD.Fields.Append FD
  DB.TableDefs.Append TD   'append the table to the database
  
  'create NozzleAngleList table **********************************
  Set TD = New TableDef    'create a new table
  TD.Name = "NozzleAngleList" + FSstr
  Set FD = New Field       'create a new field
  FD.Name = "NozzleAngle"
  FD.Type = dbInteger
  TD.Fields.Append FD
  DB.TableDefs.Append TD   'append the table to the database
  
  'create NozzleRPMList table **********************************
  Set TD = New TableDef    'create a new table
  TD.Name = "NozzleRPMList" + FSstr
  Set FD = New Field       'create a new field
  FD.Name = "NozzleRPM"
  FD.Type = dbInteger
  TD.Fields.Append FD
  DB.TableDefs.Append TD   'append the table to the database
  
  'create PressureList table ************************************
  Set TD = New TableDef    'create a new table
  TD.Name = "PressureList" + FSstr
  Set FD = New Field       'create a new field
  FD.Name = "Pressure"
  FD.Type = dbSingle
  TD.Fields.Append FD
  DB.TableDefs.Append TD   'append the table to the database

  'create WindSpeedList table ************************************
  Set TD = New TableDef    'create a new table
  TD.Name = "WindSpeedList" + FSstr
  Set FD = New Field       'create a new field
  FD.Name = "WindSpeed"
  FD.Type = dbSingle
  TD.Fields.Append FD
  DB.TableDefs.Append TD   'append the table to the database

  'close the database
  DB.Close
End Sub

Sub AddEvaporationRecords()
'read the Evaporation source file and stuff the data into tables
  
  Dim DB As Database
  Dim fn As String

  'local variables for record data storage
  Dim subst As String      'substance
  Dim erate As Single      'Evaporation rate
  Dim nvfrac As Single     'Nonvolatile Fraction
  
  Dim dsevap As Dynaset
  
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  
  'create dynaset of records to work on
  Set dsevap = DB.CreateDynaset("Evaporation")

  'open the source file
  fn = GD.SrcPath & GD.Src.evap
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Input #1, subst, erate, nvfrac

    'Add the new record to the dropsize table
    dsevap.AddNew
    dsevap.Fields("Substance") = subst
    dsevap.Fields("EvaporationRate") = erate
    dsevap.Fields("NonvolFraction") = nvfrac
    dsevap.Update

  Wend
  Close #1
  
  'close the dynaset
  dsevap.Close
  
  AddToLog Format$(numrecs) + " evaporation records added"

  DB.Close
End Sub

Sub AddEvaporationTables()
'Add tables for Evaporation info to database
'
'*********************************************************
' Database:
'   Table: Evaporation
'     Field: Substance
'     Field: EvaporationRate
'     Field: NonvolFraction
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Dropsize table ****************************************
  Dim TDevp As New TableDef    'create a new table
  TDevp.Name = "Evaporation"   'name the new table
  
  Dim FDsub As New Field       'create a new field
  FDsub.Name = "Substance"     'name the field
  FDsub.Type = dbText         'type the field
  FDsub.Size = 10              'size the field
  TDevp.Fields.Append FDsub    'append the field to the table
  
  Dim FDrate As New Field      'create a new field
  FDrate.Name = "EvaporationRate" 'name the field
  FDrate.Type = dbSingle      'type the field
  TDevp.Fields.Append FDrate   'append the field to the table
  
  Dim FDsg As New Field        'create a new field
  FDsg.Name = "NonvolFraction" 'name the field
  FDsg.Type = dbSingle        'type the field
  TDevp.Fields.Append FDsg     'append the field to the table
  
  DB.TableDefs.Append TDevp    'append the table to the database

  'close the database
  DB.Close
End Sub

Sub AddMaterialRecords()
'read the Nozzle source file and stuff the data into tables
  
  Dim DB As Database
  Dim fn As String

  'local variables for record data storage
  Dim Substance As String      'substance
  Dim DynSurfTens As Single
  Dim ShearVisc As Single
  Dim Density As Single
  Dim ElongVisc As Single
  
  Dim dsmat As Dynaset
  
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  
  'create dynaset of records to work on
  Set dsmat = DB.CreateDynaset("Materials")

  'open the source file
  fn = GD.SrcPath & GD.Src.nd
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Input #1, Substance, DynSurfTens, ShearVisc, Density, ElongVisc

    'Add the new record to the dropsize table
    dsmat.AddNew
    dsmat.Fields("Substance") = Substance
    dsmat.Fields("DynSurfTens") = DynSurfTens
    dsmat.Fields("ShearVisc") = ShearVisc
    dsmat.Fields("Density") = Density
    dsmat.Fields("ElongVisc") = ElongVisc
    dsmat.Update

  Wend
  Close #1
  
  'close the dynaset
  dsmat.Close
  
  AddToLog Format$(numrecs) + " material records added"

  DB.Close
End Sub

Sub AddMaterialTables()
'Add tables for Materials info to database
'
'*********************************************************
' Database:
'   Table: Materials
'     Field: Substance
'     Field: DynSurfTens
'     Field: ShearVisc
'     Field: Density
'     Field: ElongVisc
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Dropsize table ****************************************
  Dim TDmat As New TableDef    'create a new table
  TDmat.Name = "Materials"     'name the new table
  
  Dim FD1 As New Field         'create a new field
  FD1.Name = "Substance"       'name the field
  FD1.Type = dbText           'type the field
  FD1.Size = 10                'size the field
  TDmat.Fields.Append FD1      'append the field to the table
  
  Dim FD2 As New Field         'create a new field
  FD2.Name = "DynSurfTens"     'name the field
  FD2.Type = dbSingle         'type the field
  TDmat.Fields.Append FD2      'append the field to the table
  
  Dim FD3 As New Field         'create a new field
  FD3.Name = "ShearVisc"       'name the field
  FD3.Type = dbSingle         'type the field
  TDmat.Fields.Append FD3      'append the field to the table
  
  Dim FD4 As New Field         'create a new field
  FD4.Name = "ElongVisc"         'name the field
  FD4.Type = dbSingle         'type the field
  TDmat.Fields.Append FD4      'append the field to the table
  
  Dim FD5 As New Field         'create a new field
  FD5.Name = "Density"         'name the field
  FD5.Type = dbSingle         'type the field
  TDmat.Fields.Append FD5      'append the field to the table
  
  DB.TableDefs.Append TDmat    'append the table to the database

  'close the database
  DB.Close
End Sub

Sub AddInfoRecords()
  Dim DB As Database
  Dim fn As String

  Dim DS As Dynaset
  
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset("Info")

  'Add the new record to the dropsize table
  DS.AddNew
  DS.Fields("Version") = LIBRARYVERSION
  DS.Update

  'close the dynaset
  DS.Close
  
  AddToLog "1 info record added"

  DB.Close
End Sub

Public Sub AddInfoTables()
'Add tables for Info to database
'
'*********************************************************
' Database:
'   Table: Info
'     Field: Version

  Dim DB As Database
  Dim TD As TableDef
  Dim FD As Field

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Basic DSD table ****************************************
  Set TD = New TableDef      'create a new table
  TD.Name = "Info"           'name the new table
  
  Set FD = New Field
  FD.Name = "Version"
  FD.Type = dbInteger
  TD.Fields.Append FD

  DB.TableDefs.Append TD     'append the table to the database

  'close the database
  DB.Close
End Sub

Sub AddNozzleRecords()
'read the Nozzle source file and stuff the data into tables
  
  Dim DB As Database
  Dim fn As String

  'local variables for record data storage
  Dim nozzle As String
  Dim VMD As Single
  Dim RelSpan As Single
  Dim Diameter As Single
  Dim SprayAngle As Single
  
  Dim DS As Dynaset
  
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset("Nozzles")

  'open the source file
  fn = GD.SrcPath & GD.Src.nozzle
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Input #1, nozzle, VMD, RelSpan, Diameter, SprayAngle

    'Add the new record to the dropsize table
    DS.AddNew
    DS.Fields("Nozzle") = nozzle
    DS.Fields("VMD") = VMD
    DS.Fields("RelSpan") = RelSpan
    DS.Fields("Diameter") = Diameter
    DS.Fields("SprayAngle") = SprayAngle
    DS.Update

  Wend
  
  Close #1
  DS.Close
  AddToLog Format$(numrecs) + " nozzle records added"
  DB.Close
End Sub

Public Sub AddNozzleTables()
'Add tables for Evaporation info to database
'
'*********************************************************
' Database:
'   Table: Nozzles
'     Field: Nozzle
'     Field: VMD
'     Field: RelSpan
'     Field: Diameter
'     Field: SprayAngle
'
  Dim DB As Database
  Dim TD As TableDef
  Dim FD As Field

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Dropsize table ****************************************
  Set TD = New TableDef    'create a new table
  TD.Name = "Nozzles"       'name the new table
  
  Set FD = New Field        'create a new field
  FD.Name = "Nozzle"        'name the field
  FD.Type = dbText         'type the field
  FD.Size = 20              'size the field
  TD.Fields.Append FD       'append the field to the table
  
  Set FD = New Field        'create a new field
  FD.Name = "VMD"           'name the field
  FD.Type = dbSingle       'type the field
  TD.Fields.Append FD       'append the field to the table
  
  Set FD = New Field        'create a new field
  FD.Name = "RelSpan"         'name the field
  FD.Type = dbSingle       'type the field
  TD.Fields.Append FD       'append the field to the table
  
  Set FD = New Field        'create a new field
  FD.Name = "Diameter"      'name the field
  FD.Type = dbSingle       'type the field
  TD.Fields.Append FD       'append the field to the table
  
  Set FD = New Field        'create a new field
  FD.Name = "SprayAngle"    'name the field
  FD.Type = dbSingle       'type the field
  TD.Fields.Append FD       'append the field to the table
  
  DB.TableDefs.Append TD    'append the table to the database

  'close the database
  DB.Close
End Sub

Sub AddNozzleTMRecords()
'read the Nozzle source file and stuff the data into tables
  
  Dim DB As Database
  Dim fn As String

  'local variables for record data storage
  Dim nozzle As String
  Dim noztype As String
  Dim mfg As String
  
  Dim DS As Dynaset
  
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset("NozzlesTypMfg")

  'open the source file
  fn = GD.SrcPath & GD.Src.nozzletm
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Input #1, nozzle, noztype, mfg

    'Add the new record to the dropsize table
    DS.AddNew
    DS.Fields("Nozzle") = nozzle
    DS.Fields("Type") = noztype
    DS.Fields("Manufacturer") = mfg
    DS.Update

  Wend
  
  Close #1
  DS.Close
  AddToLog Format$(numrecs) + " nozzle type/mfg records added"
  DB.Close
End Sub

Public Sub AddNozzleTMTables()
'Add tables to database
'
'*********************************************************
' Database:
'   Table: NozzlesTypMfg
'     Field: Nozzle
'     Field: Type
'     Field: Manufacturer
'
  Dim DB As Database
  Dim TD As TableDef
  Dim FD As Field

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Dropsize table ****************************************
  Set TD = New TableDef    'create a new table
  TD.Name = "NozzlesTypMfg" 'name the new table
  
  Set FD = New Field        'create a new field
  FD.Name = "Nozzle"        'name the field
  FD.Type = dbText         'type the field
  FD.Size = 20              'size the field
  TD.Fields.Append FD       'append the field to the table
  
  Set FD = New Field        'create a new field
  FD.Name = "Type"          'name the field
  FD.Type = dbText         'type the field
  FD.Size = 24              'size the field
  TD.Fields.Append FD       'append the field to the table
  
  Set FD = New Field        'create a new field
  FD.Name = "Manufacturer"  'name the field
  FD.Type = dbText         'type the field
  FD.Size = 20              'size the field
  TD.Fields.Append FD       'append the field to the table
  
  DB.TableDefs.Append TD    'append the table to the database

  'close the database
  DB.Close
End Sub

Sub AddToLog(s As String)
'Add a string to the Log control
  Dim fsave As Control
  frmMain.lstLog.AddItem Time$ + " " + s
  frmMain.lstLog.Refresh
  Set fsave = frmMain.ActiveControl 'save current control
  frmMain.lstLog.SetFocus           'set focus to list box
  SendKeys "{END}"                  'send an END key to the list box
  DoEvents
  frmMain.lstLog.Selected(frmMain.lstLog.ListIndex) = False
  fsave.SetFocus                    'restore original focus
End Sub

Sub AddTrialRecords()
'read the Field Trial Data source file and stuff the data into tables
  
  Dim DB As Database
  Dim DS As Dynaset
  Dim fn As String

  'local variables for record data storage
  Dim Title As String * 32
  Dim SprayMaterial As Integer
  Dim TestNumber As Integer
  Dim aircraft As Integer
  Dim F141 As Single
  Dim Height As Single
  Dim WindSpeed As Single
  Dim D10 As Single
  Dim D50 As Single
  Dim D90 As Single
  Dim swathdisp As Single
  Dim AppEff As Single
  Dim DwndDrift As Single
  Dim Airborne As Single
  Dim EvapFrac As Single
  Dim DSDNumDrop As Integer
  ReDim DSDDiam(1) As Single
  ReDim DSDMassFrac(1) As Single
  Dim ACName As String * 32
  Dim ACWingType As Integer
  Dim ACSemiSpan As Single
  Dim ACTypSpeed As Single
  Dim ACWeight As Single
  Dim ACBiplSep As Single
  Dim ACPlanArea As Single
  Dim ACPropRPM As Single
  Dim ACPropRad As Single
  Dim ACEngVert As Single
  Dim ACEngFwd As Single
  Dim NZNumNoz As Integer
  ReDim NZPosHoriz(1) As Single
  Dim NZPosVert As Single
  Dim NZPosFwd As Single
  Dim SMSpecGrav As Single
  Dim METWS As Single
  Dim METCanopyHeight As Single
  Dim METWD As Single
  Dim METtemp As Single
  Dim METHumidity As Single
  Dim SMEvapRate As Single
  Dim SMFlowRate As Single
  Dim SMNVrate As Single
  Dim SMACrate As Single
  Dim CTLNumLines As Integer
  Dim CTLSwathWidth As Single
  Dim CTLHeight As Single
  Dim CALCNumPredDep As Integer
  ReDim CALCPredDepDist(1) As Single
  ReDim CALCPredDepVal(1) As Single
  Dim CALCNumMeasDep As Integer
  ReDim CALCMeasDepDist(1) As Single
  ReDim CALCMeasDepVal1(1) As Single
  ReDim CALCMeasDepVal2(1) As Single
  ReDim CALCMeasDepVal3(1) As Single
  ReDim CALCMeasDepVal4(1) As Single
  
  Dim str As String

  'Open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  
' Field Trial Data ***************************************
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset("FieldTrial")

  'open the source file
  fn = GD.SrcPath & GD.Src.fieldrun
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Line Input #1, Title
    Input #1, DSDNumDrop
    If DSDNumDrop > 0 Then
      ReDim DSDDiam(DSDNumDrop - 1) As Single
      ReDim DSDMassFrac(DSDNumDrop - 1) As Single
      For i = 0 To DSDNumDrop - 1
        Input #1, DSDDiam(i), DSDMassFrac(i)
      Next
    End If
    Line Input #1, ACName
    Input #1, ACWingType
    Input #1, ACSemiSpan
    Input #1, ACTypSpeed
    Input #1, ACWeight
    Input #1, ACBiplSep
    Input #1, ACPlanArea
    Input #1, ACPropRPM
    Input #1, ACPropRad
    Input #1, ACEngVert
    Input #1, ACEngFwd
    Input #1, NZNumNoz
    If NZNumNoz > 0 Then
      ReDim NZPosHoriz(NZNumNoz - 1) As Single
      For i = 0 To NZNumNoz - 1
        Input #1, NZPosHoriz(i)
      Next
    End If
    Input #1, NZPosVert
    Input #1, NZPosFwd
    Input #1, SMSpecGrav
    Input #1, METWS
    Input #1, METCanopyHeight
    Input #1, METWD
    Input #1, METtemp
    Input #1, METHumidity
    Input #1, SMEvapRate
    Input #1, SMFlowRate
    Input #1, SMNVrate
    Input #1, SMACrate
    Input #1, CTLNumLines
    Input #1, CTLSwathWidth
    Input #1, CTLHeight
    Input #1, CALCNumPredDep
    If CALCNumPredDep > 0 Then
      ReDim CALCPredDepDist(CALCNumPredDep - 1) As Single
      ReDim CALCPredDepVal(CALCNumPredDep - 1) As Single
      For i = 0 To CALCNumPredDep - 1
        Input #1, CALCPredDepDist(i), CALCPredDepVal(i)
      Next
    End If
    Input #1, CALCNumMeasDep
    If CALCNumMeasDep > 0 Then
      ReDim CALCMeasDepDist(CALCNumMeasDep - 1) As Single
      ReDim CALCMeasDepVal1(CALCNumMeasDep - 1) As Single
      ReDim CALCMeasDepVal2(CALCNumMeasDep - 1) As Single
      ReDim CALCMeasDepVal3(CALCNumMeasDep - 1) As Single
      ReDim CALCMeasDepVal4(CALCNumMeasDep - 1) As Single
      For i = 0 To CALCNumMeasDep - 1
        Input #1, CALCMeasDepDist(i), CALCMeasDepVal1(i), CALCMeasDepVal2(i), CALCMeasDepVal3(i), CALCMeasDepVal4(i)
      Next
    End If
    Input #1, D10, D50, D90, swathdisp, AppEff, DwndDrift, Airborne, EvapFrac
    Input #1, SprayMaterial, TestNumber, aircraft, F141, Height, WindSpeed

    'Add the new record to the dropsize table
    DS.AddNew
    DS.Fields("Title") = Title
    DS.Fields("SprayMaterial") = SprayMaterial
    DS.Fields("TestNumber") = TestNumber
    DS.Fields("Aircraft") = aircraft
    DS.Fields("F141") = F141
    DS.Fields("Height") = Height
    DS.Fields("WindSpeed") = WindSpeed
    DS.Fields("D10") = D10
    DS.Fields("D50") = D50
    DS.Fields("D90") = D90
    DS.Fields("SwathDisp") = swathdisp
    DS.Fields("AppEff") = AppEff
    DS.Fields("DwndDrift") = DwndDrift
    DS.Fields("Airborne") = Airborne
    DS.Fields("EvapFrac") = EvapFrac
    
    DS.Fields("DSDNumDrop") = DSDNumDrop
    ArrayToField DS.Fields("DSDDiam"), DSDDiam(), DSDNumDrop
    ArrayToField DS.Fields("DSDMassFrac"), DSDMassFrac(), DSDNumDrop
    DS.Fields("ACName") = ACName
    DS.Fields("ACWingType") = ACWingType
    DS.Fields("ACSemiSpan") = ACSemiSpan
    DS.Fields("ACTypSpeed") = ACTypSpeed
    DS.Fields("ACWeight") = ACWeight
    DS.Fields("ACBiplSep") = ACBiplSep
    DS.Fields("ACPlanArea") = ACPlanArea
    DS.Fields("ACPropRPM") = ACPropRPM
    DS.Fields("ACPropRad") = ACPropRad
    DS.Fields("ACEngVert") = ACEngVert
    DS.Fields("ACEngFwd") = ACEngFwd
    DS.Fields("NZNumNoz") = NZNumNoz
    ArrayToField DS.Fields("NZPosHoriz"), NZPosHoriz(), NZNumNoz
    DS.Fields("NZPosVert") = NZPosVert
    DS.Fields("NZPosFwd") = NZPosFwd
    DS.Fields("SMSpecGrav") = SMSpecGrav
    DS.Fields("METWS") = METWS
    DS.Fields("METCanopyHeight") = METCanopyHeight
    DS.Fields("METWD") = METWD
    DS.Fields("METtemp") = METtemp
    DS.Fields("METHumidity") = METHumidity
    DS.Fields("SMEvapRate") = SMEvapRate
    DS.Fields("SMFlowRate") = SMFlowRate
    DS.Fields("SMNVrate") = SMNVrate
    DS.Fields("SMACrate") = SMACrate
    DS.Fields("CTLNumLines") = CTLNumLines
    DS.Fields("CTLSwathWidth") = CTLSwathWidth
    DS.Fields("CTLHeight") = CTLHeight
    DS.Fields("CALCNumPredDep") = CALCNumPredDep
    ArrayToField DS.Fields("CALCPredDepDist"), CALCPredDepDist(), CALCNumPredDep
    ArrayToField DS.Fields("CALCPredDepVal"), CALCPredDepVal(), CALCNumPredDep
    DS.Fields("CALCNumMeasDep") = CALCNumMeasDep
    ArrayToField DS.Fields("CALCMeasDepDist"), CALCMeasDepDist(), CALCNumMeasDep
    ArrayToField DS.Fields("CALCMeasDepVal1"), CALCMeasDepVal1(), CALCNumMeasDep
    ArrayToField DS.Fields("CALCMeasDepVal2"), CALCMeasDepVal2(), CALCNumMeasDep
    ArrayToField DS.Fields("CALCMeasDepVal3"), CALCMeasDepVal3(), CALCNumMeasDep
    ArrayToField DS.Fields("CALCMeasDepVal4"), CALCMeasDepVal4(), CALCNumMeasDep
    DS.Update

  Wend
  Close #1
  AddToLog Format$(numrecs) + " Field Trial records added"
  
  'close the dynaset
  DS.Close
  
  'close the database
  DB.Close
End Sub

Sub AddTrialTables()
'Add tables for Basic Data to database
'
'*********************************************************
' Database:
'   Table: FieldTrial
'     Field: Name
'
  Dim DB As Database
  Dim TD As TableDef
  Dim FD As Field

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Basic DSD table ****************************************
  Set TD = New TableDef      'create a new table
  TD.Name = "FieldTrial"     'name the new table
  
  Set FD = New Field
  FD.Name = "Title"
  FD.Type = dbText
  FD.Size = 32
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "SprayMaterial"
  FD.Type = dbInteger
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "TestNumber"
  FD.Type = dbInteger
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "Aircraft"
  FD.Type = dbInteger
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "F141"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "Height"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "WindSpeed"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "D10"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "D50"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "D90"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "SwathDisp"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "AppEff"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "DwndDrift"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "Airborne"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "EvapFrac"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "DSDNumDrop"
  FD.Type = dbInteger
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "DSDDiam"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "DSDMassFrac"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "ACName"
  FD.Type = dbText
  FD.Size = 40
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "ACWingType"
  FD.Type = dbInteger
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "ACSemiSpan"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "ACTypSpeed"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "ACWeight"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "ACBiplSep"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "ACPlanArea"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "ACPropRPM"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "ACPropRad"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "ACEngVert"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "ACEngFwd"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "NZNumNoz"
  FD.Type = dbInteger
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "NZPosHoriz"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "NZPosVert"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "NZPosFwd"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "METWS"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "METCanopyHeight"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "METWD"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "METtemp"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "METHumidity"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "SMEvapRate"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "SMFlowRate"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "SMSpecGrav"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "SMNVrate"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "SMACrate"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "CTLNumLines"
  FD.Type = dbInteger
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "CTLSwathWidth"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "CTLHeight"
  FD.Type = dbSingle
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "CALCNumPredDep"
  FD.Type = dbInteger
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "CALCPredDepDist"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "CALCPredDepVal"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "CALCNumMeasDep"
  FD.Type = dbInteger
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "CALCMeasDepDist"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "CALCMeasDepVal1"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "CALCMeasDepVal2"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "CALCMeasDepVal3"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "CALCMeasDepVal4"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  DB.TableDefs.Append TD     'append the table to the database

  'close the database
  DB.Close
End Sub

Sub CenterForm(f As Form)
'Center the form on the screen
  f.Left = (Screen.Width / 2) - (f.Width / 2)
  f.Top = (Screen.Height / 2) - (f.Height / 2)
End Sub

Sub ConditionalAddItem(s As String, n As Integer, l() As String)
'add a string to a list of strings, only if it is not
'already there
  Dim addit As Integer
  'look for matching items
  addit = True
  For i = 0 To n - 1
    If (s = l(i)) Then
      addit = False
    Exit For
  End If
  Next
  'add the new item
  If (addit) Then
    l(n) = s
    n = n + 1
  End If
End Sub

Sub CreateAgdriftDatabase()
'Create a new agdrift database from scratch
'
  Dim DB As Database

  'open the database
  Set DB = CreateDatabase(GD.DBDirPath & GD.DBFileName, dbLangGeneral)

  'close the completed database
  DB.Close
End Sub

Sub DeleteAircraftTables(TableName As String)
'Delete tables for Aircraft from database
'
'*********************************************************
'   Table: Aircraft
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  DB.TableDefs.Delete TableName
  DB.Close
End Sub

Sub DeleteBasicTables()
'Delete tables for Basic Data from database
'
'*********************************************************
'   Table: Aircraft
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  DB.TableDefs.Delete "BasicDSD"
  DB.TableDefs.Delete "BasicAC"
  DB.TableDefs.Delete "BasicDep"
  DB.TableDefs.Delete "BasicSgl"
  DB.TableDefs.Delete "BasicNZ"
  DB.TableDefs.Delete "BasicInfo"
  DB.Close
End Sub

Sub DeleteComponentsTables()
'Delete tables for Components from database
'
'*********************************************************
'   Table: Components      list of components for each substance
'   Table: ComponentList   list of all components
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  DB.TableDefs.Delete "Components"
  DB.TableDefs.Delete "ComponentList"
  
  'close the database
  DB.Close
End Sub

Sub DeleteARSTables()
'Delete tables for ARS Nozzles from database
'
'*********************************************************
'   Table: ARSNozzle      list of ARS Nozzles
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  DB.TableDefs.Delete "ARSNozzle"
  
  'close the database
  DB.Close
End Sub

Sub DeleteDropsizeTables(FSflag As Integer)
'Delete tables for Dropsize from database
'
'*********************************************************
'   Table: Dropsize        all the data with 32 mass fracs
'   Table: Dropdiam        diameter data for above
'   Table: SubstanceList   unique list of Substances
'   Table: NozzleList      unique list of Nozzles
'   Table: NozzleAngleList unique list of nozzle angles
'   Table: WindSpeedList   unique list of wind speeds
'
'
  Dim DB As Database
  Dim FSstr As String

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  If FSflag = 0 Then
    FSstr = ""
  Else
    FSstr = "FS"
  End If
  DB.TableDefs.Delete "Dropsize" + FSstr
  DB.TableDefs.Delete "Dropdiam" + FSstr
  DB.TableDefs.Delete "SubstanceList" + FSstr
  DB.TableDefs.Delete "NozzleList" + FSstr
  DB.TableDefs.Delete "NozzleAngleList" + FSstr
  DB.TableDefs.Delete "WindSpeedList" + FSstr
  
  'close the database
  DB.Close
End Sub

Sub DeleteEvaporationTables()
'Delete tables for Evaporation from database
'
'*********************************************************
'   Table: Evaporation
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  DB.TableDefs.Delete "Evaporation"
  
  'close the database
  DB.Close
End Sub

Sub DeleteMaterialTables()
'Delete tables for Materials from database
'
'*********************************************************
'   Table: Material
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  DB.TableDefs.Delete "Materials"
  
  'close the database
  DB.Close
End Sub

Sub DeleteNozzleTables()
'Delete tables for Nozzles from database
'
'*********************************************************
'   Table: Nozzles
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  DB.TableDefs.Delete "Nozzles"
  
  'close the database
  DB.Close
End Sub

Sub DeleteNozzleTMTables()
'Delete tables for Nozzles from database
'
'*********************************************************
'   Table: NozzlesTM
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  DB.TableDefs.Delete "NozzlesTM"
  
  'close the database
  DB.Close
End Sub

Sub DeleteInfoTables()
'Delete tables for Info from database
'
'*********************************************************
'   Table: Info
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  DB.TableDefs.Delete "Info"
  
  'close the database
  DB.Close
End Sub

Sub DeleteTrialTables()
'Delete tables for Basic Data from database
'
'*********************************************************
'   Table: FieldTrial
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  DB.TableDefs.Delete "FieldTrial"
  DB.Close
End Sub

Function GetCurrRec(DS As Dynaset) As String

  Dim i As Integer
  Static FieldStr As String
  Static RecStr As String
  RecStr = ""

  'Step through each field in the current record and accumulate

  'the contents of each field into a string
  For i = 0 To DS.Fields.Count - 1

    'Pad out to the right size
    FieldStr = Space(DS.Fields(i).Size)

    Select Case DS.Fields(i).Type

      'Copy the binary representation of the field to a string (FieldStr)

      Case 1, 2       'Bytes
        CopyMemory ByVal FieldStr, CInt(DS.Fields(i).Value), DS.Fields(i).Size

      Case 3          'Integers
        CopyMemory ByVal FieldStr, CInt(DS.Fields(i).Value), DS.Fields(i).Size

      Case 4          'Long integers
        CopyMemory ByVal FieldStr, CLng(DS.Fields(i).Value), DS.Fields(i).Size

      Case 5          'Currency
        CopyMemory ByVal FieldStr, CCur(DS.Fields(i).Value), DS.Fields(i).Size

      Case 6          'Singles
        CopyMemory ByVal FieldStr, CSng(DS.Fields(i).Value), DS.Fields(i).Size

      Case 7, 8       'Doubles
        CopyMemory ByVal FieldStr, CDbl(DS.Fields(i).Value), DS.Fields(i).Size

      Case 9, 10      'String types
        CopyMemory ByVal FieldStr, ByVal CStr(DS.Fields(i).Value), Len(DS.Fields(i).Value)

      Case 11, 12     'Memo and long binary
        FieldStr = DS.Fields(i).GetChunk(0, DS.Fields(i).FieldSize())

    End Select

    'Accumulate the field string into a record string
    RecStr = RecStr & FieldStr

  Next

  'Return the accumulated string containing the contents of all
  'fields in the current record
  GetCurrRec = RecStr

End Function

Sub ArrayToField(fld As Field, X() As Single, NumX As Integer)
'Store an array in a Longbinary database field
'
' fld must be a longbinary field of an open recordset
  Dim bb() As Byte
  Dim nb As Long
  
  nb = CLng(NumX) * Len(X(0))       'number of bytes in array
  ReDim bb(nb)                'make room for data
  CopyMemory bb(0), X(0), nb  'transfer data to byte array
  fld.AppendChunk bb()        'store in DB field
End Sub

Sub FieldToArray(fld As Field, X() As Single, NumX As Integer)
'Recover an array from a Longbinary database field
'
' fld must be a longbinary field of an open recordset
  Dim bb() As Byte
  
  bb = fld.GetChunk(0, fld.FieldSize) 'retrieve raw data
  CopyMemory X(0), bb(0), UBound(bb)  'transfer it to the array
  NumX = (UBound(bb) + 1) / Len(X(0)) 'elements = bytes / bytes/element
End Sub

Public Sub AddWindroseRecords()
'read the Wind Rose source files and stuff the data into tables
'
' Database:
'   Table: WindRose
'     Field: Name
'     Field: SamsonID
'     Field: Latitude
'     Field: Longitude
'     Field: Elevation
'     Field: T55075
'     Field: RH255075
'     Field: WS255075
'     Field: WD
'     Field: Temperature
'     Field: Humidity
'     Field: MaxSpeed
'     Field: Frequency
  
  Dim DB As Database
  Dim DS As Dynaset
  Dim fn As String
  
  Dim numrecs As Integer
  Dim numspeeds As Integer
  Dim speed As Integer
  Dim RecStr As String

  'local variables for record data storage
  Dim Location As String
  Dim SamsonID As String
  Dim Lat As String
  Dim Lon As String
  Dim Elev As Long
  Dim T255075(12 * 3 - 1) As Single '(12, 3)
  Dim RH255075(12 * 3 - 1) As Single '(12, 3)
  Dim WS255075(12 * 3 - 1) As Single '(12, 3)
  Dim WD(11) As Single
  Dim Temperature(11) As Single
  Dim Humidity(11) As Single
  Dim MaxSpeed As Integer
  Dim Frequency(35, 11, 19) As Integer '(36, 12, 20)
  Dim bb() As Byte
  Dim nb As Long
  
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset("WindRose")

  'Open the index file
  numrecs = 0
  Open GD.SrcPath & "Windrose\Ordering.lst" For Input As #2
  While Not EOF(2)
    numrecs = numrecs + 1
    Line Input #2, RecStr
    Location = Trim$(Mid$(RecStr, 1, 25))
    SamsonID = Trim$(Mid$(RecStr, 27, 7))
    Lat = Trim$(Mid$(RecStr, 37, 6))
    Lon = Trim$(Mid$(RecStr, 46, 7))
    Elev = Val(Mid$(RecStr, 54, 6))
    
    'open the Samson file
    fn = GD.SrcPath & "Windrose\" + SamsonID + ".out"
    AddToLog "Reading from " + SamsonID + ".out (" & Location & ")"
    Open fn For Input As #1
  
    'init the Frequency array
    For k = 0 To 19
      For j = 0 To 11
        For i = 0 To 35
          Frequency(i, j, k) = 0
        Next
      Next
    Next
    
    'read an input record
    Input #1, dummy 'len of location string
    Line Input #1, dummy 'location and SamsonID
    'set 0 - Temp, RH, WS, and WD percentiles
    Input #1, dummy  'a zero
    n = 0
    For j = 0 To 2 '25, 50, and 75 percentiles
      For i = 0 To 11  'Jan to Dec
        Input #1, T255075(n)
        n = n + 1
      Next
    Next
    n = 0
    For j = 0 To 2 '25, 50, and 75 percentiles
      For i = 0 To 11  'Jan to Dec
        Input #1, RH255075(n)
        n = n + 1
      Next
    Next
    n = 0
    For j = 0 To 2 '25, 50, and 75 percentiles
      For i = 0 To 11  'Jan to Dec
        Input #1, WS255075(n)
        n = n + 1
      Next
    Next
    For i = 0 To 11  'Jan to Dec
      Input #1, WD(i)
    Next
    'Set 1 - average Temp and RH
    Input #1, dummy  'a one
    For i = 0 To 11
      Input #1, Temperature(i)
    Next
    For i = 0 To 11
      Input #1, Humidity(i)
    Next
    'Sets 2 and up - Wind Speed Frequency Tables
    'The "set number" is the wind speed, and the table below it
    'is the frequency chart. This repeats to the end of the file.
    MaxSpeed = 0
    While Not EOF(1)
      Input #1, speed
      If speed > MaxSpeed Then MaxSpeed = speed
      For i = 0 To 35
        For j = 0 To 11
          Input #1, Frequency(i, j, speed - 1)
        Next
      Next
    Wend
    Close #1

    'Add the new record to the dropsize table
    DS.AddNew
    DS.Fields("Name") = Location
    DS.Fields("SamsonID") = SamsonID
    DS.Fields("Latitude") = Lat
    DS.Fields("Longitude") = Lon
    DS.Fields("Elevation") = Elev
    ArrayToField DS.Fields("T255075"), T255075(), 3 * 12
    ArrayToField DS.Fields("RH255075"), RH255075(), 3 * 12
    ArrayToField DS.Fields("WS255075"), WS255075(), 3 * 12
    ArrayToField DS.Fields("WD"), WD(), 12
    ArrayToField DS.Fields("Temperature"), Temperature(), 12
    ArrayToField DS.Fields("Humidity"), Humidity(), 12
    DS.Fields("MaxSpeed") = MaxSpeed
    'this is a special case.
    'Frequency is a multidimensional Integer array and ArrayToField normally
    'works with 1D Single arrays. So we do the work of ArrayToField here.
    nb = 36 * 12 * 20 * Len(Frequency(0, 0, 0))  'number of bytes in array
    ReDim bb(nb)                                 'make room for data
    CopyMemory bb(0), Frequency(0, 0, 0), nb     'transfer data to byte array
    DS.Fields("Frequency").AppendChunk bb()      'store in DB field
    
    DS.Update

  Wend
  Close #2
  
  'close the dynaset
  DS.Close
  
  AddToLog Format$(numrecs) + " wind rose records added"

  DB.Close
End Sub

Public Sub AddWindroseTables()
'Add tables for Wind Rose Data to database
'
'*********************************************************
' Database:
'   Table: WindRose
'     Field: Name
'     Field: SamsonID
'     Field: Latitude
'     Field: Longitude
'     Field: Elevation
'     Field: T55075
'     Field: RH255075
'     Field: WS255075
'     Field: WD
'     Field: Temperature
'     Field: Humidity
'     Field: MaxSpeed
'     Field: Frequency
'
  Dim DB As Database
  Dim TD As TableDef
  Dim FD As Field

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Basic DSD table ****************************************
  Set TD = New TableDef      'create a new table
  TD.Name = "WindRose"       'name the new table
  
  Set FD = New Field
  FD.Name = "Name"
  FD.Type = dbText
  FD.Size = 32
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "SamsonID"
  FD.Type = dbText
  FD.Size = 7
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "Latitude"
  FD.Type = dbText
  FD.Size = 12
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "Longitude"
  FD.Type = dbText
  FD.Size = 13
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "Elevation"
  FD.Type = dbLong
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "T255075"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "RH255075"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "WS255075"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "WD"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "Temperature"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "Humidity"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "MaxSpeed"
  FD.Type = dbInteger
  TD.Fields.Append FD
  
  Set FD = New Field
  FD.Name = "Frequency"
  FD.Type = dbLongBinary
  TD.Fields.Append FD
  
  DB.TableDefs.Append TD     'append the table to the database

  'close the database
  DB.Close
End Sub

Public Sub DeleteWindroseTables()
'Delete tables for Wind Rose Data from database
'
'*********************************************************
'   Table: WindRose
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  DB.TableDefs.Delete "WindRose"
  DB.Close
End Sub

Public Sub AddCanopyTables(TableName As String)
'Add tables for Canopy info to database
'
'*********************************************************
' Database:
'   Table: Canopy
'     Field: Name
'     Field: LAI
'     Field: Height
'     Field: B
'     Field: C
'
  Dim DB As Database
  Dim TD As TableDef
  Dim FD As Field

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)

  'create Dropsize table ****************************************
  Set TD = New TableDef
  TD.Name = TableName      'name the new table
  
  Set FD = New Field
  FD.Name = "Name"        'name the field
  FD.Type = dbText       'type the field
  FD.Size = 40            'size the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "LAI"         'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "Height"      'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "B"           'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  Set FD = New Field
  FD.Name = "C"           'name the field
  FD.Type = dbSingle     'type the field
  TD.Fields.Append FD     'append the field to the table
  
  DB.TableDefs.Append TD  'append the table to the database

  'close the database
  DB.Close
End Sub

Public Sub AddCanopyRecords(TableName As String, SrcFile As String)
'read the Canopy source file and stuff the data into tables
  
  Dim fn As String
  Dim DB As Database

  'local variables for record data storage
  Dim FD1 As String   'name
  Dim FD2 As Single
  Dim FD3 As Single   'semi-span (m)
  Dim FD4 As Single   'semi-span (m)
  Dim FD5 As Single   'semi-span (m)

  Dim DS As Dynaset
  
  'Open the Database (Library) file
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  
  'create dynaset of records to work on
  Set DS = DB.CreateDynaset(TableName)

  'open the source file
  fn = GD.SrcPath & SrcFile
  AddToLog "Reading from " & fn
  Open fn For Input As #1
  
  numrecs = 0
  While Not EOF(1)
    'read an input record
    numrecs = numrecs + 1
    Input #1, FD1, FD2, FD3, FD4, FD5

    'Add the new record to the dropsize table
    DS.AddNew
    DS.Fields("Name") = FD1
    DS.Fields("LAI") = FD2
    DS.Fields("Height") = FD3
    DS.Fields("B") = FD4
    DS.Fields("C") = FD5

    DS.Update
  Wend
  Close #1
  
  'close the dynaset
  DS.Close
  
  AddToLog Format$(numrecs) + " canopy records added"

  DB.Close
End Sub

Public Sub DeleteCanopyTables()
'Delete tables for Canopy from database
'
'*********************************************************
'   Table: Canopy
'
'
  Dim DB As Database

  'open the database
  Set DB = OpenDatabase(GD.DBDirPath & GD.DBFileName)
  DB.TableDefs.Delete "Canopy"
  DB.Close
End Sub
