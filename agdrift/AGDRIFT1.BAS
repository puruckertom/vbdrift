Attribute VB_Name = "basAGDRIFT1"
' $Id: agdrift1.bas,v 1.28 2011/12/27 17:47:00 tom Exp $
'AGDRIFT1.BAS - main code module for AgDRIFT
'
'Current AgDRIFT version
'The AgDRIFT version notation picked up the form "major.middle.minor'
'at some point along the way. Up to version 2.0.10, this was easy to
'handle because the middle number was zero. The number "2.05", for
'example, could refer to version 2.0.5, and 2.1 could refer to 2.0.10.
'Starting with 2.1.1, a new notation needed to be adopted. So now,
'we pack the major and middle numbers to the left of the decimal point,
'and the minor number to the right. So, 21.01 now refers to
'version 2.1.1
Public Const AGDRIFTVERSION = 21.01    'Current Agdrift Verion: 2.1.01

'Agricultural audience
'True:  Agricultural model = "Regulatory"
'False: Agricultural model = "Public Use"
'This flag allows two versions of AgDRIFT to be
'compiled. The first, "Regulatory", is for members of the
'SDTF and their designates, and includes Tier II stuff.
'The other "Public Use" is intended for downloading from
'a web site for the general public. This flag has no effect
'on the Forestry side.
Public Const AGDRIFTREGULATORY = True
'
'**************** User-defined data types ****************
'
'======= User Input data area ============================
'
' The structures defined below contain all of the input data
' that is specified by the user. These structures are duplicated
' for the fortran part of this program in a file called
' "agdstruc.inc" in the dll directory.
'
' IF CHANGES ARE MADE TO THE USER DATA STRUCTURES HERE, MAKE
' COMPLIMENTARY CHANGES TO THE FORTRAN (AND VICE VERSA).
'
'define constants for sizing arrays
Public Const MAX_DSD = 3      'max number of drop distributions
Public Const MAX_DROPS = 100  'max number of drop categories
Public Const MAX_NOZZLES = 60 'max number of nozzles
Public Const MAX_ENGINES = 2  'max number of aircraft engines
Public Const MAX_ENVELOPE = 25 'Max canopy envelope levels
Public Const MAX_LAI = 10     'max canopy LAI levels
Public Const MAX_CALCDATA = 4900 'max size of calc arrays
Public Const MAX_ASAE = 6     'Max ASAE drop distributions
'tbc add these sometime
'public const MAX_BASICAC
'etc
'
'Ground Application Data
Public Type GroundApplData
  BasicType As Integer     '0-2=low boom m,c,vc 3-5=high boom m,c,vc
  NumSwaths As Integer     'Number of Swaths for ground sprayer
End Type

'Orchard Airblast Data
Public Type OrchardAirblastData
  BasicType As Integer     '0-2=combination 3-13=individual
  BegTrow As Integer       'Beginning Tree Row
  EndTrow As Integer       'End Tree Row
End Type

'Drop Size Distribution Data
Public Type DropSizeDistData
  Type As Integer          '0=basic (ASAE) 1=dropkick 2=user-def 3,4=lib 5=ARS
  BasicType As Integer     'Basic DSD category: 0-17
  Name As String * 40      'Name for DSD
  LName As Integer         'Length of Name
  NumDrop As Integer       'Number of Drop Categories
  Diam(MAX_DROPS - 1) As Single     'Drop Diameters (um)
  MassFrac(MAX_DROPS - 1) As Single 'Drop Volume Fractions
  LibSelFlag As Integer    'AGDSL flag: 1=call AGDSL w/key 3 2=key 4
End Type

'USDA ARS Nozzle Models data
Public Type DropKirkData
  MaxErrorLevel As Integer 'Max error level encountered during run 0=none 1=warn 2=err
  NozType As Integer       'Nozzle Type: 0-3
  NameNoz As String * 40   'Nozzle Name
  LNameNoz As Integer      'Length of Name
  Orifice As Single        'Orifice Size (in) or Orifice Number
  Speed As Single          'Tunnel Speed (m/s)
  NozAngle As Single       'Nozzle angle (deg) or Restrictor Number
  Pressure As Single       'Nozzle pressure (bar)
  SprayType As Integer     'Model Output Style: 0=Spray Quality 1=DSD
  SpectrumSource As Integer 'Drop Spectrum Data Source: 0=Malvern 1=PMS
End Type

'USDA FS Rotary Atomizer Data
Public Type HKData
  MaxErrorLevel As Integer 'Max error level encountered during run 0=none 1=warn 2=err
  MatType As Integer       'Material Type 0=Water 1=Water/1% StaPut 2=Water/0.25% Hasten
  RotType As Integer       'Atomizer Type: 0=AU4000 1=AU5000
  Speed As Single          'Tunnel Speed (m/s)
  BladeAngle As Single     'Blade angle (deg)
  BladeRPM As Single       'Rotation Rate (rpm)
  FlowRate As Single       'Nozzle flowrate (L/min)
  SprayType As Integer     'Model Output Style: 0=Spray Quality 1=DSD
End Type

'DropKick data
Public Type DropKickData
  MaxErrorLevel As Integer 'Max error level encountered during run 0=none 1=warn 2=err
  NozType As Integer       'Nozzle Type: 0=user 1=lib
  NameNoz As String * 40   'Nozzle Name
  LNameNoz As Integer      'Length of Name
  VMD As Single            'Volume Median Diameter (um)
  RelSpan As Single        'Relative Span
  EffDiam As Single        'Effective Diameter (cm)
  SprayAngle As Single     'Spray Angle (deg)
  MatType As Integer       'Material Type: 0=user 1=lib
  NameMat As String * 40   'Material Name
  LNameMat As Integer      'Length of Name
  DynSurfTens As Single    'Dynamic Surface Tension (dynes/cm)
  ShearVisc As Single      'Shear Viscosity (cp)
  Density As Single        'Density (gm/cm3)
  ElongVisc As Single      'Elongational Viscosity (cp)
  Speed As Single          'Tunnel Speed (m/s)
  NozAngle As Single       'Nozzle angle (deg)
  Pressure As Single       'Nozzle pressure (bar)
  FlowType As Integer      '0=Flow obtained from "scaled" 1=input
  flow As Single           'Flow rate per nozzle (l/min)
  SprayType As Integer     'Model Output Style: 0=Spray Quality 1=DSD
End Type

'Spray Material Data
Public Type SprayMaterialData
  Type As Integer          'Type: 0=basic 1=user-def 2=lib
  BasicType As Integer     'Basic Type: 0=oil 1=water
  Name As String * 40      'Spray Mat Library Substance Name
  LName As Integer         'Length of Name
  CalcInputSelect As Integer 'Calulation input selection 0=rates 1=tank mix
  NVFrac As Single         'Nonvolatile Fraction
  ACFrac As Single         'Active fraction
  ActSolFrac As Single     'Fraction of tank mix that is Active solution
  AddSolFrac As Single     'Fraction of tank mix that is Additive Solution
  ActNVFrac As Single      'Fraction of Active solution that is nonvolatile
  AddNVFrac As Single      'Fraction of Additive solution that is nonvolatile
  FlowRate As Single       'Flow Rate (L/ha or L/min)
  FlowRateUnits As Integer 'Flow Rate units Flag (0=L/ha 1=L/min)
  SpecGrav As Single       'Specific Gravity of tank mix
  NonVGrav As Single       'Specific Gravity of nonvolatiles
  EvapRate As Single       'Evaporation Rate ()
End Type

'Aircraft Data
Public Type AircraftData
  Type As Integer          '0=Basic 1=user-def 2=lib
  BasicType As Integer     '0=AgHusky 1=AT502 2=Wasp 3=AT401
  Name As String * 40      'Aircraft Name
  LName As Integer         'Length of Name
  WingType As Integer      'Aircraft Wing Type: 3=fixed 4=helicopter
  SemiSpan As Single       'Semispan or Rotor Radius (m)
  TypSpeed As Single       'Typical Spraying Speed (m/s)
  BiplSep As Single        'Biplane Distance Between Wings (m)
  Weight As Single         'Weight (kg)
  PlanArea As Single       'Planform Area (m2)
  PropRPM As Single        'Propeller/Rotor RPM
  PropRad As Single        'Propeller Radius (m)
  PropEff As Single        'Propeller Efficiency
  EngVert As Single        'Engine Vertical position (m)
  EngFwd As Single         'Engine Forward position (m)
  NumEng As Integer        'Number of engines
  EngHoriz(MAX_ENGINES - 1) As Single  'Engine Horizontal positions (m)
  WingVert As Single       'Dist from wingtip vortex to trailing edge (m)
  BoomVert As Single       'Vert dist from boom to trailing edge (m)
  BoomFwd As Single        'Fwd dist from boom to trailing edge (m)
  DragCoeff As Single      'Drag Coefficient
End Type

'Nozzle data
Public Type NozzleData
  Type As Integer          'Distribution type: 0=basic 1=user-def
  BasicType As Integer     'Basic Type: (same as Basic Aircraft)
  Name As String * 40      'Distribution Name
  LName As Integer         'Length of Name
  NumNoz As Integer        'Number of Nozzles
  NozType(MAX_NOZZLES - 1) As Integer 'Nozzle (DSD) Type
  PosHoriz(MAX_NOZZLES - 1) As Single 'Horizontal Positions (m)
  PosVert(MAX_NOZZLES - 1) As Single  'Vertical Position (m)
  PosFwd(MAX_NOZZLES - 1) As Single   'Forward (Axial) Position (m)
  PosHorizLimit As Single  'Wingspan % for user-def Dist Limit
  BoomWidth As Single      'Dist Limit % for basic dists.
End Type
  
'Meteorological Data
Public Type MetData
  WS As Single             'Wind Speed (m/s)
  WD As Single             'Wind Direction (deg)
  WindHeight As Single     'Height of WS (m)
  temp As Single           'Temperature (deg C)
  Humidity As Single       'Humidity (%)
  Pressure As Single       'Barometric Pressure (mb)
  VortexDecay As Single    'Vortex Decay Rate (m/s)
  SurfRough As Single      'Surface Roughness (m)
  Insolation As Integer    'Insolation index (0-7: Strong, Mod, Slight, etc.)
End Type

'Canopy Data
Public Type CanopyData
  Type As Integer          'Canopy Type 0=none 1=story 2=optical 3=basic
  Name As String * 40      'Canopy Name
  LName As Integer         'Length of Name
  EleSiz As Single         'Element Size (m)
  StanDen As Single        'Stand Density (stems/ha)
  NumEnv As Integer        'Number of tree envelope levels
  EnvHgt(MAX_ENVELOPE - 1) As Single 'Height (m)
  EnvDiam(MAX_ENVELOPE - 1) As Single 'Diameter (m)
  EnvPop(MAX_ENVELOPE - 1) As Single 'Probability of Penetration
  optType As Integer       'Optical Canopy Type 1=user-def 2=library
  LibHgt As Single         'Height (m)
  LibLAI As Single         'Leaf Area Index
  LibB As Single           'Coefficient B
  LibC As Single           'Coefficient C
  NumLAI As Integer        'Number of LAI levels
  LAIHgt(MAX_LAI - 1) As Single 'Height (m)
  LAICum(MAX_LAI - 1) As Single 'Cumulative LAI
  temp As Single           'Temperature within canopy (deg C)
  Humidity As Single       'Relative Humidity Within anopy (%)
  NDRuff As Single         'Nondimensional Surface Roughness
  NDDisp As Single         'Nondimensional Displacement
  Height As Single         'Canopy Height for Tier II/FS/Aerial (m)
End Type

'Terrain Data
Public Type TerrainData
  Zref As Single           'Ground Reference (m)
  Upslope As Single        'Ground upslope angle (deg)
  Sideslope As Single      'Ground sideslope angle (deg)
End Type

'Control Data
Public Type ControlData
  Height As Single         'Boom Height (m)
  NumLines As Integer      'Number of Spray Lines
  SwathWidthType As Integer 'SW type: 0=fixed 1=1.2WS 2=fact*WS
  SwathWidth As Single     'Swath Width (none or m)
  SwathDispType As Integer 'Disp type: 0=frac sw 1=frac ar 2=fixed 3=ac ctrline
  SwathDisp As Single      'Swath Displacement (none or m)
  FluxPlane As Single      'Flux Plane Location (m)
  TransHgtMin As Single    'Min Transport Height (m)
  TransHgtMax As Single    'Max Transport Height (m)
  MaxComputeTime As Single 'Max Compute time (sec)
  MaxDownwindDist As Single 'Max Downwind distance for computations (m)
  HalfBoom As Integer      'Half-boom application flag 0=full boom 1=half boom
End Type

'define a data type to hold the user input data
'This type is passed to the Fortran DLL. If you change
'its structure, you must also change it in the DLL.
Public Type UserData
  Tier As Integer          'Current Tier: 1, 2, or 3
  ApplMethod As Integer    'Application Method 0=aerial 1=ground 2=orchard/airblast
  Smokey As Integer        'Audience 0=regulatory (SDTF) 1=application (FS)
  Title As String          'Run title
  Notes As String          'User notes
  GA As GroundApplData     'Ground Application Data
  OA As OrchardAirblastData 'Orchard Airblast Data
  DSD(2) As DropSizeDistData  'Drop Szie Distribution Data
  BK(2) As DropKirkData          'DropKirk Data
  HK(2) As HKData          'FS Rotary Atomizer Data
  DK(2) As DropKickData    'DropKick Data
  SM As SprayMaterialData  'Spray Material Data
  AC As AircraftData       'Aircraft Data
  NZ As NozzleData         'Nozzle Data
  MET As MetData           'Meteorological Data
  CAN As CanopyData        'Canopy Data
  TRN As TerrainData       'Terrain Data
  CTL As ControlData       'Control Data
End Type

'======= end of User Input data area =====================

'======= User Calculation data area ============================
' The structures defined below are also duplicated in agdstruc.inc

'define a type to hold user calculation results
Public Type UserCalc
  Valid As Integer         'If true, the calcs have been performed
  CodeVersion As Single    'the version of AgDRIFT that performed these calcs
  StartDate As String      'the date when the calcs were started
  StartTime As String      'the time when the calcs were started
  MaxErrorLevel As Integer 'Max error level encountered during run 0=none 1=warn 2=err
  MessageLog As String     'A log of calculation messages
  NumDep As Integer        'Number of Deposition Data Points
  DepDist() As Single      'Deposition Distance (m)
  DepVal() As Single       'Deposition (fraction of applied)
  DepExtrap As Single      'if > 0, X value at which extrapolation begins
  NumPID As Integer        'Number of Pond-Int Depos Data Points
  PIDDist() As Single      'PID Distance (m)
  PIDVal() As Single       'PID (fraction of applied)
  PIDExtrap As Single      'if > 0, X value at which extrapolation begins
  NumFlux As Integer       'Number of Flux Data Points
  FluxDist() As Single     'Flux Distance (m)
  FluxVal() As Single      'Flux (mg/m2)
  NumCOV As Integer        'Number of COV/ESW data points
  COVVal() As Single       'COV points
  COVESW() As Single       'Effective Swath Width values for COV's
  NumConc As Integer       'Number of concentration data points
  ConcDist() As Single     'Conc distance (m)
  ConcVal() As Single      'Conc values (ng/L)(ppt)
  SwathDisp As Single      'Swath Displacement (m)
  SBCOV As Single          'Spray Block COV
  SBMeanDep As Single      'Spray Block Mean Deposition
  Swath As Single          'Swath Width (m)
  AirborneDrift As Single  '% Airborne Drift
  EvapFrac As Single       'Evaporated Fraction
  AppEff As Single         'Application Efficiency (%)
  DownwindDep As Single    'Downwind Deposition (%)
  CanopyDep As Single      'Canopy Deposition
  NumSgl As Integer        'Number of Single-Line Depos Points
  SglDist() As Single      'Single-Line Depos Distance (m)
  SglVal() As Single       'Single-Line Depos (%)
  NumFA As Integer         'Number of Fraction Aloft Points
  FADist() As Single       'Fraction Aloft Distance (m)
  FAVal() As Single        'Fraction Aloft
  NumCOVM As Integer       'Number of COV Mean Deposition Points
  COVMDist() As Single     'COV Mean Deposition Distance (m)
  COVMVal() As Single      'COV Mean Deposition
  NumHalf As Integer       'Number Sgl Swath Upwind Half Boom
  HalfDist() As Single     'Sgl Swath Upwind Half Boom Dist(m)
  HalfVal() As Single      'Sgl Swath Upwind Half Boom depos
  NumSBD As Integer        'Number of Spray Block Depos pts
  SBDDist() As Single      'Spray Block Depos Distance (m)
  SBDVal() As Single       'Spray Block Depos
  NumCAN As Integer        'Number of Canopy Depos Points
  CANDist() As Single      'Canopy Deposition Distance (m)
  CANVal() As Single       'Canopy Deposition
  NumTAA As Integer        'Num Time Accountancy Aloft Pts
  TAATime() As Single      'Time Accountancy Aloft Time (sec)
  TAAVal() As Single       'Time Accountancy Aloft
  NumTAV As Integer        'Num Time Accountancy Vapor Pts
  TAVTime() As Single      'Time Accountancy Vapor Time (sec)
  TAVVal() As Single       'Time Accountancy Vapor
  NumTAC As Integer        'Num Time Accountancy Canopy Pts
  TACTime() As Single      'Time Accountancy Canopy Time (sec)
  TACVal() As Single       'Time Accountancy Canopy
  NumTAG As Integer        'Num Time Accountancy Ground Pts
  TAGTime() As Single      'Time Accountancy Ground Time (sec)
  TAGVal() As Single       'Time Accountancy Ground
  NumDAA As Integer        'Num Distance Accountancy Aloft Pts
  DAADist() As Single      'Distance Accountancy Aloft Distance (sec)
  DAAVal() As Single       'Distance Accountancy Aloft
  NumDAV As Integer        'Num Distance Accountancy Vapor Pts
  DAVDist() As Single      'Distance Accountancy Vapor Distance (sec)
  DAVVal() As Single       'Distance Accountancy Vapor
  NumDAC As Integer        'Num Distance Accountancy Canopy Pts
  DACDist() As Single      'Distance Accountancy Canopy Distance (sec)
  DACVal() As Single       'Distance Accountancy Canopy
  NumDAG As Integer        'Num Distance Accountancy Ground Pts
  DAGDist() As Single      'Distance Accountancy Ground Distance (sec)
  DAGVal() As Single       'Distance Accountancy Ground
  NumHAA As Integer        'Num Height Accountancy Aloft Pts
  HAAHgt() As Single       'Height Accountancy Aloft Height (m)
  HAAVal() As Single       'Height Accountancy Aloft
  NumHAV As Integer        'Num Height Accountancy Vapor Pts
  HAVHgt() As Single       'Height Accountancy Vapor Height (m)
  HAVVal() As Single       'Height Accountancy Vapor
  NumHAC As Integer        'Num Height Accountancy Canopy Pts
  HACHgt() As Single       'Height Accountancy Canopy Height (m)
  HACVal() As Single       'Height Accountancy Canopy
  NumSBDSD As Integer      'Num Spray Block Drop Size Dist Pts
  SBDSDDiam() As Single    'Spray Block DSD Diameter (micron)
  SBDSDFrac() As Single    'Spray Block DSD Mass Frac
  NumDWDSD As Integer      'Num Downwind Drop Size Dist Pts
  DWDSDDiam() As Single    'Downwind DSD Diameter (micron)
  DWDSDFrac() As Single    'Downwind DSD Mass Frac
  NumFXDSD As Integer      'Num Vert Flux Drop Size Dist Pts
  FXDSDDiam() As Single    'Vert Flux DSD Diameter (micron)
  FXDSDFrac() As Single    'Vert Flux DSD Mass Frac
  NumCNDSD As Integer      'Num Canopy Drop Size Dist Pts
  CNDSDDiam() As Single    'Canopy DSD Diameter (micron)
  CNDSDFrac() As Single    'Canopy DSD Mass Frac
  NumSBAC As Integer       'Num Spray Block Area Coverage Pts
  SBACRate() As Single     'Spray Block Fraction of Application Rate
  SBACFrac() As Single     'Spray Block Fraction of Area Covered
  NumLAY As Integer        'Num Application Layout Pts
  LAYDist() As Single      'Application Layout Distance (m)
  LAYFrac() As Single      'Application Layout Frac Tank Mix
End Type

'======= end of User Calculation data area ============================

'Interface-related data
Public Type InterfaceData
  HasConfidentialData As Boolean  'True if SDTF library is present
  FileName As String         'Current user data file name
  LibraryPath As String      'Full path to library database
  MAALibraryPath As String   'Full path to MAA library database
  DataChanged As Integer     'If True, data was changed, but not saved
  CalcsBatchMode As Integer  'If true, bring up calcs form in batch mode
  OkToDoCalcs As Integer     'Used to monitor calc progress
  DataNeedsChecking As Integer 'Check data if true
  CalcsInProgress As Integer 'the calculations are happening
  StartCalcsOnLoad As Integer 'if true, start calcs right away
  RevertCalcsAvailable As Integer 'If true, a saved set of calcs is available for reloading
  PlotVar As Long         'current plot variable
  LastPlotVar As Long     'previous value of PlotVar
End Type

'define a type to hold preferences
Public Type UserPrefsData
  InitialTier As Integer      'tier level at program start
  InitialAM As Integer        'application method at program start
  InitialAUD As Integer       'audience at application start
  WarnOnTierChange As Integer 'warn user when changing tiers
  PauseBeforeCalc As Integer  'If true, the user starts the calcs
  SuppressTier3Warn As Integer 'if true, don't ask user about calc warnings
  Units As Integer            '0=imperial 1=metric
  UserLib As String           'path to user library file
End Type

'******************** Global variables *******************
Public UI As InterfaceData
Public UD As UserData
Public UC As UserCalc
Public UP As UserPrefsData
Public NZ2 As NozzleData    'work area used by frmNozzles and frmDropKick
Public DK2 As DropKickData  'work area used by frmDropSize and frmDropKick
Public BK2 As DropKirkData  'work area used by frmDropSize and frmDropKirk

Public gobjErrors As clsErrors 'Error handling

'*************** AgDRIFT Global constants ****************
'User file format version
'0 -> AgDRIFT 1.0 - 1.05: Initial file format
'1 -> AgDRIFT 1.06 - 1.07 : Added UD.NZ.BoomWidth
'2 -> AgDRIFT 1.08        : Added UD.OA .BasicType, .BegTRow, .EndTrow
'                           Added UD.SM .NonVGrav
'                           Added UD.AC .NumEng .EngHoriz(1) .WingVert
'                                       .BoomVert .BoomFwd
'                           Array UD.NZ .PosVert(60), .PosFwd(60),
'                           Added UD.NZ .NozType(60)
'                           Change UD.MET .HGTCAN to .SurfRough
'                           Added UD.Ctl .Zref .HalfBoom
'                           Array UD.DSD(3), UD.DK(3)
'3 -> AgDRIFT 1.09        : Added UD.BK(3)
'                           Added UD.DK.SprayAngle, .SprayType
'4 -> AgDRIFT 2.05 (2.0.5): Write entire UD.AC.EngHoriz array, rather than
'                           looping on NumEng, since NumEng does not
'                           describe the number of valid array elements.
'5 -> AgDRIFT 2.09        : Added UD.Met.Insolation
'                           Added UD.HK()
'                           Added UD.BK().SpectrumSource
'                           Added UD.SM.CalcInputSelect .ActSolFrac .AddSolFrac
'                                 .ActNVFrac .AddNVFrac .FlowRateUnits
'                           Added UD.MET.Insolation
'
Public Const USERFILEVERSION = 5   'version number for user data files

'Library database format version
'Libraries must have a table "Info" with an integer field "Version"
'which matches this number.
'
'1 -> 1.09: Began file version checking
'---------
'Lib vers 1: Baseline library format
'Lib vers 2: Updated BasicDSD BasicDep BasicSgl contents
'            Added SprayQuality to atomize.out
'Lib vers 3: BasicDSD - split SwathDisp in two
'Lib vers 4: BasicDSD - split SwathDisp in three (Pub,Reg,FS)
'            Added ARSNozzle table
'Lib vers 5: Switch to 32-bit environment
'Lib vers 6: No format change. New lib vers number to lock out
'            users of previous agdrift versions.
'Lib vers 7: Added prefix support for ARS (DropKirk) Nozzle Orifice
'            and Modifier names. Expanded to 48 chars Nozzle Name.
'            Expanded to twenty max. permissable values in
'            Orifice values and Modifier values.
Public Const LIBRARYVERSION = 7

'Tiers
Public Const TIER_1 = 1
Public Const TIER_2 = 2
Public Const TIER_3 = 3

'Audience (Smokey)
Public Const AUD_SDTF = 0
Public Const AUD_FS = 1

'Application Method
Public Const AM_AERIAL = 0
Public Const AM_GROUND = 1
Public Const AM_ORCHARD = 2

'Options for FileDialog
Public Const FD_OPEN = 1
Public Const FD_SAVEAS = 2
Public Const FD_TYPE_USER = 1
Public Const FD_TYPE_TEXT = 2
Public Const FD_TYPE_LIB = 3

'**************** Links to DLLs and APIs *****************
Public Declare Function WinHelp Lib "User32" Alias "WinHelpA" _
      (ByVal hwnd As Long, ByVal lpHelpFile As String, _
      ByVal wCommand As Long, ByVal dwData As Long) As Long

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetPrivateProfileInt Lib "Kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetProfileInt Lib "Kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Public Declare Function GetProfileString Lib "Kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function WriteProfileString Lib "Kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetVersionString(VersionNumber) As String
  If VersionNumber <= 2.1 Then
    'Support versions 2.0.10 and earlier
    GetVersionString = Format$(Int(VersionNumber)) & "." & _
                       Format$((VersionNumber - Fix(VersionNumber)), "0.00#")
  Else
    'Version 2.1.1 and newer
    GetVersionString = Format$(Int(VersionNumber / 10)) & "." & _
                       Format$(Int(VersionNumber) - (Int(VersionNumber / 10) * 10)) & "." & _
                       Format$(Int((VersionNumber - Fix(VersionNumber)) * 100))
  End If
End Function

Public Function CheckData(outctl As Control) As Integer
'Init fortran calc routines and perform a data check
'
'outctl - a listbox to receive status messages
'
  'args for agread
  Dim iunits As Long
  Dim istat As Long
  Dim idk As Long
  Dim itype As Long
  ReDim idat(2) As Long
  ReDim adat(2) As Single
  Dim cdat As String
  Dim clen As Long
  Dim s As String
  Dim squal As Integer
  
  Dim DKAction As Integer
  Dim DKshowmsg As Integer
  Dim DKanswer As Integer
  Dim np As Integer
  ReDim Diam(MAX_DROPS - 1) As Single
  ReDim mfrac(MAX_DROPS - 1) As Single
  Dim Msg As String
  Dim MBType As Integer
  Dim outstr As String
  Dim minstr As String
  Dim maxstr As String

  'Add a message to the calc log
  Msg = ""
  AppendStr Msg, Date$ & " " & Time$, False
  AppendStr Msg, " Checking data", False
  AddToLog outctl, Msg
  AppendStr UC.MessageLog, Msg, True

  CheckData = False 'default return value

  'Pass the User Data to the fortran
  Call aginit(UD, AGINIT_NORMAL)

  'Get and display summary messages
  ' agread info: istat 0=don't write data, keep reading
  '                    1=write data, keep reading
  '                    2=error: write data, STOP
  '                    3=end of data
  '                    4=DropKick message, allow continue
  '                    5=DropKick message, do not allow continue
  '              idk   0=density mismatch in DropKick
  '                    1=speed mismatch in DropKick
  '                    2=Q mismatch in DropKick
  '                      Note: idk=2 is a special case. No re-
  '                      calculation is possible, so we can't
  '                      display a recompute button.
  '              itype 0=char data only
  '                    1=int and char data
  '                    2=real and char data
  '          adat/idat for istat=1-3: (0)=value (1)=min (2)=max
  '                    for istat=4-5: (0)=index of offending DSD
  '                                   (1)=AgDRIFT value
  '                                   (2)=DropKick value
  UC.MaxErrorLevel = 0 'reset the max error level
  iunits = UP.Units    'set the units flag
  cdat = Space$(40)    'make space for the return string
  Do
    Call agread(iunits, istat, idk, itype, adat(0), cdat, clen)
    Select Case istat  'error level
      Case 0: 'normal, keep reading
        Select Case itype
          Case 0: 'no message or data
          Case 1: 'message
            Msg = Trim$(cdat)
            AddToLog outctl, Msg
            AppendStr UC.MessageLog, Msg, True
          Case 2: 'message and float data
            Msg = cdat & " " & AGFormat$(adat(0))
            AddToLog outctl, Msg
            AppendStr UC.MessageLog, Msg, True
          Case 3: 'message and integer data
            Msg = cdat & " " & Format$(CInt(adat(0)))
            AddToLog outctl, Msg
            AppendStr UC.MessageLog, Msg, True
        End Select
      Case 1: 'warning with string
        If UC.MaxErrorLevel < 1 Then UC.MaxErrorLevel = 1
        Msg = ""
        AppendStr Msg, "Warning!", True
        Select Case itype
          Case 0: 'no message or data
          Case 1: 'message
            AddToLog outctl, Trim$(cdat)
            AppendStr Msg, cdat, True
          Case 2: 'message and float data
            AddToLog outctl, cdat & " " & AGFormat$(adat(0))
            AppendStr Msg, Chr$(34) & Trim$(cdat) & Chr$(34), True
            AppendStr Msg, "is out of range. The limits are:", True
            AppendStr Msg, "", True
            AppendStr Msg, "Min: " & AGFormat$(adat(1)), True
            AppendStr Msg, "Val: " & AGFormat$(adat(0)), True
            AppendStr Msg, "Max: " & AGFormat$(adat(2)), True
          Case 3: 'message and integer data
            AddToLog outctl, cdat & " " & Format$(CInt(adat(0)))
            AppendStr Msg, Chr$(34) & Trim$(cdat) & Chr$(34), True
            AppendStr Msg, "is out of range. The limits are:", True
            AppendStr Msg, "", True
            AppendStr Msg, "Min: " & Format$(CInt(adat(1))), True
            AppendStr Msg, "Val: " & Format$(CInt(adat(0))), True
            AppendStr Msg, "Max: " & Format$(CInt(adat(2))), True
        End Select
        'Add the Warning to the log
        AppendStr UC.MessageLog, Msg, False
        'Warn the user, if appropriate
        If Not UP.SuppressTier3Warn And Not UI.CalcsBatchMode Then
          AppendStr Msg, "", True
          AppendStr Msg, "Continue with calculations?", True
          If MsgBox(Msg, vbExclamation + vbYesNo) = vbNo Then
            Exit Do
          End If
        End If
      Case 2: 'error with string; stop
        If UC.MaxErrorLevel < 2 Then UC.MaxErrorLevel = 2
        Msg = ""
        AppendStr Msg, "Error!", True
        Select Case itype
          Case 0: 'no message or data
          Case 1: 'message
            AddToLog outctl, Trim$(cdat)
            AppendStr Msg, cdat, True
          Case 2: 'message and float data
            AddToLog outctl, cdat & " " & AGFormat$(adat(0))
            AppendStr Msg, Chr$(34) & Trim$(cdat) & Chr$(34), True
            AppendStr Msg, "is out of range. The limits are:", True
            AppendStr Msg, "", True
            AppendStr Msg, "Min: " & AGFormat$(adat(1)), True
            AppendStr Msg, "Val: " & AGFormat$(adat(0)), True
            AppendStr Msg, "Max: " & AGFormat$(adat(2)), True
          Case 3: 'message and integer data
            AddToLog outctl, cdat & " " & Format$(CInt(adat(0)))
            AppendStr Msg, Chr$(34) & Trim$(cdat) & Chr$(34), True
            AppendStr Msg, "is out of range. The limits are:", True
            AppendStr Msg, "", True
            AppendStr Msg, "Min: " & Format$(CInt(adat(1))), True
            AppendStr Msg, "Val: " & Format$(CInt(adat(0))), True
            AppendStr Msg, "Max: " & Format$(CInt(adat(2))), True
        End Select
        'Add the Error to the log
        AppendStr UC.MessageLog, Msg, False
        'Tell the user about the error and stop
        If Not UI.CalcsBatchMode Then
          MsgBox Msg, vbCritical + vbOKOnly
        End If
        Exit Do
      Case 3: 'normal end of data
        UI.DataNeedsChecking = False 'data has been checked and is okay
        CheckData = True
        Exit Do
      Case 4, 5: 'Special DropKick message: allow/do not allow continue
        'If certain of the input data to DropKick do not match the
        'corresponding input data in DropKick, there is a problem.
        'The user may recompute DropKick with the AgDRIFT values,
        'continue with everything as it is (only if ier=4), or cancel
        'cancel the whole thing. The value that does not match is
        'indicated by IDK. If IDK=2 (Q), do not allow the option of
        'recomputing.
        If UC.MaxErrorLevel < 1 Then UC.MaxErrorLevel = 1
        Msg = ""
        AppendStr Msg, "Warning!", True
        AddToLog outctl, cdat
        AppendStr Msg, Trim$(cdat), True
        AppendStr Msg, "", True
        AppendStr Msg, "    AgDRIFT:  " & AGFormat$(adat(1)), True
        AppendStr Msg, "    DropKick: " & AGFormat$(adat(2)), True
        AppendStr Msg, "", True
        'Add the message to the log
        AppendStr UC.MessageLog, Msg, False
        'DropKick may need to be rerun.
        'Set DKAction to:
        ' 1=recompute dropkick calcs, continue
        ' 2=do not recompute dropkick calcs, continue
        ' 3=abort entire calculation
        If UI.CalcsBatchMode Then
          'This is a batch calc. Do not prompt the user for action.
          DKshowmsg = False 'do not show DropKick messages
          'If possible, rerun the DK calcs without
          'asking the user anything.
          Select Case idk
          Case 0, 1 'density, speed
            DKAction = 1 'recompute
          Case 2    'Q
            DKAction = 3 'halt
          End Select
        Else
          'This is a regular calc. Prompt the user for action.
          DKshowmsg = True 'show DropKick messages
          AppendStr Msg, "", True
          Select Case istat
          Case 4 'allow calcs to continue
            Select Case idk
            Case 0, 1 'density, speed
              AppendStr Msg, "Recompute DropKick before continuing?", True
              Select Case MsgBox(Msg, vbExclamation + vbYesNoCancel)
              Case vbYes
                DKAction = 1 'recompute
              Case vbNo
                DKAction = 2 'don't recompute
              Case vbCancel
                DKAction = 3 'halt
              End Select
            Case 2  'Q
              AppendStr Msg, "Continue?", True
              Select Case MsgBox(Msg, vbExclamation + vbOKCancel)
              Case vbOK
                DKAction = 2 'don't recompute
              Case vbCancel
                DKAction = 3 'halt
              End Select
            End Select
          Case 5 'do not allow calcs to continue
            Select Case idk
            Case 0, 1 'density, speed
              AppendStr Msg, "DropKick will be recomputed.", True
              Select Case MsgBox(Msg, vbExclamation + vbOKCancel)
              Case vbOK
                DKAction = 1 'recompute
              Case vbCancel
                DKAction = 3 'halt
              End Select
            Case 2 'Q
              AppendStr Msg, "Cannot continue calculations.", True
              Select Case MsgBox(Msg, vbExclamation + vbOKOnly)
              Case vbOK
                DKAction = 3 'halt
              End Select
            End Select
          End Select
        End If
        'Take the appropriate DropKick action
        Select Case DKAction
        Case 1 'recompute
          iDSD = Int(adat(0)) 'Recover which DSD had the problem
          Msg = "Recomputing DropKick with new parameters"
          AppendStr UC.MessageLog, Msg, True
          AddToLog outctl, Msg
          'override certain DK values with those of the main screen
          Select Case idk
          Case 0: 'Density
            UD.DK(iDSD).Density = adat(1)
          Case 1: 'Speed
            UD.DK(iDSD).Speed = adat(1)
          Case 2: 'Q
            'We should never get here, since we can't recalc for Q
            Exit Do 'jump out of the checking loop to halt calcs
          End Select
          'redo dropkick calcs
          If CalcDropKick(UD.DK(iDSD), DKshowmsg, UD.DK(iDSD).MaxErrorLevel, _
                          squal, np, Diam(), mfrac()) Then
            'DropKick done, store new dsd
            If squal < 0 Then
              UD.DSD(iDSD).NumDrop = np
              For i = 0 To UD.DSD(iDSD).NumDrop - 1
                UD.DSD(iDSD).Diam(i) = Diam(i)
                UD.DSD(iDSD).MassFrac(i) = mfrac(i)
              Next
            Else
              'spray quality: get basic DSD using returned index
              GetBasicDataDSD squal, UD.DSD(iDSD)
            End If
            'update the agdrift calc routines
            Call agupds(UD, CLng(iDSD))
            're-init the checking loop
            Msg = "Re-checking all input data"
            AppendStr UC.MessageLog, Msg, True
            AddToLog outctl, Msg
            Call aginit(UD, AGINIT_NORMAL) 'cause agread to start over
          Else
            Msg = "DropKick calcs unsuccessful"
            AppendStr UC.MessageLog, Msg, True
            AddToLog outctl, Msg
            Exit Do 'DropKick failed, halt calcs.
          End If
        Case 2 'don't recompute
          Msg = "Not recomputing DropKick"
          AppendStr UC.MessageLog, Msg, True
          AddToLog outctl, Msg
        Case 3 'halt
          Exit Do 'jump out of the checking loop
        End Select
      Case Else 'just in case...
        Exit Do
    End Select
    DoEvents
  Loop While True
End Function

Public Sub AddToLog(c As Control, s As String)
'Add a string to the Log control
  On Error Resume Next
  Dim fsave As Control
  c.AddItem s
  c.Refresh
  Set fsave = ActiveControl 'save current control
  c.SetFocus           'set focus to list box
  SendKeys "{END}"                  'send an END key to the list box
  DoEvents
  c.Selected(c.ListIndex) = False   'turn off highlight
  fsave.SetFocus                    'restore original focus
End Sub

Public Function CalcDropKick(xDK As DropKickData, showmsg As Integer, _
  MaxErrLev As Integer, squal As Integer, _
  np As Integer, Diam() As Single, mfrac() As Single) As Integer
'Calculate DropKick drop distribution
'
' xUD       i   UserData structure containing input DK data
' showmsg   i   if true, show calc warning and error messages
' MaxErrLev o   maximum error level during calculations
' squal     o   spray quality -1:none,use np/diam/mfrac  0-10:basic dsd
' np        o   number of points in diam/mfrac arrays
' Diam      o   array of drop diameters
' mfrac     o   array of mass fractions
'
' function returns true if calcs are successful, false if not
'
  'agkick arguments
  Dim nd As Long
  Dim ier As Long
  Dim iunits As Long
  Dim lfl As Long
  Dim iqual As Long
  ReDim realwd(2) As Single
  Dim cdat As String
  Dim clen As Long
  
  Dim minstr As String
  Dim maxstr As String
  Dim outstr As String

  'call agkick to do the dropkick calcs
  '
  'lfl   - agkick internal counter: init to 0
  'iqual - spray quality -1:none  0-10: basic dsd index
  'ier   - returned status 0=OK, done.
  '                        1=warning with msg and data
  '                        2=err with msg and data
  '                        3=warning with msg
  '                        4=err with msg
  '                        5=msg
  'realwd - error data
  'cdat,clen - error message and length
  '
  CalcDropKick = False      'default return value
  MaxErrLev = 0             'reset max error level
  iunits = UP.Units         'get units from preferences
  lfl = 0                   'reset agkick counter
  cdat = Space$(40)         'allocate string space
  Do
    Call agkick(xDK, iunits, lfl, iqual, _
                nd, Diam(0), mfrac(0), ier, realwd(0), cdat, clen)
    squal = CInt(iqual) 'convert to integer
    np = CInt(nd) 'convert to integer
    Select Case ier
      Case 0  'success, calcs are done
        CalcDropKick = True
        Exit Do
      Case 1 'warning with msg and data
        If MaxErrLev < 1 Then MaxErrLev = 1
        If showmsg Then
          outstr = AGFormat$(realwd(0))
          minstr = AGFormat$(realwd(1))
          maxstr = AGFormat$(realwd(2))
  
          Msg = "Warning!" + Chr$(13)
          Msg = Msg + Chr$(34) + Trim$(cdat) + Chr$(34) + Chr$(13)
          Msg = Msg + "is out of the suggested range. The limits are:" + Chr$(13)
          Msg = Msg + Chr$(13)
          Msg = Msg + "Min: " + minstr + Chr$(13)
          Msg = Msg + "Val: " + outstr + Chr$(13)
          Msg = Msg + "Max: " + maxstr + Chr$(13)
          Msg = Msg + Chr$(13)
          Msg = Msg + "Continue with calculations?"
          If MsgBox(Msg, vbExclamation + vbYesNo) = vbNo Then
            Exit Do
          End If
        End If
      Case 2 'error with msg and data
        If MaxErrLev < 2 Then MaxErrLev = 2
        If showmsg Then
          outstr = AGFormat$(realwd(0))
          minstr = AGFormat$(realwd(1))
          maxstr = AGFormat$(realwd(2))
  
          Msg = "Error!" + Chr$(13)
          Msg = Msg + Chr$(34) + Trim$(cdat) + Chr$(34) + Chr$(13)
          Msg = Msg + "is out of range. The limits are:" + Chr$(13)
          Msg = Msg + Chr$(13)
          Msg = Msg + "Min: " + minstr + Chr$(13)
          Msg = Msg + "Val: " + outstr + Chr$(13)
          Msg = Msg + "Max: " + maxstr
          MsgBox Msg, vbCritical + vbOKOnly
        End If
        Exit Do
      Case 3 'warning with msg
        If MaxErrLev < 1 Then MaxErrLev = 1
        If showmsg Then
          Msg = "Warning! "
          Msg = Msg & Left$(cdat, clen)
          Msg = Msg + Chr$(13)
          Msg = Msg + "Continue with calculations?"
          If MsgBox(Msg, vbExclamation + vbYesNo) = vbNo Then
            Exit Do
          End If
        End If
      Case 4 'error with msg
        If MaxErrLev < 2 Then MaxErrLev = 2
        If showmsg Then
          Msg = "Error! "
          Msg = Msg & Left$(cdat, clen)
          MsgBox Msg, vbCritical + vbOKOnly
        End If
        Exit Do
      Case 5 'informational msg
        If showmsg Then
          Msg = Left$(cdat, clen)
          MsgBox Msg, vbInformation + vbOKOnly
        End If
      Case Else 'just in case...
        Exit Do
    End Select
  Loop
End Function

Public Function CalcDropKirk(xBK As DropKirkData, showmsg As Integer, _
  MaxErrLev As Integer, squal As Integer, _
  np As Integer, Diam() As Single, mfrac() As Single) As Integer
'Calculate DropKirk drop distribution
'
' xUD       i   UserData structure containing input BK data
' showmsg   i   if true, show calc warning and error messages
' MaxErrLev o   maximum error level during calculations
' squal     o   spray quality -1:none,use np/diam/mfrac  0-10:basic dsd
' np        o   number of points in diam/mfrac arrays
' Diam      o   array of drop diameters
' mfrac     o   array of mass fractions
'
' function returns true if calcs are successful, false if not
'
  'agkirk arguments
  Dim nd As Long
  Dim ier As Long
  Dim iunits As Long
  Dim lfl As Long
  Dim iqual As Long
  ReDim realwd(2) As Single
  Dim cdat As String
  Dim clen As Long
  
  Dim minstr As String
  Dim maxstr As String
  Dim outstr As String

  'call agkirk to do the dropkirk calcs
  '
  'lfl   - agkirk internal counter: init to 0
  'iqual - spray quality -1:none  0-10:basic dsd index
  'ier   - returned status 0=OK, done.
  '                        1=warning with msg and data
  '                        2=err with msg and data
  '                        3=warning with msg
  '                        4=err with msg
  '                        5=msg
  'realwd - error data
  'cdat,clen - error message and length
  '
  CalcDropKirk = False      'default return value
  MaxErrLev = 0             'reset max error level
  iunits = UP.Units         'get units from preferences
  lfl = 0                   'reset agkirk counter
  cdat = Space$(40)         'allocate string space
  Do
    Call agkirk(xBK, iunits, lfl, iqual, _
                nd, Diam(0), mfrac(0), ier, realwd(0), cdat, clen)
    squal = CInt(iqual) 'convert to integer
    np = CInt(nd) 'convert to integer
    Select Case ier
      Case 0  'success, calcs are done
        CalcDropKirk = True
        Exit Do
      Case 1 'warning with msg and data
        If MaxErrLev < 1 Then MaxErrLev = 1
        If showmsg Then
          outstr = AGFormat$(realwd(0))
          minstr = AGFormat$(realwd(1))
          maxstr = AGFormat$(realwd(2))
  
          Msg = "Warning!" + Chr$(13)
          Msg = Msg + Chr$(34) + Trim$(cdat) + Chr$(34) + Chr$(13)
          Msg = Msg + "is out of the suggested range. The limits are:" + Chr$(13)
          Msg = Msg + Chr$(13)
          Msg = Msg + "Min: " + minstr + Chr$(13)
          Msg = Msg + "Val: " + outstr + Chr$(13)
          Msg = Msg + "Max: " + maxstr + Chr$(13)
          Msg = Msg + Chr$(13)
          Msg = Msg + "Continue with calculations?"
          If MsgBox(Msg, vbExclamation + vbYesNo) = vbNo Then
            Exit Do
          End If
        End If
      Case 2 'error with msg and data
        If MaxErrLev < 2 Then MaxErrLev = 2
        If showmsg Then
          outstr = AGFormat$(realwd(0))
          minstr = AGFormat$(realwd(1))
          maxstr = AGFormat$(realwd(2))
  
          Msg = "Error!" + Chr$(13)
          Msg = Msg + Chr$(34) + Trim$(cdat) + Chr$(34) + Chr$(13)
          Msg = Msg + "is out of range. The limits are:" + Chr$(13)
          Msg = Msg + Chr$(13)
          Msg = Msg + "Min: " + minstr + Chr$(13)
          Msg = Msg + "Val: " + outstr + Chr$(13)
          Msg = Msg + "Max: " + maxstr
          MsgBox Msg, vbCritical + vbOKOnly
        End If
        Exit Do
      Case 3 'warning with msg
        If MaxErrLev < 1 Then MaxErrLev = 1
        If showmsg Then
          Msg = "Warning! "
          Msg = Msg & Left$(cdat, clen)
          Msg = Msg + Chr$(13)
          Msg = Msg + "Continue with calculations?"
          If MsgBox(Msg, vbExclamation + vbYesNo) = vbNo Then
            Exit Do
          End If
        End If
      Case 4 'error with msg
        If MaxErrLev < 2 Then MaxErrLev = 2
        If showmsg Then
          Msg = "Error! "
          Msg = Msg & Left$(cdat, clen)
          MsgBox Msg, vbCritical + vbOKOnly
        End If
        Exit Do
      Case 5 'informational msg
        If showmsg Then
          Msg = Left$(cdat, clen)
          MsgBox Msg, vbInformation + vbOKOnly
        End If
      Case Else 'just in case...
        Exit Do
    End Select
  Loop
End Function

Sub LoadTier1Data(xUD As UserData, xUC As UserCalc)
'Load all tier1 data and "calcs" into the given UserData structure
'The Following elements of the UserData Structure must be set:
' xUD.ApplMethod
' xUD.GA.BasicType, .NumSwaths  (if xUD.ApplMethod = 1) or
' xUD.OA.BasicType, .BegTrow, .EndTrow  or
' xUD.DSD(0).BasicType (if xUD.ApplMethod = 0)
  
  'fill in input data
  Select Case xUD.ApplMethod
  Case AM_AERIAL
    GetBasicDataDSD xUD.DSD(0).BasicType, xUD.DSD(0) 'Dropsize distribution
    GetBasicDataAC 3, xUD.AC                         'Air Tractor 401
    GetBasicDataNZ 3, xUD.NZ                         'Air Tractor 401
    'Adjust Swath Displacement to suit new DSD data
    xUD.CTL.SwathDispType = 0 'Fraction of swath width
    GetBasicDataDSDSwathDisp xUD.DSD(0).BasicType, xUD.CTL.SwathDisp
  End Select
  
  'Now get the calcs from the database
  Select Case xUD.ApplMethod
  Case AM_AERIAL
    GetTier1Calcs xUD.ApplMethod, xUD.DSD(0).BasicType, 0, 0, xUC
    xUC.Valid = True 'that's it for tier 1 calcs
  Case AM_GROUND
    GetTier1Calcs xUD.ApplMethod, xUD.GA.BasicType, xUD.GA.NumSwaths, 0, xUC
    xUC.Valid = True 'that's it for tier 1 calcs
  Case AM_ORCHARD
    GetTier1Calcs xUD.ApplMethod, xUD.OA.BasicType, xUD.OA.BegTrow, xUD.OA.EndTrow, xUC
    xUC.Valid = True 'that's it for tier 1 calcs
  End Select
End Sub

Public Function AdjustBasicNozzles(BoomWidth As Single, SemiSpan As Single, xNZ As NozzleData)
'Given a set of nozzle data fresh from GetBasicDataNZ, adjust the horizontal
'positions to be consistent with the BoomWidth (a percentage) and the
'aircraft SemiSpan. Note that the BoomWidth stored in xNZ is not affected.
  Dim NumNoz As Long
  Dim PosHoriz(MAX_NOZZLES - 1) As Single
  
  'generate the adjusted distribution with agnozl
  Call agnozl(CLng(xNZ.NumNoz), xNZ.PosHoriz(0), _
              SemiSpan, BoomWidth, _
              NumNoz, PosHoriz(0))
  'copy the new distribution
  xNZ.NumNoz = NumNoz
  CopyMemory xNZ.PosHoriz(0), PosHoriz(0), NumNoz * Len(xNZ.PosHoriz(0))
End Function

Function AGFormat$(X)
'Format numbers for display using special rules
  Dim s As String
  If Abs(X) > 999999.9999 Then                'huge numbers
    s = Format$(X, "Scientific")
  ElseIf Abs(X) > 0 And Abs(X) < 0.0001 Then  'tiny numbers
    s = Format$(X, "Scientific")
  ElseIf Abs(X) >= 1 Then                     'smallish numbers
    s = Format$(X, "#####0.##")
  Else                                        '0 and everything else
    s = Format$(X, "#####0.####")
  End If
  'get rid of hanging decimal points
  If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
  'return the formatted string
  AGFormat$ = s
End Function

Sub AppendStr(buf As String, newstr As String, addcr As Integer)
'append newstr to buf, add a <crlf> optionally
  buf = buf & newstr
  If addcr Then buf = buf & vbCrLf
End Sub

Sub CenterForm(f As Form)
'Center the form on the screen
  f.Left = (Screen.Width / 2) - (f.Width / 2)
  f.Top = (Screen.Height / 2) - (f.Height / 2)
End Sub

Public Sub UserDataDefault(xUD As UserData)
'Set program data to defaults as set by UD.Smokey
  Select Case UD.Smokey
  Case AUD_SDTF
    If AGDRIFTREGULATORY Then
      UserDataDefaultAgReg xUD
    Else
      UserDataDefaultAgPub xUD
    End If
  Case AUD_FS
    UserDataDefaultFS xUD
  End Select
End Sub

Public Sub UserDataDefaultFS(xUD As UserData)
'Set program data to Forestry default values

  'The Basic DSD is referenced in a couple of places
  'here. This constant makes sure that use is consistant.
  Const DEFAULT_BASICDSD = 4 'ASAE Fine to Medium
  
  'General
  With xUD
    .Tier = .Tier             'don't change
    .Title = "Untitled"
    .Notes = ""
    .ApplMethod = .ApplMethod 'don't change
    .Smokey = .Smokey         'don't change
  End With

  'Ground
  With xUD.GA
    .BasicType = 3       'high boom, medium
    .NumSwaths = 20      'Swaths for ground sprayer
  End With

  'Airblast
  With xUD.OA
    .BasicType = 0       'low boom, medium
    .BegTrow = 1         'Beginning tree row
    .EndTrow = 20        'Ending Tree Row
  End With
  
  'DSD
  GetBasicDataDSD DEFAULT_BASICDSD, xUD.DSD(0)
  GetBasicDataDSD DEFAULT_BASICDSD, xUD.DSD(1)
  GetBasicDataDSD DEFAULT_BASICDSD, xUD.DSD(2)

  'Rotary Atomizer
  For i = LBound(xUD.HK()) To UBound(xUD.HK())
    With xUD.HK(i)
      .MaxErrorLevel = 0
      .MatType = 0 'Water
      .RotType = 0 'AU4000
      .Speed = 0   'will be set to match aircraft
      .BladeAngle = 50 'deg
      .BladeRPM = 5000 'rpm
      .FlowRate = 0 'will be set to match SM flowrate
      .SprayType = 0   'Spray Quality
    End With
  Next
  
  'DropKirk
  For i = 0 To 2
    With xUD.BK(i)
      .MaxErrorLevel = 0
      .NozType = 0
      .NameNoz = "CP"
      .LNameNoz = 2
      .Orifice = 0.061
      .Speed = 44.7    'm/s (100 mph)
      .NozAngle = 30
      .Pressure = 4.137 'bar (60 psig)
      .SprayType = 0  'Spray Quality
      .SpectrumSource = 0 'Malvern
    End With
  Next

  'DropKick
  For i = 0 To 2
    With xUD.DK(i)
      .MaxErrorLevel = 0
      .NozType = 1       'library
      .NameNoz = "D4-45"
      .LNameNoz = 5
      .VMD = 280.3
      .RelSpan = 0.9578
      .EffDiam = 0.11
      .SprayAngle = 65
      .MatType = 1       'library
      .NameMat = "IWS-0069"
      .LNameMat = 8
      .DynSurfTens = 73
      .ShearVisc = 1
      .Density = 0.997
      .ElongVisc = 3
      .Speed = 44.7    'm/s (100 mph)
      .NozAngle = 45
      .Pressure = 4.137 'bar (60 psig)
      .FlowType = 0
      .flow = 0
      .SprayType = 0  'Default to Spray Quality
    End With
  Next

  'Spray Mat
  With xUD.SM
    .Type = 0        'Basic
    .BasicType = 1   'Water
    .Name = GetBasicNameSM(.BasicType)
    .LName = Len(.Name)
    .CalcInputSelect = 0 'Input rates
    .NVFrac = 0.03
    .ACFrac = 0.015
    .ActSolFrac = 0.015
    .AddSolFrac = 0.015
    .ActNVFrac = 1
    .AddNVFrac = 1
    .FlowRate = 18.71 'L/ha
    .FlowRateUnits = 0 'L/ha
    .SpecGrav = 1
    .NonVGrav = 1
    .EvapRate = 84.76
  End With
  
  'Aircraft Data
  GetBasicDataAC 3, xUD.AC      'AT 401

  'Nozzle Distribution Data
  GetBasicDataNZ 3, xUD.NZ      'Same as aircraft

  'Met
  With xUD.MET
    .WS = 2.235           '5 mph
    .WD = -90             '-90 deg
    .WindHeight = 2
    .temp = 18.333        '65 deg F
    .Humidity = 50
    .Pressure = 1013
    .VortexDecay = 0.56
    .SurfRough = 0.0075   '.0246 ft
    .Insolation = 4       'Overcast (Day)
  End With

  'Canopy
  With xUD.CAN
    .Type = 3 'Basic
    .Name = ""
    .LName = 0
    .EleSiz = 0.02
    .StanDen = 494.2088  '200 st/ac
    .NumEnv = 0
'    .EnvHgt()
'    .EnvDiam()
'    .EnvPop()
    .optType = 1
    .LibHgt = 24.32
    .LibLAI = 2.57
    .LibB = 0.426
    .LibC = 1.856
    .NumLAI = 0
'    .LAIHgt()
'    .LAICum()
    .temp = xUD.MET.temp          'same as Met
    .Humidity = xUD.MET.Humidity  'same as Met
    .NDRuff = 0.14
    .NDDisp = 0.7
    .Height = 21.336  '70 ft
  End With
  
  'Terrain
  With xUD.TRN
    .Zref = 0
    .Upslope = 0
    .Sideslope = 0
  End With
  
  'Control
  With xUD.CTL
    .Height = 30.48       '100 ft
    .NumLines = 20
    .SwathWidthType = 0
    .SwathWidth = 18.2882 '60 ft
    .SwathDispType = 0    'Frac of swath width
    GetBasicDataDSDSwathDisp DEFAULT_BASICDSD, .SwathDisp
    .FluxPlane = 0
    .TransHgtMin = 0
    .TransHgtMax = 0
    .MaxComputeTime = 600 '600 sec
    .MaxDownwindDist = 795 '795 m
    .HalfBoom = 0 'full boom
  End With

End Sub

Public Sub UserDataDefaultAgPub(xUD As UserData)
'Set program data to default values

  'The Basic DSD is referenced in a couple of places
  'here. This constant makes sure that use is consistant.
  Const DEFAULT_BASICDSD = 6 'ASAE Medium to Coarse
  
  'General
  With xUD
    .Tier = .Tier             'don't change
    .Title = "Untitled"
    .Notes = ""
    .ApplMethod = .ApplMethod 'don't change
    .Smokey = .Smokey         'don't change
  End With

  'Ground
  With xUD.GA
    .BasicType = 3       'high boom, medium
    .NumSwaths = 20      'Swaths for ground sprayer
  End With

  'Airblast
  With xUD.OA
    .BasicType = 0       'low boom, medium
    .BegTrow = 1         'Beginning tree row
    .EndTrow = 20        'Ending Tree Row
  End With
  
  'DSD
  GetBasicDataDSD DEFAULT_BASICDSD, xUD.DSD(0)
  GetBasicDataDSD DEFAULT_BASICDSD, xUD.DSD(1)
  GetBasicDataDSD DEFAULT_BASICDSD, xUD.DSD(2)

  'Rotary Atomizer
  For i = LBound(xUD.HK()) To UBound(xUD.HK())
    With xUD.HK(i)
      .MaxErrorLevel = 0
      .MatType = 0 'Water
      .RotType = 0 'AU4000
      .Speed = 0   'will be set to match aircraft
      .BladeAngle = 50 'deg
      .BladeRPM = 5000 'rpm
      .FlowRate = 0 'will be set to match SM flowrate
      .SprayType = 0   'Spray Quality
    End With
  Next
  
  'DropKirk
  For i = 0 To 2
    With xUD.BK(i)
      .MaxErrorLevel = 0
      .NozType = 0
      .NameNoz = "CP"
      .LNameNoz = 2
      .Orifice = 0.061
      .Speed = 44.7    'm/s (100 mph)
      .NozAngle = 30
      .Pressure = 4.137 'bar (60 psig)
      .SprayType = 0  'Spray Quality
      .SpectrumSource = 0 'Malvern
    End With
  Next

  'DropKick
  For i = 0 To 2
    With xUD.DK(i)
      .MaxErrorLevel = 0
      .NozType = 1       'library
      .NameNoz = "D4-45"
      .LNameNoz = 5
      .VMD = 280.3
      .RelSpan = 0.9578
      .EffDiam = 0.11
      .SprayAngle = 65
      .MatType = 1       'library
      .NameMat = "IWS-0069"
      .LNameMat = 8
      .DynSurfTens = 73
      .ShearVisc = 1
      .Density = 0.997
      .ElongVisc = 3
      .Speed = 44.7    'm/s (100 mph)
      .NozAngle = 45
      .Pressure = 4.137 'bar (60 psig)
      .FlowType = 0
      .flow = 0
      .SprayType = 0  'Default to Spray Quality
    End With
  Next

  'Spray Mat
  With xUD.SM
    .Type = 0        'Basic
    .BasicType = 1   'Water
    .Name = GetBasicNameSM(.BasicType)
    .LName = Len(.Name)
    .CalcInputSelect = 0 'Input rates
    .NVFrac = 0.03
    .ACFrac = 0.015
    .ActSolFrac = 0.015
    .AddSolFrac = 0.015
    .ActNVFrac = 1
    .AddNVFrac = 1
    .FlowRate = 18.71 'L/ha
    .FlowRateUnits = 0 'L/ha
    .SpecGrav = 1
    .NonVGrav = 1
    .EvapRate = 84.76
  End With
  
  'Aircraft Data
  GetBasicDataAC 3, xUD.AC      'AT 401

  'Nozzle Distribution Data
  GetBasicDataNZ 3, xUD.NZ      'Same as aircraft

  'Met
  With xUD.MET
    .WS = 1.79            '4 mph
    .WD = -90             '-90 deg
    .WindHeight = 2
    .temp = 21.11         '70 deg F
    .Humidity = 75
    .Pressure = 1013
    .VortexDecay = 0.56
    .SurfRough = 0.01524  '.05 ft
    .Insolation = 4       'Overcast (Day)
  End With

  'Canopy
  With xUD.CAN
    .Type = 0
    .Name = ""
    .LName = 0
    .EleSiz = 0.02
    .StanDen = 494.2088  '200 st/ac
    .NumEnv = 0
'    .EnvHgt()
'    .EnvDiam()
'    .EnvPop()
    .optType = 1
    .LibHgt = 24.32
    .LibLAI = 2.57
    .LibB = 0.426
    .LibC = 1.856
    .NumLAI = 0
'    .LAIHgt()
'    .LAICum()
    .temp = xUD.MET.temp          'same as Met
    .Humidity = xUD.MET.Humidity  'same as Met
    .NDRuff = 0.14
    .NDDisp = 0.7
    .Height = 0
  End With
  
  'Terrain
  With xUD.TRN
    .Zref = 0
    .Upslope = 0
    .Sideslope = 0
  End With
  
  'Control
  With xUD.CTL
    .Height = 2.4384      '8 ft
    .NumLines = 20
    .SwathWidthType = 0
    .SwathWidth = 18.2882 '60 ft
    .SwathDispType = 0    'Frac of swath width
    GetBasicDataDSDSwathDisp DEFAULT_BASICDSD, .SwathDisp
    .FluxPlane = 0
    .TransHgtMin = 0
    .TransHgtMax = 0
    .MaxComputeTime = 600 '600 sec
    .MaxDownwindDist = 795 '795 m
    .HalfBoom = 0 'full boom
  End With

End Sub

Public Sub UserDataDefaultAgReg(xUD As UserData)
'Set program data to default values

  'The Basic DSD is referenced in a couple of places
  'here. This constant makes sure that use is consistant.
  Const DEFAULT_BASICDSD = 4 'ASAE Fine to Medium
  
  'General
  With xUD
    .Tier = .Tier             'don't change
    .Title = "Untitled"
    .Notes = ""
    .ApplMethod = .ApplMethod 'don't change
    .Smokey = .Smokey         'don't change
  End With

  'Ground
  With xUD.GA
    .BasicType = 3       'high boom, medium
    .NumSwaths = 20      'Swaths for ground sprayer
  End With

  'Airblast
  With xUD.OA
    .BasicType = 0       'low boom, medium
    .BegTrow = 1         'Beginning tree row
    .EndTrow = 20        'Ending Tree Row
  End With
  
  'DSD
  GetBasicDataDSD DEFAULT_BASICDSD, xUD.DSD(0)
  GetBasicDataDSD DEFAULT_BASICDSD, xUD.DSD(1)
  GetBasicDataDSD DEFAULT_BASICDSD, xUD.DSD(2)

  'Rotary Atomizer
  For i = LBound(xUD.HK()) To UBound(xUD.HK())
    With xUD.HK(i)
      .MaxErrorLevel = 0
      .MatType = 0 'Water
      .RotType = 0 'AU4000
      .Speed = 0   'will be set to match aircraft
      .BladeAngle = 50 'deg
      .BladeRPM = 5000 'rpm
      .FlowRate = 0 'will be set to match SM flowrate
      .SprayType = 0   'Spray Quality
    End With
  Next
  
  'DropKirk
  For i = 0 To 2
    With xUD.BK(i)
      .MaxErrorLevel = 0
      .NozType = 0
      .NameNoz = "CP"
      .LNameNoz = 2
      .Orifice = 0.061
      .Speed = 44.7    'm/s (100 mph)
      .NozAngle = 30
      .Pressure = 4.137 'bar (60 psig)
      .SprayType = 0  'Spray Quality
      .SpectrumSource = 1 'Malvern
    End With
  Next

  'DropKick
  For i = 0 To 2
    With xUD.DK(i)
      .MaxErrorLevel = 0
      .NozType = 1       'library
      .NameNoz = "D4-45"
      .LNameNoz = 5
      .VMD = 280.3
      .RelSpan = 0.9578
      .EffDiam = 0.11
      .SprayAngle = 65
      .MatType = 1       'library
      .NameMat = "IWS-0069"
      .LNameMat = 8
      .DynSurfTens = 73
      .ShearVisc = 1
      .Density = 0.997
      .ElongVisc = 3
      .Speed = 44.7    'm/s (100 mph)
      .NozAngle = 45
      .Pressure = 4.137 'bar (60 psig)
      .FlowType = 0
      .flow = 0
      .SprayType = 0  'Default to Spray Quality
    End With
  Next

  'Spray Mat
  With xUD.SM
    .Type = 0        'Basic
    .BasicType = 1   'Water
    .Name = GetBasicNameSM(.BasicType)
    .LName = Len(.Name)
    .CalcInputSelect = 0 'Input rates
    .NVFrac = 0.03
    .ACFrac = 0.015
    .ActSolFrac = 0.015
    .AddSolFrac = 0.015
    .ActNVFrac = 1
    .AddNVFrac = 1
    .FlowRate = 18.71 'L/ha
    .FlowRateUnits = 0 'L/ha
    .SpecGrav = 1
    .NonVGrav = 1
    .EvapRate = 84.76
  End With
  
  'Aircraft Data
  GetBasicDataAC 3, xUD.AC      'AT 401

  'Nozzle Distribution Data
  GetBasicDataNZ 3, xUD.NZ      'Same as aircraft

  'Met
  With xUD.MET
    .WS = 4.4704          '10 mph
    .WD = -90             '-90 deg
    .WindHeight = 2
    .temp = 30            '86 deg F
    .Humidity = 50
    .Pressure = 1013
    .VortexDecay = 0.56
    .SurfRough = 0.0075   '.0246 ft
    .Insolation = 4       'Overcast (Day)
  End With

  'Canopy
  With xUD.CAN
    .Type = 0
    .Name = ""
    .LName = 0
    .EleSiz = 0.02
    .StanDen = 494.2088  '200 st/ac
    .NumEnv = 0
'    .EnvHgt()
'    .EnvDiam()
'    .EnvPop()
    .optType = 1
    .LibHgt = 24.32
    .LibLAI = 2.57
    .LibB = 0.426
    .LibC = 1.856
    .NumLAI = 0
'    .LAIHgt()
'    .LAICum()
    .temp = xUD.MET.temp          'same as Met
    .Humidity = xUD.MET.Humidity  'same as Met
    .NDRuff = 0.14
    .NDDisp = 0.7
    .Height = 0
  End With
  
  'Terrain
  With xUD.TRN
    .Zref = 0
    .Upslope = 0
    .Sideslope = 0
  End With
  
  'Control
  With xUD.CTL
    .Height = 3.048       '10 ft
    .NumLines = 20
    .SwathWidthType = 0
    .SwathWidth = 18.2882 '60 ft
    .SwathDispType = 0    'Frac of swath width
    GetBasicDataDSDSwathDisp DEFAULT_BASICDSD, .SwathDisp
    .FluxPlane = 0
    .TransHgtMin = 0
    .TransHgtMax = 0
    .MaxComputeTime = 600 '600 sec
    .MaxDownwindDist = 795 '795 m
    .HalfBoom = 0 'full boom
  End With

End Sub

Function FileDialog(act, ftype, fname As String) As Integer
'Dialog box for obtaining a file name for UserData
' act - common dialog box action
'       1=FD_OPEN=Open 2=FD_SAVEAS=Save As
' ftype file filter type
'       1=FD_TYPE_USER=.agd 2=FD_TYPE_TEXT=.txt 3=FD_TYPE_LIB=.mdb
' fname the selected file name, if selection is successful.
'       fname is unchanged otherwise
'
' returns: true on OK, false on Cancel
'
  Dim d As Control
  Dim fn As String

  Set d = frmMain.CMDialog1
  '
  On Error GoTo ErrHandlerFND
  'Turn on CancelError
  d.CancelError = True
  'Set Default Extension
  Select Case ftype
  Case FD_TYPE_USER 'AgDRIFT user files
    'added if a file name is entered without an extension
    d.DefaultExt = "agd"
    'Set filter list
    d.Filter = "All Files (*.*)|*.*|Data Files (*.agd)|*.agd"
  Case FD_TYPE_TEXT 'Text files
    'added if a file name is entered without an extension
    d.DefaultExt = "txt"
    'Set filter list
    d.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt"
  Case FD_TYPE_LIB 'User library
    'added if a file name is entered without an extension
    d.DefaultExt = "mdb"
    'Set filter list
    d.Filter = "All Files (*.*)|*.*|Library Files (*.mdb)|*.mdb"
  End Select
  'Specify current filter
  d.FilterIndex = 2
  'Set the default file name
  d.FileName = fname
  'Set dialog flags
  d.Flags = cdlOFNHideReadOnly
  'Display the dialog box
  Select Case act
  Case FD_OPEN
    d.ShowOpen
  Case FD_SAVEAS
    d.ShowSave
  End Select
  'full file path is CMDialog1.FileName
  'file name only is CMDialog1.FileTitle
  
  'check for file existance for Save As Function
  fn = d.FileTitle
  If act = FD_SAVEAS And FileExists(fn) Then
    If MsgBox("File exists. Replace?", vbQuestion + vbYesNo) = vbNo Then
      FileDialog = False
      Exit Function
    End If
  End If
  
  If ftype = FD_TYPE_USER Then UI.FileName = d.FileTitle
  
  If ftype = FD_TYPE_LIB Then
    fname = d.FileName 'save whole path
  Else
    fname = d.FileTitle
  End If
  FileDialog = True
  Exit Function

ErrHandlerFND:
'User pressed cancel button
  FileDialog = False
  Exit Function

End Function

Function FileExists(fn)
'return true if the given file exists, false otherwise
  On Error GoTo ErrHandlerFE  'start error trapping
  X = GetAttr(fn) 'if this fails, the file doesn't exist
  FileExists = True
  Exit Function
  
ErrHandlerFE:
  FileExists = False
  Exit Function

End Function

Public Function GenRegNozDist(xNZ As NozzleData, SemiSpan As Single, _
  sNozzles As String, sExtent As String, sSpacing As String) As Boolean
'generate a regularly-spaced nozzle distribution
'
'Two of the three parameters (nozzles, extent, spacing) must be provided.
'
'When the input parameters cannot be implemented exactly, Spacing is given
'priority over Extent or Number of Nozzles.
'
'The nozzle distribution in xNZ is replaced by the new one.
'If the routine is unsuccessful, it returns False and the nozzle distribution
'is unchanged.

  Dim Msg As String
  Dim inflag As Integer
  Dim limit As Single
  Dim nnoz As Integer
  Dim spacing As Single
  Dim tmpspc As Single
  
  Dim hmax As Single
  Dim i As Integer
  Dim n As Integer
  Dim dh As Single
  
  GenRegNozDist = False 'default value
  
  'extract parameter values
  inflag = 0
  'Extent
  If Len(Trim$(sExtent)) > 0 Then
    inflag = inflag + 1  'Bit 1: Extent
    limit = SemiSpan * (Abs(Val(sExtent)) / 100)
  End If
  'Number of Nozzles
  If Len(Trim$(sNozzles)) > 0 Then
    inflag = inflag + 2  'Bit 2: Number
    nnoz = Abs(Val(sNozzles))
    If nnoz = 0 Then
      Msg = "Nozzles must be at least 1"
      MsgBox Msg, vbInformation + vbOKOnly
      nnoz = 1
    ElseIf nnoz > MAX_NOZZLES Then
      Msg = "Nozzles may not be greater than " & Format$(MAX_NOZZLES)
      MsgBox Msg, vbInformation + vbOKOnly
      nnoz = MAX_NOZZLES
    End If
  End If
  'Spacing
  If Len(Trim$(sSpacing)) > 0 Then
    inflag = inflag + 4 'Bit 3: Spacing
    spacing = UnitsInternal(Abs(Val(sSpacing)), UN_LENGTH)
  End If
    
  'set up for the generation
  'The first 3 bits of inflag show which parameters were supplied.
  Select Case inflag
  Case 3 '(011) limit, number: need spacing
    If nnoz > 1 Then
      spacing = 2 * limit / (nnoz - 1)
    Else
      spacing = 0
    End If
  Case 5 '(101) limit, spacing: need number
    If spacing = 0 Or limit = 0 Then
      MsgBox "Enter positive Extent and Spacing.", vbCritical + vbOKOnly
      Exit Function
    End If
    'Find the number of nozzles we need
    nnoz = (2 * limit / spacing) + 1
    'don't exceed the max number of nozzles
    If nnoz > MAX_NOZZLES Then
      Msg = "This distribution requires more than " + _
            Format$(MAX_NOZZLES) + " nozzles. The Extent " + _
            "will be reduced to match this limit."
      MsgBox Msg, vbInformation + vbOKOnly
      nnoz = MAX_NOZZLES
    End If
    'adjust limit to exactly match spacing
    limit = (nnoz - 1) * spacing * 0.5
  Case 6 '(110) number, spacing: need limit
    limit = (nnoz - 1) * spacing * 0.5
  Case Else 'other than exactly two parameters were specified
    Msg = "Two values are needed to generate a regular distribution."
    MsgBox Msg, vbCritical + vbOKOnly
    Exit Function
  End Select

  'generate the distribution
  For i = 0 To nnoz - 1
    xNZ.PosHoriz(i) = limit - spacing * (nnoz - i - 1)
    If Abs(xNZ.PosHoriz(i)) < 0.001 Then xNZ.PosHoriz(i) = 0
    xNZ.PosVert(i) = 0
    xNZ.PosFwd(i) = 0
  Next
  xNZ.NumNoz = nnoz
  xNZ.Type = 1 'user-defined

  'Success
  GenRegNozDist = True
End Function

Sub GetBasicDataAC(BasicType As Integer, xAC As AircraftData)
'Get Basic Aircraft data
'
  Dim DB As Database
  Dim RS As Recordset
  Dim s As String

  If Not LibOpenDB(DB) Then Exit Sub
  Set RS = DB.OpenRecordset("BasicAC")
  
  'position the Recordset to the correct record
  For i = 1 To BasicType
    RS.MoveNext
  Next

  With xAC
    .Type = 0
    .BasicType = BasicType
    .Name = GetBasicNameAC(BasicType)
    .LName = Len(.Name)
    If .BasicType = 2 Then
      .WingType = 4 'Helicopter
    Else
      .WingType = 3 'fixed wing
    End If
    .SemiSpan = RS.Fields("SemiSpan")
    .TypSpeed = RS.Fields("TypSpeed")
    .BiplSep = 0
    .Weight = RS.Fields("Weight")
    .PlanArea = RS.Fields("PlanArea")
    .PropRPM = RS.Fields("PropRPM")
    .PropRad = RS.Fields("PropRad")
    .PropEff = 0.8
    .EngVert = RS.Fields("EngVert")
    .EngFwd = RS.Fields("EngFwd")
    .NumEng = RS.Fields("NumEng")
    FieldToArray RS.Fields("EngHoriz"), .EngHoriz()
    .WingVert = RS.Fields("WingVert")
    .BoomVert = RS.Fields("BoomVert")
    .BoomFwd = RS.Fields("BoomFwd")
    .DragCoeff = 0.1
  End With
End Sub

Sub GetBasicDataSGL(BasicType As Integer, np As Integer, X() As Single, dep() As Single)
'Get Basic (Tier 1) Single-Swath Deposition data
'
  Dim DB As Database
  Dim DS As Recordset

  If Not LibOpenDB(DB) Then Exit Sub
  If Not LibOpenRS(DB, "BasicSgl", DS) Then Exit Sub
  
  'position the Recordset to the correct record
  DS.FindFirst "Type = " & Format$(BasicType)
  If DS.NoMatch Then Exit Sub
  
  np = DS.Fields("NumDep")
  FieldToArray DS.Fields("Distance"), X()
  FieldToArray DS.Fields("Deposition"), dep()
End Sub

Sub GetBasicDataDEP(BasicType As Integer, np As Integer, X() As Single, dep() As Single)
'Get Basic Deposition data
'
  Dim DB As Database
  Dim DS As Recordset

  If Not LibOpenDB(DB) Then Exit Sub
  If Not LibOpenRS(DB, "BasicDep", DS) Then Exit Sub
  
  'position the Recordset to the correct record
  DS.FindFirst "Type = " & Format$(BasicType)
  If DS.NoMatch Then Exit Sub
  
  np = DS.Fields("NumDep")
  FieldToArray DS.Fields("Distance"), X()
  FieldToArray DS.Fields("Deposition"), dep()
  
  DS.Close
End Sub


Sub GetBasicDataDSD(BasicType, xDSD As DropSizeDistData)
'Get Basic Drop Distribution data
'
  Dim DB As Database
  Dim DS As Recordset
  Dim s As String
  Dim xyzzy As Single

  If Not LibOpenDB(DB) Then Exit Sub
  If Not LibOpenRS(DB, "BasicDSD", DS) Then Exit Sub
  
  'position the Recordset to the correct record
  DS.FindFirst "Type = " & Format$(BasicType)
  If DS.NoMatch Then Exit Sub
  
  With xDSD
    .Type = 0   'basic
    .BasicType = BasicType
    .Name = GetBasicNameDSD(BasicType)
    .LName = Len(.Name)
    .NumDrop = DS.Fields("NumDrop")
    FieldToArray DS.Fields("Diam"), .Diam()
    FieldToArray DS.Fields("Frac"), .MassFrac()
    .LibSelFlag = 0
  End With
End Sub

Sub GetBasicDataDSDSwathDisp(BasicType, SwathDisp)
'Get Swath displacement information associated with Basic Drop Distribution data
'
'returns: Swath Displacement as a fraction of swath width
'
  Dim DB As Database
  Dim DS As Recordset
  Dim s As String

  If Not LibOpenDB(DB) Then Exit Sub
  If Not LibOpenRS(DB, "BasicDSD", DS) Then Exit Sub
  
  'position the Recordset to the correct record
  DS.FindFirst "Type = " & Format$(BasicType)
  If DS.NoMatch Then Exit Sub
  
  Select Case UD.Smokey
  Case AUD_SDTF
    If AGDRIFTREGULATORY Then
      SwathDisp = DS.Fields("SwathDispAgReg")
    Else
      SwathDisp = DS.Fields("SwathDispAgPub")
    End If
  Case AUD_FS
    SwathDisp = DS.Fields("SwathDispFS")
  End Select
  
  DS.Close
  DB.Close
End Sub

Sub GetBasicDataNZ(BasicType As Integer, xNZ As NozzleData)
'Get Basic Nozzle data
'
  Dim DB As Database
  Dim DS As Recordset
  Dim s As String

  If Not LibOpenDB(DB) Then Exit Sub
  Set DS = DB.OpenRecordset("BasicNZ")
  
  'position the Recordset to the correct record
  For i = 1 To BasicType
    DS.MoveNext
  Next

  With xNZ
    .Type = 0  'Basic
    .BasicType = BasicType
    .Name = GetBasicNameNZ(BasicType)
    .LName = Len(.Name)
    .NumNoz = DS.Fields("NumNoz")
    FieldToArray DS.Fields("PosHoriz"), .PosHoriz()
    .NozType(0) = 0 'DSD 1
    .PosVert(0) = 0
    .PosFwd(0) = 0
    For i = 1 To .NumNoz - 1
      .NozType(i) = 0 'DSD 1
      .PosVert(i) = 0
      .PosFwd(i) = 0
    Next
    .PosHorizLimit = 0
    Select Case BasicType
    Case 0 'Ag Husky
      .BoomWidth = 69.4
    Case 1 'AT502
      .BoomWidth = 75.8
    Case 2 'Wasp
      .BoomWidth = 72.8
    Case 3 'AT401
      .BoomWidth = 76.3
    End Select
  End With
End Sub

Sub GetBasicDataSM(BasicType As Integer, xSM As SprayMaterialData)
'Get Basic Spray Material data
  With xSM
    Select Case BasicType
    Case 0 'Oil
      .SpecGrav = 0.92
      .NonVGrav = 0.92
      .EvapRate = 0
    Case 1 'water
      .SpecGrav = 1
      .NonVGrav = 1
      .EvapRate = 84.76
    End Select
  End With
End Sub

Function GetBasicNameAC(BasicType) As String
'return the name of the Basic Aircraft of
'the specified type
  ReDim Basiclbl(3) As String
  Basiclbl(0) = "Ag Husky"
  Basiclbl(1) = "Air Tractor AT-502"
  Basiclbl(2) = "Wasp Helicopter"
  Basiclbl(3) = "Air Tractor AT-401"
  
  GetBasicNameAC = Basiclbl(BasicType)
End Function

Function GetBasicNameAC2(BasicType) As String
'return the name of the Basic Aircraft Type of
'the specified type
  ReDim Basiclbl(3) As String
  Basiclbl(0) = "Slow Fixed-wing"
  Basiclbl(1) = "Fast Fixed-wing"
  Basiclbl(2) = "Helicopter"
  Basiclbl(3) = "Slow Fixed-wing"
  
  GetBasicNameAC2 = Basiclbl(BasicType)
End Function

Function GetBasicNameAM(BasicType As Integer)
'return the name of the Application Method
  ReDim Basiclbl(2) As String
  Basiclbl(0) = "Aerial"
  Basiclbl(1) = "Ground"
  Basiclbl(2) = "Orchard Airblast"
  
  GetBasicNameAM = Basiclbl(BasicType)
End Function

Function GetBasicNameDSD(BasicType) As String
'return the name of the Basic Drop Size Distribution of
'the specified type
  Dim DB As Database
  Dim DS As Recordset
  Dim crit As String
  
  If Not LibOpenDB(DB) Then GetBasicNameDSD = "": Exit Function
  
  Set DS = DB.OpenRecordset("BasicDSD", dbOpenDynaset)
  
  'position the Recordset to the correct record
  crit = "Type = " & Format$(BasicType)
  DS.FindFirst crit
  If DS.NoMatch Then GetBasicNameDSD = "": Exit Function
  GetBasicNameDSD = Trim$(DS.Fields("Name"))
  DS.Close
  DB.Close
End Function

Function GetBasicNameGA(BasicType As Integer)
'return the name of the Basic Ground type
  
  GetBasicNameGA = ""
  
  'bit 2: Boom
  If (BasicType And 2) = 0 Then
    GetBasicNameGA = GetBasicNameGA + "Low Boom"
  Else
    GetBasicNameGA = GetBasicNameGA + "High Boom"
  End If
  'bit 1: DSD
  If (BasicType And 1) = 0 Then
    GetBasicNameGA = GetBasicNameGA + ", ASAE Very Fine to Fine"
  Else
    GetBasicNameGA = GetBasicNameGA + ", ASAE Fine to Medium/Coarse"
  End If
  'bit 3: Percentile
  If (BasicType And 4) = 0 Then
    GetBasicNameGA = GetBasicNameGA + ", 50th Percentile"
  Else
    GetBasicNameGA = GetBasicNameGA + ", 90th Percentile"
  End If
  
End Function

Function GetBasicNameNZ(BasicType) As String
'return the name of the Basic Nozzle Distribution of
'the specified type
  ReDim Basiclbl(3) As String
  Basiclbl(0) = "Nozzles for Ag Husky"
  Basiclbl(1) = "Nozzles for Air Tractor AT-502"
  Basiclbl(2) = "Nozzles for Wasp Helicopter"
  Basiclbl(3) = "Nozzles for Air Tractor AT-401"
  
  GetBasicNameNZ = Basiclbl(BasicType)
End Function

Function GetBasicNameOA(BasicType As Integer)
'return the name of the Basic Orchard Airblast type
  Dim sInfo As String
  
  GetBasicNameOA = "" 'Default
  sInfo = GetTier1Info(AM_ORCHARD, BasicType) 'Get the whole info string
  'The first line if the Info is the name
  If InStr(sInfo, vbCrLf) > 0 Then
    GetBasicNameOA = Left(sInfo, InStr(sInfo, vbCrLf) - 1)
  End If
End Function

Function GetBasicNameSM(BasicType)
'return the name of the Basic Spray Material
  ReDim Basiclbl(1) As String
  Basiclbl(0) = "Oil"
  Basiclbl(1) = "Water"
  
  GetBasicNameSM = Basiclbl(BasicType)
End Function

Function GetBasicNameT1(xUD As UserData)
'Return the "basic" name of Tier I data
  GetBasicNameT1 = "" 'default value
  Select Case xUD.Tier
  Case TIER_1
    Select Case xUD.ApplMethod
    Case AM_AERIAL
      GetBasicNameT1 = GetBasicNameDSD(xUD.DSD(0).BasicType)
    Case AM_GROUND
      GetBasicNameT1 = GetBasicNameGA(xUD.GA.BasicType)
    Case AM_ORCHARD
      GetBasicNameT1 = GetBasicNameOA(xUD.OA.BasicType)
    End Select
  End Select
End Function

Function GetARSNozName(Index) As String
'return the name of the ARS Nozzle
  Dim DB As Database
  Dim RS As Recordset
  
  GetARSNozName = "" 'default return value
  
  If Not LibOpenDB(DB) Then Exit Function
  If Not LibOpenRS(DB, "ARSNozzle", RS) Then Exit Function
  
  'position the Recordset to the correct record
  RS.FindFirst "Index=" & Format$(Index)
  If RS.NoMatch Then Exit Function
  GetARSNozName = Trim$(RS.Fields("Name"))
  RS.Close
  DB.Close
End Function

Sub GetARSNozData(Index, _
                  OfcUnits As Integer, OfcLabel As String, OfcPrefix As String, _
                  NumOfc As Integer, OfcVal() As Single, _
                  ModUnits As Integer, ModLabel As String, ModPrefix As String, _
                  NumMod As Integer, ModVal() As Single)
'Get ARS Nozzle Data from the library
  Dim DB As Database
  Dim RS As Recordset
  Dim s As String
  
  If Not LibOpenDB(DB) Then Exit Sub
  If Not LibOpenRS(DB, "ARSNozzle", RS) Then Exit Sub
  
  'position the Recordset to the correct record
  RS.FindFirst "Index = " & Format$(Index)
  If RS.NoMatch Then Exit Sub
  
  OfcUnits = RS.Fields("OfcUnits")
  OfcLabel = Trim$(RS.Fields("OfcLabel"))
  OfcPrefix = Trim$(RS.Fields("OfcPrefix"))
  NumOfc = RS.Fields("NumOfc")
  FieldToArray RS.Fields("OfcVal"), OfcVal()
  ModUnits = RS.Fields("ModUnits")
  ModLabel = Trim$(RS.Fields("ModLabel"))
  ModPrefix = Trim$(RS.Fields("ModPrefix"))
  NumMod = RS.Fields("NumMod")
  FieldToArray RS.Fields("ModVal"), ModVal()
  
  RS.Close
  DB.Close
End Sub

Sub GetHelp(Topic As Integer)
' Displays the selected help topic selected from either
' Editors;' or Viewer's help menu.
'
    Dim R As Integer

    If HelpTopic = -1 Then
        '
        ' "Using Help" was selected so display the Standard Windows Help
        ' Topic for "Using Help".
        '
        R = WinHelp(frmMain.hwnd, dummy$, cdlHelpHelpOnHelp, 0)
    Else
        ' A help topic other the "Using help" was selected.
        '
        R = WinHelp(frmMain.hwnd, App.HelpFile, cdlHelpContext, ByVal CLng(Topic))
    End If

End Sub

Sub GetLibDataDSD(key As String, nd As Integer, Diam() As Single, frac() As Single)
'Get Library Drop Distribution data
'
' key  i  a "master" search key comprised of a set of 6
'         subkeys separated by commas. This is exactly what
'         frmDropLib returns.
'
' xDSD o  a fully formatted Drop Size Distribution retrieved
'         from the library.
'
  Dim DB As Database
  Dim DS As Recordset
  Dim crit As String
  Dim s As String
  Dim keyLib As String
  Dim keySubst As String
  Dim keyNoz As String
  Dim keyNozAng As String
  Dim keyNozRPM As String
  Dim keyPressure As String
  Dim keyWS As String
  Dim ndlong As Long
  Dim dslflag As Integer

  If Trim$(key) = "" Then Exit Sub

  'Extract search keys from the master key
  start = 1
  comma = InStr(start, key, ",")
  keylen = comma - start
  keyLib = Mid$(key, start, keylen)
  start = comma + 1
  comma = InStr(start, key, ",")
  keylen = comma - start
  keySubst = Mid$(key, start, keylen)
  start = comma + 1
  comma = InStr(start, key, ",")
  keylen = comma - start
  keyNoz = Mid$(key, start, keylen)
  start = comma + 1
  comma = InStr(start, key, ",")
  keylen = comma - start
  keyNozAng = Mid$(key, start, keylen)
  start = comma + 1
  comma = InStr(start, key, ",")
  keylen = comma - start
  keyNozRPM = Mid$(key, start, keylen)
  start = comma + 1
  comma = InStr(start, key, ",")
  keylen = comma - start
  keyPressure = Mid$(key, start, keylen)
  start = comma + 1
  keyWS = Mid$(key, start)

  If Not LibOpenDB(DB) Then Exit Sub
  If keyLib = "0" Then
    Set DS = DB.OpenRecordset("Dropsize", dbOpenDynaset)
  Else
    Set DS = DB.OpenRecordset("DropsizeFS", dbOpenDynaset)
  End If
  
  'position the Recordset to the correct record
  crit = "Substance = '" & keySubst & "'"
  crit = crit & " and Nozzle = '" & keyNoz & "'"
  crit = crit & " and NozzleAngle = " & keyNozAng
  crit = crit & " and NozzleRPM = " & keyNozRPM
  crit = crit & " and Pressure = " & keyPressure
  crit = crit & " and WindSpeed = " & keyWS
  DS.FindFirst crit
  If DS.NoMatch Then
    nd = 0
    Exit Sub
  End If

  'Retrieve the data from the library
  nd = 32 'all lib entries have 32 categories
  FieldToArray DS.Fields("MassFrac"), frac()
  dslflag = DS.Fields("DSLflag") 'flag used to get diams below
  DS.Close

  'retrieve the diameter data from a different table
  If keyLib = "0" Then
    Set DS = DB.OpenRecordset("Dropdiam", dbOpenDynaset)
  Else
    Set DS = DB.OpenRecordset("DropdiamFS", dbOpenDynaset)
  End If
  DS.FindFirst "DSLflag = " & CStr(dslflag)
  If DS.NoMatch Then
    nd = 0
    Exit Sub
  End If
  FieldToArray DS.Fields("Diameter"), Diam()
  DS.Close
  DB.Close
End Sub

Function GetRunID() As String
'return a runid associated with the current data set.
'the ID is a chracter string consisting of:
' - the file name
' - the version of AgDRIFT that performed the calculations
' - the date and time when the calculations were started
  Dim s As String
  s = App.Title + "® " 'application name
  If UD.Tier = TIER_1 Then
    s = s & " " & GetVersionString(AGDRIFTVERSION)
  Else
    s = s + UI.FileName
    If UC.Valid Then
      s = s & " " & GetVersionString(UC.CodeVersion)
      s = s & " " & UC.StartDate
      s = s & " " & UC.StartTime
    Else
      s = s & " " & GetVersionString(AGDRIFTVERSION)
      s = s & " 00-00-0000"
      s = s & " 00:00:00"
    End If
  End If
  GetRunID = s
End Function

Sub GetTier1Calcs(ApplMethod As Integer, BasicType As Integer, _
                  Param1 As Integer, Param2 As Integer, _
                  xUC As UserCalc)
'Get the Tier1 calculations from storage and load into
'a UserCalc Structure.
' ApplMethod  i  Application method 0=aerial 1=ground 2=orchard
' BasicType   i  Basic type (if ApplMethod = 0 use UD.DSD().BasicType)
'                           (if ApplMethod = 1 use UD.GA.BasicType)
'                           (if ApplMethod = 2 use UD.OA.BasicType)
' Param1, Param2 i  AppMethod Param 1            Param 2
'                   ========= ================== ===============
'                   aerial    not used           not used
'                   ground    Number of swaths   not used
'                   orchard   Beginning Tree Row Ending Tree Row
'
' xUC         o  calculation data to be filled
'
' NOTE: The current Tier affects the data that is returned. If
'       you GetTier1Calcs for Tier 1, you'll get less data than
'       if you make the same call under Tier 3.
'
  Dim s As String
  Dim NXtmp As Long
  Dim NStmp As Long
  Dim BegTrow As Long
  Dim EndTrow As Long
  Dim Sdum As Single
  Dim NXY As Integer 'actual number of points in the following arrays
  Dim X(MAX_CALCDATA - 1) As Single 'temp array for calc data
  Dim Y(MAX_CALCDATA - 1) As Single 'temp array for calc data

  Dim Tier As Long
  Dim MaxDownwindDist As Single 'For Tier 3, aglibr requires this param
  
  'Get the current tier. If you GetTier1Calcs for Tier 1, you'll get
  'less data than if you make the same call under Tier 3.
  Tier = CLng(UD.Tier)
  MaxDownwindDist = UD.CTL.MaxDownwindDist
  
  Select Case ApplMethod
  Case AM_AERIAL
    'Deposition
    'get the raw distribution from the lib
    GetBasicDataDEP BasicType, NXY, X(), Y()
    'clean up the distribution
    NXtmp = CLng(NXY)
    Call aglibr(Tier, 0, MaxDownwindDist, NXtmp, X(0), Y(0))
    NXY = CInt(NXtmp)
    'transfer to final storage
    xUC.NumDep = CInt(NXY)
    ReDim xUC.DepDist(NXY)
    ReDim xUC.DepVal(NXY)
    CopyMemory xUC.DepDist(0), X(0), Len(X(0)) * NXY
    CopyMemory xUC.DepVal(0), Y(0), Len(Y(0)) * NXY
    xUC.DepExtrap = 0 'No extrapolation for aerial
      
    'Pond-Integrated Deposition
    'get the raw distribution from the lib
    GetBasicDataDEP BasicType, NXY, X(), Y()
    'clean up the distribution
    NXtmp = CLng(NXY)
    Call aglibr(Tier, 1, MaxDownwindDist, NXtmp, X(0), Y(0))
    NXY = CInt(NXtmp)
    'transfer to final storage
    xUC.NumPID = CInt(NXY)
    ReDim xUC.PIDDist(NXY)
    ReDim xUC.PIDVal(NXY)
    CopyMemory xUC.PIDDist(0), X(0), Len(X(0)) * NXY
    CopyMemory xUC.PIDVal(0), Y(0), Len(Y(0)) * NXY
    xUC.PIDExtrap = 0 'No extrapolation for aerial
      
    'Vertical Deposition
    xUC.NumFlux = 0                   'no vert depos
    ReDim xUC.FluxDist(0)
    ReDim xUC.FluxVal(0)
  
    'Single-swath deposition
    GetBasicDataSGL BasicType, NXY, X(), Y()
    xUC.NumSgl = NXY
    ReDim xUC.SglDist(NXY)
    ReDim xUC.SglVal(NXY)
    CopyMemory xUC.SglDist(0), X(0), Len(X(0)) * NXY
    CopyMemory xUC.SglVal(0), Y(0), Len(Y(0)) * NXY
  
  Case AM_GROUND
    'deposition
    NStmp = CLng(Param1) 'number of swaths
    Call aggrnd(CLng(BasicType), Tier, MaxDownwindDist, NStmp, AGGRND_DEPOS, NXtmp, X(0), Y(0))
    NXY = CInt(NXtmp)
    Param1 = CInt(NStmp) 'NStmp may have been corrected
    xUC.NumDep = NXY
    ReDim xUC.DepDist(NXY)
    ReDim xUC.DepVal(NXY)
    CopyMemory xUC.DepDist(0), X(0), Len(X(0)) * NXY
    CopyMemory xUC.DepVal(0), Y(0), Len(Y(0)) * NXY
    
    'Pond-Integrated Deposition
    NStmp = CLng(Param1) 'number of swaths
    Call aggrnd(CLng(BasicType), Tier, MaxDownwindDist, NStmp, AGGRND_PID, NXtmp, X(0), Y(0))
    NXY = CInt(NXtmp)
    Param1 = CInt(NStmp) 'NStmp may have been corrected
    xUC.NumPID = NXY
    ReDim xUC.PIDDist(NXY)
    ReDim xUC.PIDVal(NXY)
    CopyMemory xUC.PIDDist(0), X(0), Len(X(0)) * NXY
    CopyMemory xUC.PIDVal(0), Y(0), Len(Y(0)) * NXY
  
    'Vertical Deposition
    xUC.NumFlux = 0                   'no vert depos
    ReDim xUC.FluxDist(0)
    ReDim xUC.FluxVal(0)
  
    'Single-swath Deposition
    'call aggrnd with:
    '  Tier set to 3
    '  Dep/PID flag set to 0 (Dep)
    '  NumSwaths set to 1
    '  MaxDist set to 748m
    Tier = 3
    NStmp = 1
    MaxDownwindDist = 748
    Call aggrnd(CLng(BasicType), Tier, MaxDownwindDist, NStmp, AGGRND_DEPOS, NXtmp, X(0), Y(0))
    NXY = CInt(NXtmp)
    xUC.NumSgl = NXY
    ReDim xUC.SglDist(NXY)
    ReDim xUC.SglVal(NXY)
    CopyMemory xUC.SglDist(0), X(0), Len(X(0)) * NXY
    CopyMemory xUC.SglVal(0), Y(0), Len(Y(0)) * NXY
  
  Case AM_ORCHARD
    'deposition
    BegTrow = CLng(Param1)
    EndTrow = CLng(Param2)
    Call agorch(CLng(BasicType), Tier, MaxDownwindDist, BegTrow, EndTrow, AGGRND_DEPOS, NXtmp, X(0), Y(0))
    NXY = CInt(NXtmp)
    xUC.NumDep = NXY
    ReDim xUC.DepDist(NXY)
    ReDim xUC.DepVal(NXY)
    CopyMemory xUC.DepDist(0), X(0), Len(X(0)) * NXY
    CopyMemory xUC.DepVal(0), Y(0), Len(Y(0)) * NXY
    
    'Pond-Integrated Deposition
    BegTrow = CLng(Param1)
    EndTrow = CLng(Param2)
    Call agorch(CLng(BasicType), Tier, MaxDownwindDist, BegTrow, EndTrow, AGGRND_PID, NXtmp, X(0), Y(0))
    xUC.NumPID = NXY
    ReDim xUC.PIDDist(NXY)
    ReDim xUC.PIDVal(NXY)
    CopyMemory xUC.PIDDist(0), X(0), Len(X(0)) * NXY
    CopyMemory xUC.PIDVal(0), Y(0), Len(Y(0)) * NXY
  
    'Vertical Deposition
    xUC.NumFlux = 0                   'no vert depos
    ReDim xUC.FluxDist(0)
    ReDim xUC.FluxVal(0)
  
    'Single-swath Deposition
    xUC.NumSgl = 0                    'no single-swath
    ReDim xUC.SglDist(0)
    ReDim xUC.SglVal(0)
  End Select
End Sub

Public Function GetTier1Info(ApplMethod As Integer, BasicType As Integer) As String
'Return a descriptive string relating to Tier1 Ground or Orchard
  Dim DB As Database
  Dim DS As Recordset

  If Not LibOpenDB(DB) Then Exit Function
  Set DS = DB.OpenRecordset("Select Info from BasicInfo where" + _
                            " ApplMeth=" + Format$(ApplMethod) + _
                            " and Type=" + Format$(BasicType), _
                            dbOpenSnapshot)

  If DS.BOF And DS.EOF Then
    GetTier1Info = ""
  Else
    GetTier1Info = DS.Fields("Info")
  End If
  DS.Close
  DB.Close
End Function

Function GetTypeNameAC(INTYPE As Integer)
'return the name of the Aircraft type
  ReDim Typelbl(3) As String
  Typelbl(0) = "Basic"
  Typelbl(1) = "User-defined"
  Typelbl(2) = "Library" 'SDTF
  Typelbl(3) = "Library" 'FS
  
  GetTypeNameAC = Typelbl(INTYPE)
End Function

Function GetTypeNameACWing(INTYPE As Integer)
'return the name of the Aircraft WingType
  ReDim Typelbl(3 To 4) As String
  Typelbl(3) = "Fixed-Wing"
  Typelbl(4) = "Helicopter"
  
  GetTypeNameACWing = Typelbl(INTYPE)
End Function

Function GetTypeNameDKCalc(INTYPE As Integer)
'return the name of the DropKick Calculation type
  ReDim Typelbl(2) As String
  Typelbl(0) = "Root Normal"
  Typelbl(1) = "Rosin-Rammler"
  Typelbl(2) = "Log Normal"
  
  GetTypeNameDKCalc = Typelbl(INTYPE)
End Function

Function GetTypeNameDKMat(INTYPE As Integer)
'return the name of the DropKick Material type
  ReDim Typelbl(1) As String
  Typelbl(0) = "User-defined"
  Typelbl(1) = "Library"
  
  GetTypeNameDKMat = Typelbl(INTYPE)
End Function

Function GetTypeNameDKNoz(INTYPE As Integer)
'return the name of the DropKick Nozzle type
  ReDim Typelbl(1) As String
  Typelbl(0) = "User-defined"
  Typelbl(1) = "Library"
  
  GetTypeNameDKNoz = Typelbl(INTYPE)
End Function

Function GetNameHKMat(INTYPE As Integer)
'return the name of the Rotary Nozzle Spray Material
  Dim s As String
  Select Case INTYPE
  Case 0: s = "Water"
  Case 1: s = "Water with 1% w/w Sta-Put®"
  Case 2: s = "Water with 0.25% w/w Hasten®"
  Case 3: s = "Foray 76B neat"
  Case Else: s = ""
  End Select
  GetNameHKMat = s
End Function

Function GetTypeNameStability(INTYPE As Integer)
  Dim s As String
  Select Case INTYPE
  Case 0: s = "Strong"
  Case 1: s = "Moderate"
  Case 2: s = "Slight"
  Case 3: s = "Weak"
  Case 4: s = "Overcast"
  Case 5: s = "Thinly Overcast"
  Case 6: s = "Less Than 3/8ths Overcast"
  Case Else: s = ""
  End Select
  GetTypeNameStability = s
End Function


Function GetNameHKNoz(INTYPE As Integer)
'return the name of the Rotary Nozzle
  Dim s As String
  Select Case INTYPE
  Case 0: s = "Micronair AU4000"
  Case 1: s = "Micronair AU5000"
  Case Else: s = ""
  End Select
  GetNameHKNoz = s
End Function

Function GetTypeNameDSD(INTYPE As Integer)
'return the name of the Drop Size Distribution type
  ReDim Typelbl(6) As String
  Typelbl(0) = "Basic"
  Typelbl(1) = "DropKick"
  Typelbl(2) = "User-defined"
  Typelbl(3) = "Library" 'SDTF
  Typelbl(4) = "Library" 'FS
  Typelbl(5) = "USDA ARS"
  'Typelbl(6) = "User Library"
  
  GetTypeNameDSD = Typelbl(INTYPE)
End Function

Function GetTypeNameCN(INTYPE As Integer)
'return the name of the Canopy type
  Select Case INTYPE
  Case 0: GetTypeNameCN = "None"
  Case 1: GetTypeNameCN = "Story"
  Case 2: GetTypeNameCN = "Optical"
  Case 3: GetTypeNameCN = "Basic"
  Case Else
    GetTypeNameCN = ""
  End Select
End Function

Function GetTypeNameOP(INTYPE As Integer)
'return the name of the Optical Canopy type
  Select Case INTYPE
  Case 1: GetTypeNameOP = "User-Defined"
  Case 2: GetTypeNameOP = "Library"
  Case Else
    GetTypeNameOP = ""
  End Select
End Function

Function GetTypeNameNZ(INTYPE As Integer)
'return the name of the Nozzle type
  Select Case INTYPE
  Case 0
    GetTypeNameNZ = "Basic"
  Case 1
    GetTypeNameNZ = "User-Defined"
  Case 2
    GetTypeNameNZ = "Library"
  Case Else
    GetTypeNameNZ = ""
  End Select
End Function

Function GetTypeNameSM(INTYPE As Integer)
'return the name of the Spray Material type
  ReDim Typelbl(2) As String
  Typelbl(0) = "Basic"
  Typelbl(1) = "User-defined"
  Typelbl(2) = "Library"
  
  GetTypeNameSM = Typelbl(INTYPE)
End Function

Sub InitInterface(xUI As InterfaceData)
'Init User Interface data items
  xUI.HasConfidentialData = False
  xUI.FileName = ""
  xUI.LibraryPath = LibGetFileName()
  xUI.MAALibraryPath = App.Path & "\" & "agdmaa.mdb"
  xUI.DataChanged = False
  xUI.CalcsBatchMode = False
  xUI.OkToDoCalcs = True
  xUI.DataNeedsChecking = True
  xUI.CalcsInProgress = False
  xUI.StartCalcsOnLoad = False
  xUI.RevertCalcsAvailable = False
  xUI.PlotVar = 0
End Sub

Function LineFromString(s, start) As String
'returns the next line, delineated by CRLF, from the
'given string, without the CRLF. on return, start is set
'to a position after the next CRLF
  If start > Len(s) Then
    start = 0
    LineFromString = ""
    Exit Function
  End If
  idx = InStr(start, s, vbCrLf) 'find next crlf after the specified start
  If idx = 0 Then
    slen = Len(s) - start + 1
    LineFromString = Mid$(s, start, slen)
    start = 0
  Else
    slen = idx - start
    LineFromString = Mid$(s, start, slen)
    start = idx + Len(vbCrLf)
  End If
End Function

Function LinesInString(s)
'count lines in given string by counting the number
'of carriage return/linefeeds within the string
  lines = 0
  start = 1
  Do
    idx = InStr(start, s, vbCrLf)
    lines = lines + 1
    start = idx + Len(vbCrLf)
  Loop While idx > 0
  LinesInString = lines
End Function

Sub Main()
'This is the first routine executed when Agdrift starts up
  Set gobjErrors = New clsErrors 'Init error handling
  App.HelpFile = App.Path & "\" & "agdrift.hlp" 'define the help file
  InitInterface UI     'Init User Inteface data
  ReadGeneralPrefs     'Initialize User Preferences
  UnitsSelectSystem UP.Units 'set units system
  UD.Smokey = UP.InitialAUD 'Set up "mode" or "audience"
  UserDataDefault UD   'set user data to default values
  ClearUserCalc UC     'reset calculation data
  PlotSettingsInit PS  'Init Plot setting data
  PlotPrefsRead PS     'Override Plot Settings with User Prefs
  BackupPlotDataSources 'Init other plot data
  frmMain.Show         'Show the main form
  frmAbout.Show vbModal
End Sub

Function NewFile() As Integer
'Reset user data to default state

  NewFile = False 'default status
  'save changed data first
  If Not QuerySaveChanged() Then Exit Function
    
  UserDataDefault UD     'load default data into user data
  ClearUserCalc UC        'wipe out existing calcs
  UI.FileName = ""       'wipe out any existing file name
  UpdateDataChangedFlag False 'reset the changed flag
  NewFile = True 'Success
End Function

Function OpenFile() As Integer
'Open a user file and read its contents
  
  OpenFile = False 'default status
  'save changed data first
  If Not QuerySaveChanged() Then Exit Function
    
  'open the file dialog and get a file name
  If FileDialog(FD_OPEN, FD_TYPE_USER, UI.FileName) Then 'if the selection is good
    OpenFile = UserDataRead(UI.FileName, UD, UC, True)
    UpdateDataChangedFlag False
  End If
End Function

Function PerformCalcs(Pause As Integer) As Integer
'load the calcuations form and do the calcs on the current data
  'If Pause is true, the user must press start to continue
  UI.StartCalcsOnLoad = Not Pause
  UI.CalcsBatchMode = False 'set form mode flag
  frmCalc.Show vbModal
  Select Case frmCalc.Tag
  Case "ok"
    PerformCalcs = True
  Case Else
    PerformCalcs = False
  End Select
  Unload frmCalc
End Function

Sub PrintData(ReportText As String, pvw As Integer, pages, Mag)
'print the current UserData
' This routine handles printing to the printer object, and
' to an array of picture boxes (for print=preview).
'
'  ReportText i A String containing the raw text to be
'               printed.
'  pvw        i   the destination of the output
'                   if true, this is a preview
'                   if false, this is a print job
'  pages      o   the number of nonblank pages produced
'  Mag        i   the desired magnification for
'                 the picture box. Ignored if pvw is false.
'
  Dim s As String
  Dim f As Form
  Dim stmp As String
  Dim PVfontname As String
  Dim PVfontsize As Single
  Dim LineHeight As Single
  Dim LinesPerPage As Integer
  Dim lines As Integer
  Dim MarginTop As Single
  Dim MarginLeft As Single
  Dim MarginBottom As Single
  Dim MarginRight As Single

  On Error GoTo ErrHandPrintData
  
  'set the margins
  MarginTop = 1
  MarginBottom = 1
  MarginLeft = 1
  MarginRight = 1

  'Set the output font
  PVfontname = "Courier New"
  PVfontsize = 10
  LineHeight = PVfontsize / 72# 'Set line height in inches

  'Printer object general setup
  Printer.FontName = PVfontname     'select printer font
  Printer.FontSize = PVfontsize     'select printer fontsize
  Printer.FontBold = False
  Printer.FontItalic = False
  Printer.FontStrikethru = False
  Printer.FontTransparent = False
  Printer.FontUnderline = False
  Printer.ScaleMode = 5             'set printer scaling to inches
  PageHeight = Printer.ScaleHeight - MarginTop - MarginBottom
  PageWidth = Printer.ScaleWidth - MarginLeft - MarginRight
  'scale the output page to obtain the top and left margins
  'the Printer scale height and width are already set
  Printer.ScaleLeft = -MarginLeft
  Printer.ScaleTop = -MarginTop
  'Set the starting Position
  Printer.CurrentX = 0
  Printer.CurrentY = 0

  'preview form general setup
  If pvw Then
    pages = 1
    frmPrintPreview.picPage(1).FontName = PVfontname
    frmPrintPreview.picPage(1).FontSize = PVfontsize * Mag
    frmPrintPreview.picPage(1).FontBold = False
    frmPrintPreview.picPage(1).FontItalic = False
    frmPrintPreview.picPage(1).FontStrikethru = False
    frmPrintPreview.picPage(1).FontTransparent = False
    frmPrintPreview.picPage(1).FontUnderline = False
    'set the size of the pictureboxes
    frmPrintPreview.picPage(1).Height = Printer.Height * Mag
    frmPrintPreview.picPage(1).Width = Printer.Width * Mag
    'scale the output page
    frmPrintPreview.picPage(1).ScaleMode = 0       'set form scaling to user
    frmPrintPreview.picPage(1).ScaleLeft = Printer.ScaleLeft
    frmPrintPreview.picPage(1).ScaleWidth = Printer.ScaleWidth
    frmPrintPreview.picPage(1).ScaleTop = Printer.ScaleTop
    frmPrintPreview.picPage(1).ScaleHeight = Printer.ScaleHeight
    frmPrintPreview.picPage(1).CurrentX = 0
    frmPrintPreview.picPage(1).CurrentY = 0
  End If
  
  'Initialize Line counter and limit
  LinesPerPage = PageHeight / LineHeight
  lines = 0

  'print page header
  GoSub printHeader
  
  'Format the report text
  start = 1
  Do
    s = LineFromString(ReportText, start)
    GoSub PrintString
  Loop While start > 0
  
  If Not pvw Then
    Printer.EndDoc
  End If
PrintDataExit:
  Exit Sub

'subroutines -----------------------------------------------------

PrintUnderlineOn:
  If pvw Then
    frmPrintPreview.picPage(pages).FontUnderline = True
  Else
    Printer.FontUnderline = True
  End If
  Return

PrintUnderlineOff:
  If pvw Then
    frmPrintPreview.picPage(pages).FontUnderline = False
  Else
    Printer.FontUnderline = False
  End If
  Return

printHeader:
'subroutine for printing page header
  If pvw Then
    stmp = " " & GetRunID()
    frmPrintPreview.picPage(pages).Line (0, 0)-(PageWidth, 1.5 * LineHeight), , B
    frmPrintPreview.picPage(pages).CurrentX = 0
    frmPrintPreview.picPage(pages).CurrentY = 0.25 * LineHeight
    frmPrintPreview.picPage(pages).Print stmp;
    stmp = "Page " + Format$(pages) + " "
    frmPrintPreview.picPage(pages).CurrentX = PageWidth - Printer.TextWidth(stmp)
    frmPrintPreview.picPage(pages).CurrentY = 0.25 * LineHeight
    frmPrintPreview.picPage(pages).Print stmp;
    lines = 4
  Else
    stmp = " " & GetRunID()
    Printer.Line (0, 0)-(PageWidth, 1.5 * LineHeight), , B
    Printer.CurrentX = 0
    Printer.CurrentY = 0.25 * LineHeight
    Printer.Print stmp;
    stmp = "Page " + Format$(Printer.Page) + " "
    Printer.CurrentX = PageWidth - Printer.TextWidth(stmp)
    Printer.CurrentY = 0.25 * LineHeight
    Printer.Print stmp;
    lines = 4
  End If
  Return

'subroutine for printing line passed by string "s"
PrintString:
  'see if there's room on the current page
  If lines >= LinesPerPage Then
    GoSub PrintNewPage
    GoSub printHeader
  End If
  If pvw Then
    frmPrintPreview.picPage(pages).CurrentX = 0
    frmPrintPreview.picPage(pages).CurrentY = lines * LineHeight
    frmPrintPreview.picPage(pages).Print s;
  Else
    Printer.CurrentX = 0
    Printer.CurrentY = lines * LineHeight
    Printer.Print s;
  End If
  lines = lines + 1
  Return

PrintNewPage:
'prepare a new page for printing
  If pvw Then
    'if this is a preview, load a new picturebox
    pages = pages + 1
    Load frmPrintPreview.picPage(pages)
  Else
    'if we're actually printing, output the current page
    Printer.NewPage
  End If
  lines = 0
  Return

ErrHandPrintData:
'error handler
  MsgBox "Error " + Format$(Err.Number) + " during printing: " + Err.Description
  Resume PrintDataExit
  
End Sub

Function PrintDialog(BeginPage As Integer, EndPage As Integer, NumCopies As Integer) As Integer
'bring up a printing dialog box and return some stats
'
'This procedure does no printing. It only gathers
'user selections for printing

  Dim i As Integer
  Dim c As Control

  Const THISFORM = 0
  Const THEPRINTER = 1

  'Set the control
  Set c = frmMain.CMDialog1

  'Set Cancel to True
  c.CancelError = True
  On Error GoTo ErrHandlerPD

  'Set dialog box flags
  c.Flags = cdlPDNoSelection + cdlPDNoPageNums + cdlPDHidePrintToFile

  'Display the Print Dialog box
  c.ShowPrinter

  'Get user-selected values
  BeginPage = c.FromPage
  EndPage = c.ToPage
  NumCopies = c.Copies
  PrintDialog = True
  Exit Function

ErrHandlerPD:
  'User pressed Cancel button
  BeginPage = 0
  EndPage = 0
  NumCopies = 0
  PrintDialog = False
  Exit Function

End Function

Function QuerySaveChanged()
'check the status of DataChanged and give the user
' a chance to save data.
'
'returns:  true  if sucessful
'          false cancelled
'
  Dim ans As Integer
  Dim Msg As String
  Dim MBType As Integer

  QuerySaveChanged = True 'the default
  'give the user a chance to save the current data, if needed
  If UD.Tier > 1 And UI.DataChanged Then
    Msg = "The current data has changed." + Chr$(13)
    Msg = Msg + "Do you want to save the changes?"
    MBType = vbQuestion + vbYesNoCancel
    ans = MsgBox(Msg, MBType)
    '6="Yes" 7="No" 2="Cancel"
    Select Case ans
    Case vbYes  'Yes - go to the save routine
      If Not SaveAsFile() Then QuerySaveChanged = False
    Case vbNo  'No - just continue
    Case vbCancel  'Cancel - exit this routine
      QuerySaveChanged = False
    End Select
  End If
End Function

Sub ReadGeneralPrefs()
'read settings in the .ini file
  Dim fn As String      '.ini file name
  Dim an As String      'name for .ini file section
  Dim i As Long
  Dim s As String * 255

  'build the settings file name from the Application
  fn = App.Path & Chr$(92) & App.EXEName & ".ini"
  
  'retrieve preferences
  an = "Preferences"
  UP.InitialTier = GetPrivateProfileInt(an, "InitialTier", 1, fn)
  UP.InitialAM = GetPrivateProfileInt(an, "InitialAM", 0, fn)
  UP.InitialAUD = GetPrivateProfileInt(an, "InitialAUD", 0, fn)
  UP.WarnOnTierChange = GetPrivateProfileInt(an, "WarnOnTierChange", True, fn)
  UP.PauseBeforeCalc = GetPrivateProfileInt(an, "PauseBeforeCalc", True, fn)
  UP.Units = GetPrivateProfileInt(an, "Units", UN_IMPERIAL, fn)
  UP.SuppressTier3Warn = GetPrivateProfileInt(an, "SuppressTier3Warn", False, fn)
  i = GetPrivateProfileString(an, "UserLib", App.Path & "\userlib.mdb", s, 255, fn)
  UP.UserLib = Left$(s, i)

  'make sure there were no prefs illegal for Public use
  If Not AGDRIFTREGULATORY Then
    'no T1A
    If UP.InitialTier = TIER_1 And UP.InitialAM = AM_AERIAL Then
      UP.InitialAM = AM_GROUND
    End If
    'no T2
    If UP.InitialTier = TIER_2 Then
      UP.InitialTier = TIER_3
    End If
  End If
  
  'make sure there was no old prefs setting for T1/Forestry
  If UP.InitialTier = TIER_1 And UP.InitialAUD = AUD_FS Then
    UP.InitialTier = TIER_2 'switch to tier 2
  End If
End Sub

Function UserDataRead(fn As String, xUD As UserData, xUC As UserCalc, Notify As Integer)
'Read UserData and UserCalc from file
'
'fn     i  the file from which to read
'xUD    o  a UserData structure to receive part of the data
'Notify i Error notification flag:
'           true=notify user of errors through message box
'           false=do not notify
'Return value:
'  true=success
'  false=errors encountered
'
'Currently handles file versions 0 - 4
'Remember: "version" in this routine is an integer
'file format version, not the program version
'
'Perhaps a separate routine for each version would be better?
'That would get cumbersome after a while
'Perhaps separate routines for groups of similar versions
'
  Dim Header As String
  Dim version As Integer
  Dim xNZ As NozzleData
  Dim iDrp As Integer, nDrp As Integer
  Dim dummy As Variant

  'read user data
  On Error GoTo UserDataReadErrHand1
  Open fn For Input As #1
  
  'read header to determine format (e.g. "AGDRIFT_userdata")
  On Error GoTo UserDataReadErrHand2
  Input #1, Header
  
  'check for a non-agdrift file
  If Header <> "AGDRIFT_userdata" Then
    MsgBox "Not an AGDRIFT input file.", vbOKOnly + vbExclamation, "Read data"
    Close #1
    UserDataRead = False
    Exit Function
  End If
    
  'read file version number (not the AgDRIFT program version number!)
  Input #1, version
  
  'file version > current version (old program used with new data)
  If version > USERFILEVERSION Then
    MsgBox "Unknown file format", vbOKOnly + vbExclamation, "Read data"
    Close #1
    UserDataRead = False
    Exit Function
  End If
  
  'file format is OK. Read the rest of the file
  
  'But first, default the data and the calcs
  UserDataDefault xUD
  ClearUserCalc xUC
  
  'UserData
  With xUD
    'General
    Input #1, .Tier
    'Guard against Tier 2 files read into public use version
    If Not AGDRIFTREGULATORY Then
      If .Tier = TIER_2 Then .Tier = TIER_3
    End If
    'in 1.08 the header portion was rearranged
    'slightly and .Smokey was added
    If version >= 2 Then
      Input #1, .ApplMethod
      Input #1, .Smokey
    End If
    Line Input #1, .Title
    'Get the Notes data in multiline format
    Input #1, lines
    .Notes = "" 'start with a blank
    For i = 1 To lines
      Line Input #1, s                   'get next line
      If i < lines Then s = s & vbCrLf 'append CRLF
      .Notes = .Notes & s          'append to the total string
    Next
    'this is the old location of .ApplMethod
    If version < 2 Then
      Input #1, .ApplMethod
      .Smokey = AUD_SDTF
    End If
   
    'Ground Application
    With .GA
      Input #1, .BasicType
      Input #1, .NumSwaths
    End With
  
    'Orchard Airblast Data (added in 1.08)
    With .OA
      If version >= 2 Then
        Input #1, .BasicType
        Input #1, .BegTrow
        Input #1, .EndTrow
      Else
        .BasicType = 0
        .BegTrow = 1
        .EndTrow = 1
      End If
    End With
  
    'Drop Distribution
    'in 1.08, this section was expanded from
    'one DSD set to an array of three
    If version >= 2 Then
      nDrp = 3
    Else
      nDrp = 1
    End If
    For iDrp = 0 To nDrp - 1
      With .DSD(iDrp)
        Input #1, .Type
        Input #1, .BasicType
        Line Input #1, .Name
        Input #1, .LName
        Input #1, .NumDrop
        For i = 0 To .NumDrop - 1
          Input #1, .Diam(i), .MassFrac(i)
        Next
        Input #1, .LibSelFlag
      End With
    Next
    'for single DSD files, copy the first DSD to the other slots
    If version < 2 Then
      .DSD(1) = .DSD(0)
      .DSD(2) = .DSD(0)
    End If
  
    'DropKirk
    '(New in 1.09)
    If version >= 3 Then
      nDrp = 3
      For iDrp = 0 To nDrp - 1
        With .BK(iDrp)
          Input #1, .MaxErrorLevel
          Input #1, .NozType
          Line Input #1, .NameNoz
          Input #1, .LNameNoz
          Input #1, .Orifice
          Input #1, .Speed
          Input #1, .NozAngle
          Input #1, .Pressure
          Input #1, .SprayType
          If version >= 5 Then
            Input #1, .SpectrumSource
          End If
        End With
      Next
    End If
    
    'USDA FS Rotary Nozzle Models
    If version >= 5 Then
      For iDSD = 0 To MAX_DSD - 1
        With .HK(iDSD)
          Input #1, .MaxErrorLevel
          Input #1, .MatType
          If version >= 5 Then
            Input #1, .MatType
          End If
          Input #1, .Speed
          Input #1, .BladeAngle
          If version >= 5 Then
            Input #1, .BladeRPM
          End If
          Input #1, .FlowRate
          Input #1, .SprayType
        End With
      Next
    End If
    
    'DropKick
    'in 1.08, this section was expanded from
    'one DK set to an array of three
    If version >= 2 Then
      nDrp = 3
    Else
      nDrp = 1
    End If
    For iDrp = 0 To nDrp - 1
      With .DK(iDrp)
        'in 1.08 DropKick Solution Type went away
        If version < 2 Then
          Input #1, dummy '.CalcType
        End If
        Input #1, .MaxErrorLevel
        Input #1, .NozType
        Line Input #1, .NameNoz
        Input #1, .LNameNoz
        'in version 1.08 F50,F141,F220 were
        'dropped and .RelSpan was added
        If version < 2 Then
          Input #1, dummy '.F50
          Input #1, dummy '.F141
          Input #1, dummy '.F220
          .RelSpan = 0
          .FlowType = 0
          .flow = 0
        End If
        Input #1, .VMD
        If version >= 2 Then
          Input #1, .RelSpan
        End If
        Input #1, .EffDiam
        'Spray and was introduced in 1.09
        If version >= 3 Then
          Input #1, .SprayAngle
        End If
        Input #1, .MatType
        Line Input #1, .NameMat
        Input #1, .LNameMat
        Input #1, .DynSurfTens
        Input #1, .ShearVisc
        Input #1, .Density
        Input #1, .ElongVisc
        Input #1, .Speed
        Input #1, .NozAngle
        Input #1, .Pressure
        'in 1.08 FlowType and FLow were added
        If version >= 2 Then
          Input #1, .FlowType
          Input #1, .flow
        End If
        'in 1.09 SprayType as added
        If version >= 3 Then
          Input #1, .SprayType
        End If
      End With
    Next
    'for single DK files, copy the first DK to the other slots
    If version < 2 Then
      .DK(1) = .DK(0)
      .DK(2) = .DK(0)
    End If
  
    'Spray Material
    With .SM
      Input #1, .Type
      Input #1, .BasicType
      Line Input #1, .Name
      Input #1, .LName
      If version >= 5 Then
        Input #1, .CalcInputSelect
      End If
      'prior to 1.08 (ver 2) input was NVFrac, NVamt, ACamt
      'In 1.08 only fractions are maintained. The user enters
      'rates in the regulatory side (Smokey=0) and fractions in
      'the FS side (Smokey=1)
      If version >= 2 Then
        Input #1, .NVFrac
        Input #1, .ACFrac
      Else
        Input #1, .NVFrac
        Input #1, dummy '.NVamt
        'Read the number that was ACamt.
        'The value of .ACfrac cannot be calculated until
        'FlowRate and NonVGrav are known.
        'See below
        Input #1, .ACFrac '.ACamt - gets converted to .ACfrac below
      End If
      If version >= 5 Then
        Input #1, .ActSolFrac
        Input #1, .AddSolFrac
        Input #1, .ActNVFrac
        Input #1, .AddNVFrac
      End If
      Input #1, .FlowRate
      If version >= 5 Then
        Input #1, .FlowRateUnits
      End If
      Input #1, .SpecGrav
      '.NonVGrav added in 1.08 (ver 2)
      If version >= 2 Then
        Input #1, .NonVGrav
      Else
        .NonVGrav = .SpecGrav
      End If
      Input #1, .EvapRate
      'Now we can calculate .ACfrac if we need to
      If version < 2 Then
        If .FlowRate * .NonVGrav <> 0 Then
          .ACFrac = .ACFrac / (.FlowRate * .NonVGrav)
        Else
          .ACFrac = .NVFrac
        End If
      End If
      'Starting with AgDRIFT 2.10, Tier II requires active=novol
      If xUD.Tier = TIER_2 Then
        .ACFrac = .NVFrac
      End If
    End With
    
    'Aircraft
    With .AC
      Input #1, .Type
      Input #1, .BasicType
      Line Input #1, .Name
      Input #1, .LName
      Input #1, .WingType
      Input #1, .SemiSpan
      Input #1, .TypSpeed
      Input #1, .BiplSep
      Input #1, .Weight
      Input #1, .PlanArea
      Input #1, .PropRPM
      Input #1, .PropRad
      Input #1, .PropEff
      Input #1, .EngVert
      Input #1, .EngFwd
      'the following were introduced in 1.08
      If version >= 4 Then
        Input #1, .NumEng
        For i = 0 To MAX_ENGINES - 1
          Input #1, .EngHoriz(i)
        Next
        Input #1, .WingVert
        Input #1, .BoomVert
        Input #1, .BoomFwd
      ElseIf version >= 2 Then
        'prior to 2.05 (2.0.5) the EngHoriz array was read in a loop
        'on NumEng, which was incorrect. The end result was satisfactory
        'for 1- or 2-engine aircraft, so keep reading the old files
        'this way.
        Input #1, .NumEng
        For i = 0 To .NumEng - 1
          Input #1, .EngHoriz(i)
        Next
        Input #1, .WingVert
        Input #1, .BoomVert
        Input #1, .BoomFwd
      Else
        .NumEng = 1
        .EngHoriz(0) = 0
        .EngHoriz(1) = 0
        .WingVert = 0
        'prior to 1.08 (ver 2) there was no boom. The NozVert
        'and NozFwd elements of the Nozzles section
        'described this geometry. Leave them blank
        'for now and we will set them in the
        'Nozzles section below
        .BoomVert = 0
        .BoomFwd = 0
      End If
      Input #1, .DragCoeff
    End With
  
    'Nozzles
    With .NZ
      Input #1, .Type
      Input #1, .BasicType
      Line Input #1, .Name
      Input #1, .LName
      Input #1, .NumNoz
      'prior to 1.08, all nozzles shared the same vertial
      'and forward position. in 1.08 fully independent
      'nozzles were introduced.
      'PosVert and PosFwd became arrays, and the concept of
      'a boom was introduced in the aircraft section. Previously
      'PosVert and PosFwd described the boom position. Now when
      'we read these values, install them in the aircraft section.
      If version >= 2 Then
        For i = 0 To .NumNoz - 1
          Input #1, .NozType(i), .PosHoriz(i), _
                    .PosVert(i), .PosFwd(i)
        Next
      Else
        'Read all the horizontal positions
        For i = 0 To .NumNoz - 1
          Input #1, .PosHoriz(i)
        Next
        'Set all the nozzles types to DSD 0,
        'and position them all on the boom
        For i = 0 To .NumNoz - 1
          .NozType(i) = 0
          .PosVert(i) = 0
          .PosFwd(i) = 0
        Next
        'Now read NozVert and NozFwd as the aircraft boom
        'position
        Input #1, xUD.AC.BoomVert
        Input #1, xUD.AC.BoomFwd
      End If
      Input #1, .PosHorizLimit
      '.BoomWidth added in 1.06 (ver 1)
      If (version >= 1) Then
        Input #1, .BoomWidth 'ver 1: added
      Else
        'Get the boom width from the library
        If .Type = 0 Then 'basic
          GetBasicDataNZ .BasicType, xNZ 'get a copy of the basic data
          .BoomWidth = xNZ.BoomWidth
        End If
      End If
    End With
  
    'Meteorology
    With .MET
      Input #1, .WS
      Input #1, .WD
      Input #1, .WindHeight
      Input #1, .temp
      Input #1, .Humidity
      Input #1, .Pressure
      Input #1, .VortexDecay
      Input #1, .SurfRough
      If version >= 5 Then
        Input #1, .Insolation
      End If
    End With
  
    'Canopy
    'the CAN section was introduced in 1.08 (ver 2)
    With .CAN
      If version >= 2 Then
        Input #1, .Type
        Input #1, .Name
        Input #1, .LName
        Input #1, .EleSiz
        Input #1, .StanDen
        Input #1, .NumEnv
        For i = 0 To .NumEnv - 1
          Input #1, .EnvHgt(i), .EnvDiam(i), .EnvPop(i)
        Next
        Input #1, .optType
        Input #1, .LibHgt
        Input #1, .LibLAI
        Input #1, .LibB
        Input #1, .LibC
        Input #1, .NumLAI
        For i = 0 To .NumLAI - 1
          Input #1, .LAIHgt(i), .LAICum(i)
        Next
        Input #1, .temp
        Input #1, .Humidity
        Input #1, .NDRuff
        Input #1, .NDDisp
        Input #1, .Height
      Else
        .Type = 0
        .Name = ""
        .LName = 0
        .EleSiz = 0
        .StanDen = 0
        .NumEnv = 0
        .optType = 1
        .NumLAI = 0
        .temp = xUD.MET.temp
        .Humidity = xUD.MET.Humidity
        .NDRuff = 0.14
        .NDDisp = 0.7
        .Height = 0
      End If
    End With
    
    'Terrain
    'The TRN section was introduced in 1.08 (ver 2)
    With .TRN
      If version >= 2 Then
        Input #1, .Zref
        Input #1, .Upslope
        Input #1, .Sideslope
      Else
        .Zref = 0
        .Upslope = 0
        .Sideslope = 0
      End If
    End With
    
    'Control
    With .CTL
      Input #1, .Height
      Input #1, .NumLines
      Input #1, .SwathWidthType
      Input #1, .SwathWidth
      '.SwathDispType definition changed in 1.09 (ver 3)
      ' Before: 0=1/2 Swath  1=1 Swath 2=Frac App Rate 3=Fixed Dist 4=A/C Ctrline
      ' Now:    0=Frac Swath 1=Frac App Rate 2=Fixed Dist 3=A/C Ctrline
      Input #1, .SwathDispType
      Input #1, .SwathDisp
      If version < 3 Then 'adjust values from older files
        Select Case .SwathDispType
        Case 0 '1/2 sw
          .SwathDispType = 0 'frac sw
          .SwathDisp = 0.5
        Case 1 '1 sw
          .SwathDispType = 0 'frac sw
          .SwathDisp = 1
        Case 2 'frac ar
          .SwathDispType = 1 'frac sw
        Case 3 'fixed dist
          .SwathDispType = 2 'fixed dist
        Case 4 'a/c ctrlin
          .SwathDispType = 3 'a/c ctrline
        End Select
      End If
      Input #1, .FluxPlane
      '.TransHgt(Min/Max) added in 1.08 (ver 2)
      If version >= 2 Then
        Input #1, .TransHgtMin
        Input #1, .TransHgtMax
      Else
        .TransHgtMin = 0
        .TransHgtMax = 0
      End If
      Input #1, .MaxComputeTime
      Input #1, .MaxDownwindDist
      '.HalfBoom added in 1.08 (ver 2)
      If version >= 2 Then
        Input #1, .HalfBoom
      Else
        .HalfBoom = 0
      End If
    End With
  End With
  
  'Always force recalcs when reading old files
  If version < USERFILEVERSION Then
    ClearUserCalc xUC
  Else
    'Read in the current calcs
    With xUC
      Input #1, .Valid
      If .Valid Then
        'calc header data
        Input #1, .CodeVersion
        Line Input #1, .StartDate
        Line Input #1, .StartTime
        Input #1, .MaxErrorLevel
        'Get the Message Log in multiline format
        Input #1, lines
        .MessageLog = "" 'start with a blank
        For i = 1 To lines
          Line Input #1, s                   'get next line
          If i < lines Then s = s & vbCrLf 'append CRLF
          .MessageLog = .MessageLog & s 'append to the total string
        Next
      
        'deposition
        Input #1, .NumDep
        ReDim .DepDist(.NumDep)
        ReDim .DepVal(.NumDep)
        For i = 0 To .NumDep - 1
          Input #1, .DepDist(i), .DepVal(i)
        Next
        Input #1, .DepExtrap
        
        'Vertical Profile
        Input #1, .NumFlux
        ReDim .FluxDist(.NumFlux)
        ReDim .FluxVal(.NumFlux)
        For i = 0 To .NumFlux - 1
          Input #1, .FluxVal(i), .FluxDist(i)
        Next
      
        'average deposition
        Input #1, .NumPID
        ReDim .PIDDist(.NumPID)
        ReDim .PIDVal(.NumPID)
        For i = 0 To .NumPID - 1
          Input #1, .PIDDist(i), .PIDVal(i)
        Next
        Input #1, .PIDExtrap
      
        'Coeffient of Variation
        Input #1, .NumCOV
        ReDim .COVVal(.NumCOV)
        ReDim .COVESW(.NumCOV)
        For i = 0 To .NumCOV - 1
          Input #1, .COVVal(i), .COVESW(i)
        Next
      
        'Concentration
        Input #1, .NumConc
        ReDim .ConcDist(.NumConc)
        ReDim .ConcVal(.NumConc)
        For i = 0 To .NumConc - 1
          Input #1, .ConcVal(i), .ConcDist(i)
        Next
      
        'numerics
        Input #1, .SwathDisp
        Input #1, .SBCOV
        Input #1, .SBMeanDep
        Input #1, .Swath
        Input #1, .AirborneDrift
        Input #1, .EvapFrac
        Input #1, .AppEff
        Input #1, .DownwindDep
        Input #1, .CanopyDep
      
        'Single-swath deposition
        Input #1, .NumSgl
        ReDim .SglDist(.NumSgl)
        ReDim .SglVal(.NumSgl)
        For i = 0 To .NumSgl - 1
          Input #1, .SglDist(i), .SglVal(i)
        Next
      
        'Fraction Aloft
        .NumFA = 0
        Input #1, .NumFA
        ReDim .FADist(.NumFA)
        ReDim .FAVal(.NumFA)
        For i = 0 To .NumFA - 1
          Input #1, .FADist(i), .FAVal(i)
        Next
      
        'COV Mean Deposition
        .NumCOVM = 0
        Input #1, .NumCOVM
        ReDim .COVMDist(.NumCOVM)
        ReDim .COVMVal(.NumCOVM)
        For i = 0 To .NumCOVM - 1
          Input #1, .COVMDist(i), .COVMVal(i)
        Next
      
        'Single Swath Deposition
        .NumHalf = 0
        Input #1, .NumHalf
        ReDim .HalfDist(.NumHalf)
        ReDim .HalfVal(.NumHalf)
        For i = 0 To .NumHalf - 1
          Input #1, .HalfDist(i), .HalfVal(i)
        Next
      
        'Spray Block Deposition
        .NumSBD = 0
        Input #1, .NumSBD
        ReDim .SBDDist(.NumSBD)
        ReDim .SBDVal(.NumSBD)
        For i = 0 To .NumSBD - 1
          Input #1, .SBDDist(i), .SBDVal(i)
        Next
      
        'Canopy Deposition
        .NumCAN = 0
        Input #1, .NumCAN
        ReDim .CANDist(.NumCAN)
        ReDim .CANVal(.NumCAN)
        For i = 0 To .NumCAN - 1
          Input #1, .CANDist(i), .CANVal(i)
        Next
      
        'Time Accountancy Aloft
        .NumTAA = 0
        Input #1, .NumTAA
        ReDim .TAATime(.NumTAA)
        ReDim .TAAVal(.NumTAA)
        For i = 0 To .NumTAA - 1
          Input #1, .TAATime(i), .TAAVal(i)
        Next
      
        'Time Accountancy Vapor
        .NumTAV = 0
        Input #1, .NumTAV
        ReDim .TAVTime(.NumTAV)
        ReDim .TAVVal(.NumTAV)
        For i = 0 To .NumTAV - 1
          Input #1, .TAVTime(i), .TAVVal(i)
        Next
      
        'Time Accountancy Canopy
        .NumTAC = 0
        Input #1, .NumTAC
        ReDim .TACTime(.NumTAC)
        ReDim .TACVal(.NumTAC)
        For i = 0 To .NumTAC - 1
          Input #1, .TACTime(i), .TACVal(i)
        Next
      
        'Time Accountancy Ground
        .NumTAG = 0
        Input #1, .NumTAG
        ReDim .TAGTime(.NumTAG)
        ReDim .TAGVal(.NumTAG)
        For i = 0 To .NumTAG - 1
          Input #1, .TAGTime(i), .TAGVal(i)
        Next
      
        'Distance Accountancy Aloft
        .NumDAA = 0
        Input #1, .NumDAA
        ReDim .DAADist(.NumDAA)
        ReDim .DAAVal(.NumDAA)
        For i = 0 To .NumDAA - 1
          Input #1, .DAADist(i), .DAAVal(i)
        Next
      
        'Distance Accountancy Vapor
        .NumDAV = 0
        Input #1, .NumDAV
        ReDim .DAVDist(.NumDAV)
        ReDim .DAVVal(.NumDAV)
        For i = 0 To .NumDAV - 1
          Input #1, .DAVDist(i), .DAVVal(i)
        Next
      
        'Distance Accountancy Canopy
        .NumDAC = 0
        Input #1, .NumDAC
        ReDim .DACDist(.NumDAC)
        ReDim .DACVal(.NumDAC)
        For i = 0 To .NumDAC - 1
          Input #1, .DACDist(i), .DACVal(i)
        Next
      
        'Distance Accountancy Ground
        .NumDAG = 0
        Input #1, .NumDAG
        ReDim .DAGDist(.NumDAG)
        ReDim .DAGVal(.NumDAG)
        For i = 0 To .NumDAG - 1
          Input #1, .DAGDist(i), .DAGVal(i)
        Next
      
        'Height Accountancy Aloft
        .NumHAA = 0
        Input #1, .NumHAA
        ReDim .HAAHgt(.NumHAA)
        ReDim .HAAVal(.NumHAA)
        For i = 0 To .NumHAA - 1
          Input #1, .HAAHgt(i), .HAAVal(i)
        Next
      
        'Height Accountancy Vapor
        .NumHAV = 0
        Input #1, .NumHAV
        ReDim .HAVHgt(.NumHAV)
        ReDim .HAVVal(.NumHAV)
        For i = 0 To .NumHAV - 1
          Input #1, .HAVHgt(i), .HAVVal(i)
        Next
      
        'Height Accountancy Canopy
        .NumHAC = 0
        Input #1, .NumHAC
        ReDim .HACHgt(.NumHAC)
        ReDim .HACVal(.NumHAC)
        For i = 0 To .NumHAC - 1
          Input #1, .HACHgt(i), .HACVal(i)
        Next
      
        'Spray Block DSD
        .NumSBDSD = 0
        Input #1, .NumSBDSD
        ReDim .SBDSDDiam(.NumSBDSD)
        ReDim .SBDSDFrac(.NumSBDSD)
        For i = 0 To .NumSBDSD - 1
          Input #1, .SBDSDDiam(i), .SBDSDFrac(i)
        Next
      
        'Downwind DSD
        .NumDWDSD = 0
        Input #1, .NumDWDSD
        ReDim .DWDSDDiam(.NumDWDSD)
        ReDim .DWDSDFrac(.NumDWDSD)
        For i = 0 To .NumDWDSD - 1
          Input #1, .DWDSDDiam(i), .DWDSDFrac(i)
        Next
      
        'Vertical Flux DSD
        .NumFXDSD = 0
        Input #1, .NumFXDSD
        ReDim .FXDSDDiam(.NumFXDSD)
        ReDim .FXDSDFrac(.NumFXDSD)
        For i = 0 To .NumFXDSD - 1
          Input #1, .FXDSDDiam(i), .FXDSDFrac(i)
        Next
      
        'Canopy DSD
        .NumCNDSD = 0
        Input #1, .NumCNDSD
        ReDim .CNDSDDiam(.NumCNDSD)
        ReDim .CNDSDFrac(.NumCNDSD)
        For i = 0 To .NumCNDSD - 1
          Input #1, .CNDSDDiam(i), .CNDSDFrac(i)
        Next
      
        'Spray Block Area Coverage
        .NumSBAC = 0
        Input #1, .NumSBAC
        ReDim .SBACRate(.NumSBAC)
        ReDim .SBACFrac(.NumSBAC)
        For i = 0 To .NumSBAC - 1
          Input #1, .SBACRate(i), .SBACFrac(i)
        Next
      
        'Application Layout
        .NumLAY = 0
        Input #1, .NumLAY
        ReDim .LAYDist(.NumLAY)
        ReDim .LAYFrac(.NumLAY)
        For i = 0 To .NumLAY - 1
          Input #1, .LAYDist(i), .LAYFrac(i)
        Next
      End If
    End With
  End If
  
  Close #1
  UserDataRead = True
  Exit Function

UserDataReadErrHand1:
  If Notify Then
    s = "Error opening file: " + fn + Chr$(13) + Error$(Err)
    MsgBox s, vbCritical + vbOKOnly
  End If
  UserDataRead = False
  Exit Function

UserDataReadErrHand2:
  If Notify Then
    s = "Error reading file: " + fn + Chr$(13) + Error$(Err)
    MsgBox s, vbCritical + vbOKOnly
  End If
  Close #1
  UserDataRead = False
  Exit Function
End Function

Function SaveAsFile() As Integer
'Save the current user data to a file,
'always prompting the user for a file name
'
'returns true if sucessful, false if cancelled
'
  SaveAsFile = False 'default
  If Not FileDialog(FD_SAVEAS, FD_TYPE_USER, UI.FileName) Then
    Exit Function
  End If
  If UserDataWrite(UI.FileName, UD, UC, True) Then
    UpdateDataChangedFlag False
    SaveAsFile = True
  End If
End Function

Function SaveFile() As Integer
'Save the current user data to a file, prompting the user
'for a file name only if necessary
'
'returns true if sucessful, false if cancelled
'
  SaveFile = False 'default

  If UI.FileName = "" Then       'if no file name prompt for one
    If Not FileDialog(FD_SAVEAS, FD_TYPE_USER, UI.FileName) Then
      Exit Function
    End If
  End If
  If UserDataWrite(UI.FileName, UD, UC, True) Then
    UpdateDataChangedFlag False
    SaveFile = True
  End If
End Function

Function UnexpectedError(Where As String) As Integer
'General error handling routine
'Just ask the user how to continue with a general message
  Dim Msg As String
  Dim NL As String
  NL = Chr$(13) + Chr$(10)
  Msg = "Unexpected AgDRIFT Error!" + NL + NL
  Msg = Msg + Chr$(34) + Error$(Err) + Chr$(34) + NL + NL
  Msg = Msg + "Where: " + Where
  UnexpectedError = MsgBox(Msg, vbAbortRetryIgnore)
End Function

Sub UpdateDataChangedFlag(NewValue)
'Update the value of UI.DataChanged
'If it has changed, also update the main form caption
'
' NewValue: the new value for UI.DataChanged
'           True=data has changed since last save
'           False=data has not changed since last save
'
  On Error Resume Next 'In case there is no active form
  If NewValue <> UI.DataChanged Then
    UI.DataChanged = NewValue
    frmMain.ActiveForm.Caption = FormCaption
  End If
End Sub

Function FormCaption() As String
'Build a string to act as a form caption.
'Include the current
'file name and an asterisk, if the data has changed
'
  FormCaption = ""
  If UI.DataChanged Then
    FormCaption = "*"   'add an asterisk to indicate changed data
  End If
  If UI.FileName <> "" Then
    FormCaption = FormCaption & UI.FileName 'add the filename
  End If
End Function

Sub WriteGeneralPrefs()
'write settings to the .ini file
  Dim fn As String      '.ini file name
  Dim an As String      'name for .ini file section
  Dim stat As Integer   'return value for functions

  'build the settings file name from the Application
  fn = App.Path & Chr$(92) & App.EXEName & ".ini"
  
  'save preferences
  an = "Preferences"
  stat = WritePrivateProfileString(an, "InitialTier", CStr(UP.InitialTier), fn)
  stat = WritePrivateProfileString(an, "InitialAM", CStr(UP.InitialAM), fn)
  stat = WritePrivateProfileString(an, "InitialAUD", CStr(UP.InitialAUD), fn)
  stat = WritePrivateProfileString(an, "WarnOnTierChange", CStr(UP.WarnOnTierChange), fn)
  stat = WritePrivateProfileString(an, "PauseBeforeCalc", CStr(UP.PauseBeforeCalc), fn)
  stat = WritePrivateProfileString(an, "Units", CStr(UP.Units), fn)
  stat = WritePrivateProfileString(an, "SuppressTier3Warn", CStr(UP.SuppressTier3Warn), fn)
  stat = WritePrivateProfileString(an, "UserLib", UP.UserLib, fn)
End Sub

Function UserDataWrite(fn As String, xUD As UserData, xUC As UserCalc, Notify As Integer)
'Write user data to a file
'fn    i  the file to which to write
'xUD   i  a UserData structure to supply the data
'Notify i Error notification flag:
'           true=notify user of errors through message box
'           false=do not notify
'Return value:
'  true=success
'  false=errors encountered
'
'File format history:
'0: original format
'1: added Header line "AGDRIFT_Format"
'   added version line
'   removed xUI.LibSelDesc
'   added   xUD.DSD.LibSelFlag
'   added Tier III stuff temporarily
'2: big changes to go along with data struct reformat.
'   removed dependance on UI structure.
'3: corrected aircraft EngHoriz writing/reading. Before, this array
'   was incorrectly written in a loop on NumEng. Now the entire
'   2-element array is written, no matter what value NumEng has.
'AgDRIFT version 0.7:
'   changed header line to "AGDRIFT_userdata"
'   reset file version to 0
'   added MesssageLog in CALC section
'
  On Error GoTo UserDataWriteErrHand1
  Open fn For Output As #1
  'write header info
  On Error GoTo UserDataWriteErrHand2
  Print #1, "AGDRIFT_userdata"
  Print #1, USERFILEVERSION  'save the format version
  
  'Userdata
  With xUD
    'General
    Print #1, .Tier
    Print #1, .ApplMethod
    Print #1, .Smokey
    Print #1, .Title
    'write Notes
    lines = LinesInString(.Notes) 'count # of text lines in Notes
    Print #1, lines
    start = 1
    Do
      Print #1, LineFromString(.Notes, start)
    Loop While start > 0
    
    'Ground Application
    With .GA
      Print #1, .BasicType
      Print #1, .NumSwaths
    End With
    
    'Orchard Airblast
    With .OA
      Print #1, .BasicType
      Print #1, .BegTrow
      Print #1, .EndTrow
    End With
    
    'Drop Size Distribution(s)
    For iDSD = 0 To 2
      With .DSD(iDSD)
        Print #1, .Type
        Print #1, .BasicType
        Print #1, .Name
        Print #1, .LName
        Print #1, .NumDrop
        For i = 0 To .NumDrop - 1
          Print #1, .Diam(i); .MassFrac(i)
        Next
        Print #1, .LibSelFlag
      End With
    Next
    
    'Dropkirk
    For iDSD = 0 To 2
      With .BK(iDSD)
        Print #1, .MaxErrorLevel
        Print #1, .NozType
        Print #1, .NameNoz
        Print #1, .LNameNoz
        Print #1, .Orifice
        Print #1, .Speed
        Print #1, .NozAngle
        Print #1, .Pressure
        Print #1, .SprayType
        Print #1, .SpectrumSource
      End With
    Next
    
    'USDA FS Rotary Nozzle Models
    For iDSD = 0 To MAX_DSD - 1
      With .HK(iDSD)
        Print #1, .MaxErrorLevel
        Print #1, .MatType
        Print #1, .RotType
        Print #1, .Speed
        Print #1, .BladeAngle
        Print #1, .BladeRPM
        Print #1, .FlowRate
        Print #1, .SprayType
      End With
    Next
    
    'Dropkick
    For iDSD = 0 To 2
      With .DK(iDSD)
        Print #1, .MaxErrorLevel
        Print #1, .NozType
        Print #1, .NameNoz
        Print #1, .LNameNoz
        Print #1, .VMD
        Print #1, .RelSpan
        Print #1, .EffDiam
        Print #1, .SprayAngle
        Print #1, .MatType
        Print #1, .NameMat
        Print #1, .LNameMat
        Print #1, .DynSurfTens
        Print #1, .ShearVisc
        Print #1, .Density
        Print #1, .ElongVisc
        Print #1, .Speed
        Print #1, .NozAngle
        Print #1, .Pressure
        Print #1, .FlowType
        Print #1, .flow
        Print #1, .SprayType
      End With
    Next
    
    'Spray Material
    With .SM
      Print #1, .Type
      Print #1, .BasicType
      Print #1, .Name
      Print #1, .LName
      Print #1, .CalcInputSelect
      Print #1, .NVFrac
      Print #1, .ACFrac
      Print #1, .ActSolFrac
      Print #1, .AddSolFrac
      Print #1, .ActNVFrac
      Print #1, .AddNVFrac
      Print #1, .FlowRate
      Print #1, .FlowRateUnits
      Print #1, .SpecGrav
      Print #1, .NonVGrav
      Print #1, .EvapRate
    End With
    
    'Aircraft
    With .AC
      Print #1, .Type
      Print #1, .BasicType
      Print #1, .Name
      Print #1, .LName
      Print #1, .WingType
      Print #1, .SemiSpan
      Print #1, .TypSpeed
      Print #1, .BiplSep
      Print #1, .Weight
      Print #1, .PlanArea
      Print #1, .PropRPM
      Print #1, .PropRad
      Print #1, .PropEff
      Print #1, .EngVert
      Print #1, .EngFwd
      Print #1, .NumEng
      For i = 0 To MAX_ENGINES - 1
        Print #1, .EngHoriz(i)
      Next
      Print #1, .WingVert
      Print #1, .BoomVert
      Print #1, .BoomFwd
      Print #1, .DragCoeff
    End With
    
    'Nozzles
    With .NZ
      Print #1, .Type
      Print #1, .BasicType
      Print #1, .Name
      Print #1, .LName
      Print #1, .NumNoz
      For i = 0 To .NumNoz - 1
        Print #1, .NozType(i); .PosHoriz(i); _
                  .PosVert(i); .PosFwd(i)
      Next
      Print #1, .PosHorizLimit
      Print #1, .BoomWidth
    End With
    
    'Meteorology
    With .MET
      Print #1, .WS
      Print #1, .WD
      Print #1, .WindHeight
      Print #1, .temp
      Print #1, .Humidity
      Print #1, .Pressure
      Print #1, .VortexDecay
      Print #1, .SurfRough
      Print #1, .Insolation
    End With
    
    'Canopy
    With .CAN
      Print #1, .Type
      Print #1, .Name
      Print #1, .LName
      Print #1, .EleSiz
      Print #1, .StanDen
      Print #1, .NumEnv
      For i = 0 To .NumEnv - 1
        Print #1, .EnvHgt(i); .EnvDiam(i); .EnvPop(i)
      Next
      Print #1, .optType
      Print #1, .LibHgt
      Print #1, .LibLAI
      Print #1, .LibB
      Print #1, .LibC
      Print #1, .NumLAI
      For i = 0 To .NumLAI - 1
        Print #1, .LAIHgt(i); .LAICum(i)
      Next
      Print #1, .temp
      Print #1, .Humidity
      Print #1, .NDRuff
      Print #1, .NDDisp
      Print #1, .Height
    End With
    
    'Terrain
    With .TRN
      Print #1, .Zref
      Print #1, .Upslope
      Print #1, .Sideslope
    End With
    
    'Control
    With .CTL
      Print #1, .Height
      Print #1, .NumLines
      Print #1, .SwathWidthType
      Print #1, .SwathWidth
      Print #1, .SwathDispType
      Print #1, .SwathDisp
      Print #1, .FluxPlane
      Print #1, .TransHgtMin
      Print #1, .TransHgtMax
      Print #1, .MaxComputeTime
      Print #1, .MaxDownwindDist
      Print #1, .HalfBoom
    End With
  End With
  
  'Read user calcs if there are any
  With xUC
    Print #1, .Valid
    If .Valid Then
      'calc header data
      Print #1, .CodeVersion
      Print #1, .StartDate
      Print #1, .StartTime
      Print #1, .MaxErrorLevel
      'write MessageLog
      lines = LinesInString(.MessageLog) 'count # of text lines
      Print #1, lines
      start = 1
      Do
        Print #1, LineFromString(.MessageLog, start)
      Loop While start > 0
      
      'deposition
      Print #1, .NumDep
      For i = 0 To .NumDep - 1
        Print #1, .DepDist(i); .DepVal(i)
      Next
      Print #1, .DepExtrap
      
      'Vertical Profile
      Print #1, .NumFlux
      For i = 0 To .NumFlux - 1
        Print #1, .FluxVal(i); .FluxDist(i)
      Next
      
      'average deposition
      Print #1, .NumPID
      For i = 0 To .NumPID - 1
        Print #1, .PIDDist(i); .PIDVal(i)
      Next
      Print #1, .PIDExtrap
      
      'Coeffient of Variation
      Print #1, .NumCOV
      For i = 0 To .NumCOV - 1
        Print #1, .COVVal(i); .COVESW(i)
      Next
      
      'Concentration
      Print #1, .NumConc
      For i = 0 To .NumConc - 1
        Print #1, .ConcVal(i); .ConcDist(i)
      Next
      
      'numerics
      Print #1, .SwathDisp
      Print #1, .SBCOV
      Print #1, .SBMeanDep
      Print #1, .Swath
      Print #1, .AirborneDrift
      Print #1, .EvapFrac
      Print #1, .AppEff
      Print #1, .DownwindDep
      Print #1, .CanopyDep
      
      'Single-swath deposition
      Print #1, .NumSgl
      For i = 0 To .NumSgl - 1
        Print #1, .SglDist(i); .SglVal(i)
      Next
      
      'Fraction Aloft
      Print #1, .NumFA
      For i = 0 To .NumFA - 1
        Print #1, .FADist(i); .FAVal(i)
      Next
      
      'COV Mean Deposition
      Print #1, .NumCOVM
      For i = 0 To .NumCOVM - 1
        Print #1, .COVMDist(i); .COVMVal(i)
      Next
      
      'Single Swath Deposition
      Print #1, .NumHalf
      For i = 0 To .NumHalf - 1
        Print #1, .HalfDist(i); .HalfVal(i)
      Next
      
      'Spray Block Deposition
      Print #1, .NumSBD
      For i = 0 To .NumSBD - 1
        Print #1, .SBDDist(i); .SBDVal(i)
      Next
      
      'Canopy Deposition
      Print #1, .NumCAN
      For i = 0 To .NumCAN - 1
        Print #1, .CANDist(i); .CANVal(i)
      Next
      
      'Time Accountancy Aloft
      Print #1, .NumTAA
      For i = 0 To .NumTAA - 1
        Print #1, .TAATime(i); .TAAVal(i)
      Next
      
      'Time Accountancy Vapor
      Print #1, .NumTAV
      For i = 0 To .NumTAV - 1
        Print #1, .TAVTime(i); .TAVVal(i)
      Next
      
      'Time Accountancy Canopy
      Print #1, .NumTAC
      For i = 0 To .NumTAC - 1
        Print #1, .TACTime(i); .TACVal(i)
      Next
      
      'Time Accountancy Ground
      Print #1, .NumTAG
      For i = 0 To .NumTAG - 1
        Print #1, .TAGTime(i); .TAGVal(i)
      Next
      
      'Distance Accountancy Aloft
      Print #1, .NumDAA
      For i = 0 To .NumDAA - 1
        Print #1, .DAADist(i); .DAAVal(i)
      Next
      
      'Distance Accountancy Vapor
      Print #1, .NumDAV
      For i = 0 To .NumDAV - 1
        Print #1, .DAVDist(i); .DAVVal(i)
      Next
      
      'Distance Accountancy Canopy
      Print #1, .NumDAC
      For i = 0 To .NumDAC - 1
        Print #1, .DACDist(i); .DACVal(i)
      Next
      
      'Distance Accountancy Ground
      Print #1, .NumDAG
      For i = 0 To .NumDAG - 1
        Print #1, .DAGDist(i); .DAGVal(i)
      Next
      
      'Height Accountancy Aloft
      Print #1, .NumHAA
      For i = 0 To .NumHAA - 1
        Print #1, .HAAHgt(i); .HAAVal(i)
      Next
      
      'Height Accountancy Vapor
      Print #1, .NumHAV
      For i = 0 To .NumHAV - 1
        Print #1, .HAVHgt(i); .HAVVal(i)
      Next
      
      'Height Accountancy Canopy
      Print #1, .NumHAC
      For i = 0 To .NumHAC - 1
        Print #1, .HACHgt(i); .HACVal(i)
      Next
      
      'Spray Block DSD
      Print #1, .NumSBDSD
      For i = 0 To .NumSBDSD - 1
        Print #1, .SBDSDDiam(i); .SBDSDFrac(i)
      Next
      
      'Downwind DSD
      Print #1, .NumDWDSD
      For i = 0 To .NumDWDSD - 1
        Print #1, .DWDSDDiam(i); .DWDSDFrac(i)
      Next
      
      'Vertical Flux DSD
      Print #1, .NumFXDSD
      For i = 0 To .NumFXDSD - 1
        Print #1, .FXDSDDiam(i); .FXDSDFrac(i)
      Next
      
      'Canopy DSD
      Print #1, .NumCNDSD
      For i = 0 To .NumCNDSD - 1
        Print #1, .CNDSDDiam(i); .CNDSDFrac(i)
      Next
      
      'Spray Block Area Coverage
      Print #1, .NumSBAC
      For i = 0 To .NumSBAC - 1
        Print #1, .SBACRate(i); .SBACFrac(i)
      Next
    
      'Application Layout
      Print #1, .NumLAY
      For i = 0 To .NumLAY - 1
        Print #1, .LAYDist(i); .LAYFrac(i)
      Next
    End If
  End With
    
  Close #1
  UserDataWrite = True
  Exit Function

UserDataWriteErrHand1:
  If Notify Then
    s$ = "Error opening file: " + fn + Chr$(13) + Error$(Err)
    MsgBox s$, vbCritical + vbOKOnly
  End If
  UserDataWrite = False
  Exit Function

UserDataWriteErrHand2:
  If Notify Then
    s$ = "Error writing file: " + fn + Chr$(13) + Error$(Err)
    MsgBox s$, vbCritical + vbOKOnly
  End If
  Close #1
  UserDataWrite = False
  Exit Function
End Function

Sub GetTrialDataInp(key As String, xUD As UserData)
'Get Field Trial Library data. Gets only input data
'
' key  i  a search key comprised of a Title field
'         entry from frmTrialLib. If the key is blank
'         or bad, xUD will be untouched.
'
' xUD  o  a fully formatted UserData structure retrieved
'         from the library. Defaults are set for those
'         values not recovered from the library.
'
  Dim DB As Database
  Dim DS As Recordset
  Dim crit As String
  Dim s As String

  If Trim$(key) = "" Then Exit Sub

  If Not LibOpenDB(DB) Then Exit Sub
  Set DS = DB.OpenRecordset("FieldTrial", dbOpenDynaset)
  
  'position the Recordset to the correct record
  crit = "Title = '" & key & "'"
  DS.FindFirst crit
  If DS.NoMatch Then
    DS.Close
    DB.Close
    Exit Sub
  End If

  'Place default data in the UD structure
  UserDataDefault xUD
  
  '
  'Retrieve the data from the library
  '
  
  'Not used:
  'xUD = DS.Fields("TestNumber")
  'xUD = DS.Fields("WindSpeed")
  'xUD = DS.Fields("Height")
  
  'General:
  xUD.Title = DS.Fields("Title")
  
  'DSD
  xUD.DSD(0).Type = 2  'User-def
  xUD.DSD(0).Name = xUD.Title
  xUD.DSD(0).LName = Len(xUD.DSD(0).Name)
  xUD.DSD(0).NumDrop = DS.Fields("DSDNumDrop")
  FieldToArray DS.Fields("DSDDiam"), xUD.DSD(0).Diam()
  FieldToArray DS.Fields("DSDMassFrac"), xUD.DSD(0).MassFrac()
  
  'SM
  xUD.SM.Type = 1 'User-def
  Select Case DS.Fields("SprayMaterial")
    Case 0:
      xUD.SM.Name = "Standard"
      xUD.SM.LName = Len(xUD.SM.Name)
    Case 1:
      xUD.SM.Name = "Variable"
      xUD.SM.LName = Len(xUD.SM.Name)
  End Select
  xUD.SM.NVFrac = DS.Fields("SMNVrate") / (DS.Fields("SMFlowRate") * DS.Fields("SMSpecGrav"))
  xUD.SM.ACFrac = DS.Fields("SMACrate") / (DS.Fields("SMFlowRate") * DS.Fields("SMSpecGrav"))
  xUD.SM.FlowRate = DS.Fields("SMFlowRate")
  xUD.SM.SpecGrav = DS.Fields("SMSpecGrav")
  xUD.SM.EvapRate = DS.Fields("SMEvapRate")
  
  'AC
  xUD.AC.Type = 1  'user-def
  xUD.AC.BasicType = DS.Fields("Aircraft")
  xUD.AC.Name = DS.Fields("ACName")
  xUD.AC.LName = Len(xUD.AC.Name)
  xUD.AC.WingType = DS.Fields("ACWingType")
  xUD.AC.SemiSpan = DS.Fields("ACSemiSpan")
  xUD.AC.TypSpeed = DS.Fields("ACTypSpeed")
  xUD.AC.BiplSep = DS.Fields("ACBiplSep")
  xUD.AC.Weight = DS.Fields("ACWeight")
  xUD.AC.PlanArea = DS.Fields("ACPlanArea")
  xUD.AC.PropRPM = DS.Fields("ACPropRPM")
  xUD.AC.PropRad = DS.Fields("ACPropRad")
  xUD.AC.EngVert = DS.Fields("ACEngVert")
  xUD.AC.EngFwd = DS.Fields("ACEngFwd")
  xUD.AC.BoomVert = DS.Fields("NZPosVert")
  xUD.AC.BoomFwd = DS.Fields("NZPosFwd")
  
  'NZ
  xUD.NZ.Type = 1  'user-def
  xUD.NZ.Name = xUD.Title
  xUD.NZ.LName = Len(xUD.NZ.Name)
  xUD.NZ.NumNoz = DS.Fields("NZNumNoz")
  FieldToArray DS.Fields("NZPosHoriz"), xUD.NZ.PosHoriz()
  For i = 0 To xUD.NZ.NumNoz - 1
    xUD.NZ.PosVert(i) = 0
    xUD.NZ.PosFwd(i) = 0
  Next
  
  'MET
  xUD.MET.WS = DS.Fields("METWS")
  xUD.MET.WD = DS.Fields("METWD")
  xUD.MET.temp = DS.Fields("METtemp")
  xUD.MET.Humidity = DS.Fields("METHumidity")
  xUD.MET.SurfRough = DS.Fields("METCanopyHeight")
  
  'CTL
  xUD.CTL.Height = DS.Fields("CTLHeight")
  xUD.CTL.NumLines = DS.Fields("CTLNumLines")
  xUD.CTL.SwathWidthType = 0  'fixed
  xUD.CTL.SwathWidth = DS.Fields("CTLSwathWidth")
  xUD.CTL.SwathDispType = 2   'fixed
  xUD.CTL.SwathDisp = 0       '0 distance
  
  DS.Close
  DB.Close
End Sub

Sub GetTrialDataDSD(key As String, np As Integer, Diam() As Single, mfrac() As Single)
'Get Field Trial Library data. Gets only Dropsize distribution
'
' key     i  a search key comprised of a Title field
'            entry from frmTrialLib. If the key is blank
'            or bad, xUD will be untouched.
'
' np      o  the number of points returned
'
' Diam()  o  diameters returned
' mfrac() o  mass fractions returned
'
  Dim DB As Database
  Dim DS As Recordset
  Dim crit As String
  Dim s As String

  If Trim$(key) = "" Then Exit Sub

  If Not LibOpenDB(DB) Then Exit Sub
  Set DS = DB.OpenRecordset("FieldTrial", dbOpenDynaset)
  
  'position the Recordset to the correct record
  crit = "Title = '" & key & "'"
  DS.FindFirst crit
  If DS.NoMatch Then
    np = 0
    DS.Close
    DB.Close
    Exit Sub
  End If

  '
  'Retrieve the data from the library
  '
  
  'DSD
  np = DS.Fields("DSDNumDrop")
  FieldToArray DS.Fields("DSDDiam"), Diam()
  FieldToArray DS.Fields("DSDMassFrac"), mfrac()
  
  DS.Close
  DB.Close
End Sub

Sub GetTrialDataPred(PlotVar As Long, key As String, np As Integer, Dist() As Single, dep() As Single)
'Get Field Trial Library data. Gets only predicted deposition.
'
' PlotVar  i  PV_DEP   2=deposition
'             PV_PID   4=pond-integrated deposition
'
' key    i  a search key comprised of a Title field
'           entry from frmTrialLib. If the key is blank
'           or bad, xUD will be untouched.
'
' np     o  the number of points returned
'
' Dist() o  distances returned
' Dep()  o  deposition returned
'
  Dim DB As Database
  Dim DS As Recordset
  Dim crit As String
  Dim s As String
  Dim NXtmp As Long

  If Trim$(key) = "" Then Exit Sub

  If Not LibOpenDB(DB) Then Exit Sub
  Set DS = DB.OpenRecordset("FieldTrial", dbOpenDynaset)
  
  'position the Recordset to the correct record
  crit = "Title = '" & key & "'"
  DS.FindFirst crit
  If DS.NoMatch Then
    np = 0
    DS.Close
    DB.Close
    Exit Sub
  End If

  '
  'Retrieve the data from the library
  '
  
  'CALC
  np = DS.Fields("CALCNumPredDep")
  FieldToArray DS.Fields("CALCPredDepDist"), Dist()
  FieldToArray DS.Fields("CALCPredDepVal"), dep()
  
  DS.Close
  DB.Close
  
  Select Case PlotVar
  Case PV_DEP 'deposition
    NXtmp = np
    Call aglibr(CLng(UD.Tier), 0, UD.CTL.MaxDownwindDist, NXtmp, Dist(0), dep(0))
    np = NXtmp
  Case PV_PID 'PID
    NXtmp = np
    Call aglibr(CLng(UD.Tier), 1, UD.CTL.MaxDownwindDist, NXtmp, Dist(0), dep(0))
    np = NXtmp
  End Select
End Sub

Sub GetTrialDataMeas(PlotVar As Long, key As String, np As Integer, Dist() As Single, dep() As Single)
'Get Field Trial Library data. Gets only measured deposition.
'
' PlotVar  i  PV_DEP   2=deposition
'             PV_PID   4=pond-integrated deposition
'
' key    i  a search key comprised of a Title field
'           entry from frmTrialLib. If the key is blank
'           or bad, xUD will be untouched.
'
' np     o  the number of points returned
'
' Dist() o  distances returned
' Dep()  o  deposition returned
'
  Dim DB As Database
  Dim DS As Recordset
  Dim crit As String
  Dim s As String
  Dim tmpDist() As Single
  Dim tmpVal1() As Single
  Dim tmpVal2() As Single
  Dim tmpVal3() As Single
  Dim tmpVal4() As Single
  Dim NXtmp As Long

  If Trim$(key) = "" Then Exit Sub

  If Not LibOpenDB(DB) Then Exit Sub
  Set DS = DB.OpenRecordset("FieldTrial", dbOpenDynaset)
  
  'position the Recordset to the correct record
  crit = "Title = '" & key & "'"
  DS.FindFirst crit
  If DS.NoMatch Then
    np = 0
    DS.Close
    DB.Close
    Exit Sub
  End If

  '
  'Retrieve the data from the library
  '
  
  'CALC
  np = DS.Fields("CALCNumMeasDep")
  ReDim tmpDist(np)
  ReDim tmpVal1(np)
  ReDim tmpVal2(np)
  ReDim tmpVal3(np)
  ReDim tmpVal4(np)
  FieldToArray DS.Fields("CALCMeasDepDist"), tmpDist()
  FieldToArray DS.Fields("CALCMeasDepVal1"), tmpVal1()
  FieldToArray DS.Fields("CALCMeasDepVal2"), tmpVal2()
  FieldToArray DS.Fields("CALCMeasDepVal3"), tmpVal3()
  FieldToArray DS.Fields("CALCMeasDepVal4"), tmpVal4()
  
  DS.Close
  DB.Close
  
  Select Case PlotVar
  Case PV_DEP 'deposition
    'average and pack up final arrays
    For i = 0 To np - 1
      Dist(i) = tmpDist(i)
      dep(i) = (tmpVal1(i) + tmpVal2(i) + tmpVal3(i) + tmpVal4(i)) * 0.25
    Next
  Case PV_PID 'PID
    'average and pack up final arrays
    For i = 0 To np - 1
      Dist(i) = tmpDist(i)
      dep(i) = (tmpVal1(i) + tmpVal2(i) + tmpVal3(i) + tmpVal4(i)) * 0.25
    Next
  End Select
End Sub

Sub GetTrialDataCalc(key As String, xUC As UserCalc)
'Get Field Trial Library data. Gets only predicted deposition.
'
' key  i  a search key comprised of a Title field
'         entry from frmTrialLib. If the key is blank
'         or bad, xUD will be untouched.
'
' xUD  o  a fully formatted UserData structure retrieved
'         from the library. Defaults are set for those
'         values not recovered from the library.
'
  Dim DB As Database
  Dim DS As Recordset
  Dim crit As String
  Dim s As String
  Dim NXY As Long
  Dim X(MAX_CALCDATA) As Single
  Dim Y(MAX_CALCDATA) As Single

  If Trim$(key) = "" Then Exit Sub

  If Not LibOpenDB(DB) Then Exit Sub
  Set DS = DB.OpenRecordset("FieldTrial", dbOpenDynaset)
  
  'position the Recordset to the correct record
  crit = "Title = '" & key & "'"
  DS.FindFirst crit
  If DS.NoMatch Then
    DS.Close
    DB.Close
    Exit Sub
  End If

  '
  'Retrieve the data from the library
  '
  xUC.Valid = True
  
  'Deposition
  NXY = CLng(DS.Fields("CALCNumPredDep"))
  FieldToArray DS.Fields("CALCPredDepDist"), X()
  FieldToArray DS.Fields("CALCPredDepVal"), Y()
  'Massage the data to get the proper deposition
  Call aglibr(CLng(UD.Tier), 0, UD.CTL.MaxDownwindDist, NXY, X(0), Y(0))
  xUC.NumDep = NXY
  ReDim xUC.DepDist(NXY)
  ReDim xUC.DepVal(NXY)
  CopyMemory xUC.DepDist(0), X(0), Len(X(0)) * NXY
  CopyMemory xUC.DepVal(0), Y(0), Len(Y(0)) * NXY
  
  'Pond-integrated deposition
  NXY = CLng(DS.Fields("CALCNumPredDep"))
  FieldToArray DS.Fields("CALCPredDepDist"), X()
  FieldToArray DS.Fields("CALCPredDepVal"), Y()
  'Massage the data to get the proper deposition
  Call aglibr(CLng(UD.Tier), 1, UD.CTL.MaxDownwindDist, NXY, X(0), Y(0))
  xUC.NumPID = NXY
  ReDim xUC.PIDDist(NXY)
  ReDim xUC.PIDVal(NXY)
  CopyMemory xUC.PIDDist(0), X(0), Len(X(0)) * NXY
  CopyMemory xUC.PIDVal(0), Y(0), Len(Y(0)) * NXY
  
  'Flux
  xUC.NumFlux = 0
  ReDim xUC.FluxDist(0)
  ReDim xUC.FluxVal(0)
  
  'COV
  xUC.NumCOV = 0
  ReDim xUC.COVVal(0)
  ReDim xUC.COVESW(0)
  
  'Conc
  xUC.NumConc = 0
  ReDim xUC.ConcDist(0)
  ReDim xUC.ConcVal(0)
  
  xUC.SwathDisp = DS.Fields("SwathDisp")
  xUC.Swath = 0
  xUC.AirborneDrift = DS.Fields("Airborne")
  xUC.EvapFrac = DS.Fields("EvapFrac")
  xUC.DownwindDep = DS.Fields("DwndDrift")
  xUC.AppEff = DS.Fields("AppEff")
  xUC.NumSgl = 0
  ReDim xUC.SglDist(0)
  ReDim xUC.SglVal(0)
  xUC.NumFA = 0
  ReDim xUC.FADist(0)
  ReDim xUC.FAVal(0)
  
  DS.Close
  DB.Close
End Sub

Public Function LoadFieldTrial()
'Loads Field trial input and calculation data into
'the current user data space

  Dim LibKey As String
  
  LoadFieldTrial = False  'default return status
  
  'save changed data first
  If Not QuerySaveChanged() Then Exit Function
  
  'get the library key from the browse form
  frmTrialLib.Show vbModal
  LibKey = frmTrialLib.Tag
  Unload frmTrialLib
  
  If LibKey <> "" Then
    'load new data
    GetTrialDataInp LibKey, UD
    GetTrialDataCalc LibKey, UC
    UpdateDataChangedFlag True
    LoadFieldTrial = True
  End If
End Function

Public Function ClipStr$(s, wid As Integer)
'Clips a string to the specified width. If the string is
'less than or equal to wid, it is returned unchanged.
'Otherwise, the first part of the string is returned with
'"..." appended to it. The length of the string returned
'in the latter case is wid.
  Dim tmpstr As String
  
  tmpstr = Trim$(s)
  If Len(tmpstr) <= wid Then
    ClipStr$ = tmpstr
  Else
    If wid > 3 Then
      ClipStr$ = Left$(tmpstr, wid - 3) & "..."
    Else
      ClipStr$ = Left$(tmpstr, wid)
    End If
  End If
End Function

Public Function LibOpenDB(DB As Database) As Integer
'Open AgDRIFT Library Database
  Dim mbtitle As String
  Dim mbflags As Integer
  Dim mbmsg As String
  
  On Error GoTo ErrHandlerLODB
  LibOpenDB = True 'default return value
  Set DB = Workspaces(0).OpenDatabase(UI.LibraryPath, False, True)
  
ExitLibOpenDB:
  Exit Function

ErrHandlerLODB:
  Select Case Err.Number
  Case 3024  'Database file not found
    mbtitle = "AgDRIFT Library Error"
    mbmsg = "The AgDRIFT Library file was not found." + vbCrLf + vbCrLf
    mbmsg = mbmsg + "AgDRIFT cannot continue without this library and will terminate."
    mbflags = vbExclamation + vbOKOnly
    MsgBox mbmsg, mbflags, mbtitle
    End
  End Select
  LibOpenDB = False
  Resume ExitLibOpenDB
End Function

Public Function LibOpenMAADB(DB As Database) As Integer
'Open AgDRIFT MAA Library Database
  Dim mbtitle As String
  Dim mbflags As Integer
  Dim mbmsg As String
  Dim RS As Recordset
  
  On Error GoTo ErrHandlerLOMAADB
  LibOpenMAADB = True 'default return value
  Set DB = Workspaces(0).OpenDatabase(UI.MAALibraryPath, False, True)

  'Database is open. Check version.
  'First look for the Info table. If it is not there, the library
  'is very old or very wrong. Then match the Version field in the
  'database to the version wired in this program.
  If LibOpenRS(DB, "Info", RS) Then
    If RS.Fields("Version") = LIBRARYVERSION Then
      RS.Close
      Exit Function  'Success!
    Else
      'Version mismatch
      RS.Close
      Error 10000 'Raise a fake error for Version Mismatch
    End If
  Else
    'No Info table, must be an old library
    Error 10000 'Raise a fake error for Version Mismatch
  End If
  
ExitLibOpenMAADB:
  Exit Function

ErrVersionLOMAADB:
  'Version mismatch error
  mbtitle = "AgDRIFT Library Error"
  mbmsg = "The AgDRIFT MAA Library file is the wrong version." + vbCrLf + vbCrLf
  mbmsg = mbmsg + "AgDRIFT cannot open this toolbox without the correct library."
  mbflags = vbExclamation + vbOKOnly
  MsgBox mbmsg, mbflags, mbtitle
  LibOpenMAADB = False
  GoTo ExitLibOpenMAADB
  
ErrHandlerLOMAADB:
  mbtitle = "AgDRIFT Library Error"
  mbflags = vbExclamation + vbOKOnly
  Select Case Err.Number
  Case 3024  'Database file not found
    mbmsg = "The AgDRIFT MAA Library file was not found." + vbCrLf + vbCrLf
    mbmsg = mbmsg + "AgDRIFT cannot open this toolbox without the correct library."
  Case 10000  '(user-defined) Library Version mismatch
    mbmsg = "The AgDRIFT MAA Library file is the wrong version." + vbCrLf + vbCrLf
    mbmsg = mbmsg + "AgDRIFT cannot open this toolbox without the correct library."
  End Select
  MsgBox mbmsg, mbflags, mbtitle
  LibOpenMAADB = False
  Resume ExitLibOpenMAADB
End Function

Public Function LibOpenRS(DB As Database, TableName As String, RS As Recordset) As Integer
'Open a Recordset within an open Database
  
  On Error GoTo ErrHandlerLORS
  LibOpenRS = True 'default return value
  Set RS = DB.OpenRecordset(TableName, dbOpenDynaset)
  Exit Function
  
ErrHandlerLORS:
  Select Case Err.Number
  Case 3011  '
    s$ = Err.Description
    'no message, just return false
  End Select
  LibOpenRS = False
  Resume Next
End Function

Public Function FieldTrialDataExists() As Integer
'test for existance of field trial data
  Dim DB As Database
  Dim RS As Recordset
  
  FieldTrialDataExists = False
  If LibOpenDB(DB) Then
    If LibOpenRS(DB, "FieldTrial", RS) Then
      FieldTrialDataExists = True
      RS.Close
    End If
    DB.Close
  End If
End Function

Public Sub ClearUserCalc(xUC As UserCalc)
'Reset UserCalc structure to clear calc data
'Set the arrays to 1 element, rather than erasing them
'in case the empty array is referenced. This avoids
'"Subscript out of range" errors for calling a routine
'by referencing the zero-th element
  xUC.Valid = False
  xUC.CodeVersion = 0
  xUC.StartDate = ""
  xUC.StartTime = ""
  xUC.MaxErrorLevel = 0
  xUC.MessageLog = ""
  xUC.NumDep = 0
  ReDim xUC.DepDist(0)
  ReDim xUC.DepVal(0)
  xUC.DepExtrap = 0
  xUC.NumPID = 0
  ReDim xUC.PIDDist(0)
  ReDim xUC.PIDVal(0)
  xUC.PIDExtrap = 0
  xUC.NumFlux = 0
  ReDim xUC.FluxDist(0)
  ReDim xUC.FluxVal(0)
  xUC.NumCOV = 0
  ReDim xUC.COVVal(0)
  ReDim xUC.COVESW(0)
  xUC.NumConc = 0
  ReDim xUC.ConcDist(0)
  ReDim xUC.ConcVal(0)
  xUC.SwathDisp = 0
  xUC.SBCOV = 0
  xUC.SBMeanDep = 0
  xUC.Swath = 0
  xUC.AirborneDrift = 0
  xUC.EvapFrac = 0
  xUC.AppEff = 0
  xUC.DownwindDep = 0
  xUC.CanopyDep = 0
  xUC.NumSgl = 0
  ReDim xUC.SglDist(0)
  ReDim xUC.SglVal(0)
  xUC.NumFA = 0
  ReDim xUC.FADist(0)
  ReDim xUC.FAVal(0)
  xUC.NumCOVM = 0
  ReDim xUC.COVMDist(0)
  ReDim xUC.COVMVal(0)
  xUC.NumHalf = 0
  ReDim xUC.HalfDist(0)
  ReDim xUC.HalfVal(0)
  xUC.NumSBD = 0
  ReDim xUC.SBDDist(0)
  ReDim xUC.SBDVal(0)
  xUC.NumCAN = 0
  ReDim xUC.CANDist(0)
  ReDim xUC.CANVal(0)
  xUC.NumTAA = 0
  ReDim xUC.TAATime(0)
  ReDim xUC.TAAVal(0)
  xUC.NumTAV = 0
  ReDim xUC.TAVTime(0)
  ReDim xUC.TAVVal(0)
  xUC.NumTAC = 0
  ReDim xUC.TACTime(0)
  ReDim xUC.TACVal(0)
  xUC.NumTAG = 0
  ReDim xUC.TAGTime(0)
  ReDim xUC.TAGVal(0)
  xUC.NumDAA = 0
  ReDim xUC.DAADist(0)
  ReDim xUC.DAAVal(0)
  xUC.NumDAV = 0
  ReDim xUC.DAVDist(0)
  ReDim xUC.DAVVal(0)
  xUC.NumDAC = 0
  ReDim xUC.DACDist(0)
  ReDim xUC.DACVal(0)
  xUC.NumDAG = 0
  ReDim xUC.DAGDist(0)
  ReDim xUC.DAGVal(0)
  xUC.NumHAA = 0
  ReDim xUC.HAAHgt(0)
  ReDim xUC.HAAVal(0)
  xUC.NumHAV = 0
  ReDim xUC.HAVHgt(0)
  ReDim xUC.HAVVal(0)
  xUC.NumHAC = 0
  ReDim xUC.HACHgt(0)
  ReDim xUC.HACVal(0)
  xUC.NumSBDSD = 0
  ReDim xUC.SBDSDDiam(0)
  ReDim xUC.SBDSDFrac(0)
  xUC.NumDWDSD = 0
  ReDim xUC.DWDSDDiam(0)
  ReDim xUC.DWDSDFrac(0)
  xUC.NumFXDSD = 0
  ReDim xUC.FXDSDDiam(0)
  ReDim xUC.FXDSDFrac(0)
  xUC.NumCNDSD = 0
  ReDim xUC.CNDSDDiam(0)
  ReDim xUC.CNDSDFrac(0)
  xUC.NumSBAC = 0
  ReDim xUC.SBACRate(0)
  ReDim xUC.SBACFrac(0)
  xUC.NumLAY = 0
  ReDim xUC.LAYDist(0)
  ReDim xUC.LAYFrac(0)
End Sub

Public Function PrinterExists()
'Test see if there is a default printer configured
  Dim Msg As String
  If Printers.Count > 0 Then
    PrinterExists = True
  Else
    Msg = "There is no printer configured."
    MsgBox Msg, vbOKOnly
    PrinterExists = False
  End If
End Function

Public Function LibGetFileName() As String
'Get the file name of the Library Database
  
  Dim SDTFlib As String
  Dim STDlib As String
  Dim SDTFExists As Boolean
  Dim STDExists As Boolean
  Dim SDTFVersionOK As Boolean
  Dim STDVersionOK As Boolean
  
  Dim Msg As String
  Dim Mflags As Long
  Dim Mtitle As String
  
  'Set up candidate library paths
  SDTFlib = App.Path & "\" & "agdsdtf.mdb"
  STDlib = App.Path & "\" & "agdrift.mdb"
  
  'Check for file existance
  SDTFExists = FileExists(SDTFlib)
  STDExists = FileExists(STDlib)
  
  'Check that the library version is okay
  SDTFVersionOK = LibVersionMatches(SDTFlib)
  STDVersionOK = LibVersionMatches(STDlib)
    
  'Now for the big decision tree
  Mtitle = "AgDRIFT Library Error"
  Mflags = vbOKOnly
  Msg = ""
  If SDTFExists Then
    If SDTFVersionOK Then
      'SDTF Lib OK - Silently use SDTF Confidential Library
      LibGetFileName = SDTFlib
      'Remeber that this user has access to the confidential data
      '(Thisa makes a difference in some forms)
      UI.HasConfidentialData = True
    Else
      If STDExists Then
        If STDVersionOK Then
          'SDTF bad, STD good
          Msg = Msg + "The AgDRIFT Confidential Library is the wrong version." + vbCrLf + vbCrLf
          Msg = Msg + "AgDRIFT will continue with the default library."
          Mflags = Mflags + vbExclamation
          MsgBox Msg, Mflags, Mtitle
          LibGetFileName = STDlib
        Else
          'SDTF bad, STD bad
          Msg = Msg + "The AgDRIFT Library files are the wrong version." + vbCrLf + vbCrLf
          Msg = Msg + "AgDRIFT cannot "
          Msg = Msg + "continue without a correct library and will terminate."
          Mflags = Mflags + vbCritical
          MsgBox Msg, Mflags, Mtitle
          End
        End If
      Else
        'SDTF bad, STD missing
        Msg = Msg + "The AgDRIFT Confidential Library is the wrong version "
        Msg = Msg + "and no default library was found." + vbCrLf + vbCrLf
        Msg = Msg + "AgDRIFT cannot "
        Msg = Msg + "continue without the correct library and will terminate."
        Mflags = Mflags + vbCritical
        MsgBox Msg, Mflags, Mtitle
        End
      End If
    End If
  Else
    If STDExists Then
      If STDVersionOK Then
        'SDTF missing, STD good - Silently use Standard library
        LibGetFileName = STDlib
      Else
        'SDTF missing, STD bad
        Msg = Msg + "The AgDRIFT Library file is the wrong version." + vbCrLf + vbCrLf
        Msg = Msg + "AgDRIFT cannot "
        Msg = Msg + "continue without the correct library and will terminate."
        Mflags = Mflags + vbCritical
        MsgBox Msg, Mflags, Mtitle
        End
      End If
    Else
      'SDTF missing, STD missing
      Msg = Msg + "The AgDRIFT Library file was not found." + vbCrLf + vbCrLf
      Msg = Msg + "AgDRIFT cannot "
      Msg = Msg + "continue without the library and will terminate."
      Mflags = Mflags + vbCritical
      MsgBox Msg, Mflags, Mtitle
      End
    End If
  End If
End Function

Public Function LibVersionMatches(LibraryPath As String) As Integer
'Open the given library file and check its version against the program's
'Return True if all is well, False is there is any kind of problem.
  Dim DB As Database
  Dim RS As Recordset

  'Try to open the database
  On Error GoTo LibVersionMatchesErrHand
  Set DB = Workspaces(0).OpenDatabase(LibraryPath, False, True)
  On Error GoTo 0
  
  'Database is open. Check version.
  'First look for the Info table. If it is not there, the library
  'is very old or very wrong. Then match the Version field in the
  'database to the version wired in this program.
  If LibOpenRS(DB, "Info", RS) Then
    If RS.Fields("Version") = LIBRARYVERSION Then
      'Success!
      LibVersionMatches = True
    Else
      'Version mismatch
      LibVersionMatches = False
    End If
    RS.Close
  Else
    'No Info table
    LibVersionMatches = False
  End If
  DB.Close

ExitLibVersionMatches:
  Exit Function
  
LibVersionMatchesErrHand:
  LibVersionMatches = False
  Resume ExitLibVersionMatches
End Function

Public Function QueryPerformCalcs() As Integer
'Ask the user about do calcs, then proceed
' Returns: True if calcs were performed
'          False if not
'
  Dim Msg As String
  
  QueryPerformCalcs = False 'default return value
  
  Msg = "Calculations must be performed before "
  Msg = Msg + "continuing. Perform them now?"
  MBType = vbQuestion + vbOKCancel
  If MsgBox(Msg, MBType) = vbOK Then
    'try to do the calcs
    QueryPerformCalcs = PerformCalcs(UP.PauseBeforeCalc)
  End If
End Function

Public Sub AddToArray(newval As Single, nary As Integer, ary() As Single)
'Add a new value to an array so that the array
'maintains an ascending sort. It is assumed that
'the existing array is already sorting in an
'ascending direction
'
' newval  - the value to be added to the array
' nary    - the number of elements currently in ary
' ary     - the existing array

  'array is empty
  If nary = 0 Then
    nary = 1
    ary(0) = newval
  'value is greater than last element
  ElseIf newval >= ary(nary - 1) Then
    nary = nary + 1
    ary(nary - 1) = newval
  Else
    For i = 0 To nary - 1
      If newval < ary(i) Then
        For j = nary - 1 To i Step -1
          ary(j + 1) = ary(j)
        Next
        ary(i) = newval
        nary = nary + 1
        Exit For
      End If
    Next
  End If
End Sub

Public Sub ArrayToField(fld As Field, X() As Single, Optional NumX)
'Store an array in a Longbinary database field
'
' fld must be a longbinary field of an open recordset
' X() contains the data to be packed into the field
' if NumX is provided, that number of array elements will
'   be transferred.
' If NumX is not provided, the entire X() array will be transferred
'
  Dim bb() As Byte
  Dim nb As Integer
  
  If IsMissing(NumX) Then
    nb = (UBound(X) + 1) * Len(X(0))
  Else
    nb = NumX * Len(X(0))
  End If
  ReDim bb(nb)                'make room for data
  CopyMemory bb(0), X(0), nb  'transfer data to byte array
  fld.AppendChunk bb()        'store in DB field
End Sub

Public Function FieldToArray(fld As Field, X() As Single) As Long
'Recover an array from a Longbinary database field
'
' fld must be a longbinary field of an open recordset
' X() must be dimensioned large enough to hold the array
'
' returns: number of array elements recovered
'
  Dim bb() As Byte
  Dim nb As Long
  
  FieldToArray = 0 'default return value
  
  'Longbinary fields seem to end with a null, which is included in the fieldsize.
  'Subtract it from the byte count
  nb = fld.FieldSize - 1
  
  If nb <= 0 Then Exit Function
  
  bb = fld.GetChunk(0, nb) 'retrieve raw data
  CopyMemory X(0), bb(0), nb 'transfer it to the array
  FieldToArray = nb / Len(X(0)) 'elements = bytes / bytes/element
End Function

Public Sub FormatAgreadMessage(istat As Long, _
                               itype As Long, _
                               adat() As Single, _
                               cdat As String, _
                               ShortMsg As String, _
                               LongMsg As String)
'Format short and long messages based on values returned by agread
'The short messages are to be added to the calc log
'The long messages are to be displayed in a MsgBox
'
  ' agread info: istat 0=ok: don't write data, keep reading
  '                    1=warning/info: write data, keep reading
  '                    2=error: write data, STOP
  '                    3=end of data
  '                    4=DropKick message, allow continue
  '                    5=DropKick message, do not allow continue
  '              idk   0=density mismatch in DropKick
  '                    1=speed mismatch in DropKick
  '                    2=Q mismatch in DropKick
  '                      Note: idk=2 is a special case. No re-
  '                      calculation is possible, so we can't
  '                      display a recompute button.
  '              itype 0=char data only
  '                    1=int and char data
  '                    2=real and char data
  '          adat/idat for istat=1-3: (0)=value (1)=min (2)=max
  '                    for istat=4-5: (0)=new value for DropKick
  '                                   (1)=AgDRIFT value
  '                                   (2)=DropKick value
  Dim Header As String
  
  ShortMsg = ""
  LongMsg = ""
  Select Case istat  'error level
  Case 0: 'ok
    Select Case itype
    Case 0: 'no message or data
    Case 1: 'message
      ShortMsg = Trim$(cdat)
    Case 2, 3: 'message and data
      ShortMsg = cdat & " " & AGFormat$(adat(0))
    End Select
  Case 1, 2: 'warning/error
    Select Case istat
    Case 1:
      Header = "Warning!"
    Case 2:
      Header = "Error!"
    End Select
    Select Case itype
    Case 0: 'no message or data
    Case 1: 'message
      ShortMsg = Trim$(cdat)
      AppendStr LongMsg, Header, True
      AppendStr LongMsg, Trim$(cdat), True
    Case 2, 3: 'message and float/int data
      ShortMsg = Trim$(cdat)
      AppendStr LongMsg, Header, True
      AppendStr LongMsg, Chr$(34) & Trim$(cdat) & Chr$(34), True
      AppendStr LongMsg, "is out of range. The limits are:", True
      AppendStr LongMsg, "", True
      AppendStr LongMsg, "Min: " & AGFormat$(adat(1)), True
      AppendStr LongMsg, "Val: " & AGFormat$(adat(0)), True
      AppendStr LongMsg, "Max: " & AGFormat$(adat(2)), True
    End Select
  Case 3: 'normal end of data
  Case 4, 5: 'Special DropKick message
    AppendStr LongMsg, "Warning!", True
    AppendStr LongMsg, Trim$(cdat), True
    AppendStr LongMsg, "", True
    AppendStr LongMsg, "    AgDRIFT:  " & AGFormat$(adat(1)), True
    AppendStr LongMsg, "    DropKick: " & AGFormat$(adat(2)), True
    AppendStr LongMsg, "", True
  End Select
End Sub

Public Function DSDIsUsed(xUD As UserData, iDSD As Integer) As Boolean
'Checks a DSD to see if any Nozzle points to it.
  Dim i As Integer
  DSDIsUsed = False 'default return value
  For i = 0 To UD.NZ.NumNoz - 1
    If UD.NZ.NozType(i) = iDSD Then
      DSDIsUsed = True
      Exit For
    End If
  Next
End Function

Public Function NumberOfDSDsUsed(xUD As UserData) As Integer
'Counts the number of DSD's currently attached to nozzles
  Dim nDSD As Integer
  Dim i As Integer
  
  nDSD = 0
  For iDSD = 0 To MAX_DSD - 1
    For i = 0 To UD.NZ.NumNoz - 1
      If UD.NZ.NozType(i) = iDSD Then
        nDSD = nDSD + 1
        Exit For
      End If
    Next
  Next
  NumberOfDSDsUsed = nDSD
End Function

Public Sub OpenFileAndSkipComments(fn As String, funit As Integer)
'Open file fn on unit funit and position it past any leading comment
'lines. Comment lines begin with a pound sign (#). Blank lines are
'also skipped.
  Dim icmnt As Integer
  Dim buf As String
  Dim i As Integer
  
  Open fn For Input As #funit
    
  'Read past (and count) any comment lines
  icmnt = -1
  Do
    icmnt = icmnt + 1
    Line Input #funit, buf
  Loop While (Left(buf, 1) = "#") Or (Trim(buf) = "")
    
  'Reopen the file to rewind it to the beginning
  Close #funit
  Open fn For Input As #funit
    
  'Read past comment lines
  For i = 1 To icmnt
    Line Input #funit, buf
  Next

End Sub

Public Sub EditSprayMaterial()

  Dim fSM As frmSprayMat
  Set fSM = New frmSprayMat
  With fSM
    'Send the form only those SM elements that it is interested in
    .SMType = UD.SM.Type
    .BasicType = UD.SM.BasicType
    .SMName = UD.SM.Name
    .SMLName = UD.SM.LName
    .CalcInputSelect = UD.SM.CalcInputSelect
    .NVFrac = UD.SM.NVFrac
    .ACFrac = UD.SM.ACFrac
    .ActSolFrac = UD.SM.ActSolFrac
    .AddSolFrac = UD.SM.AddSolFrac
    .ActNVFrac = UD.SM.ActNVFrac
    .AddNVFrac = UD.SM.AddNVFrac
    .FlowRate = UD.SM.FlowRate
    .FlowRateUnits = UD.SM.FlowRateUnits
    .SpecGrav = UD.SM.SpecGrav
    .NonVGrav = UD.SM.NonVGrav
    .EvapRate = UD.SM.EvapRate
    
    .Show vbModal
    
    If Not .Cancelled Then
      UD.SM.Type = .SMType
      UD.SM.BasicType = .BasicType
      UD.SM.Name = .SMName
      UD.SM.LName = .SMLName
      UD.SM.CalcInputSelect = .CalcInputSelect
      UD.SM.NVFrac = .NVFrac
      UD.SM.ACFrac = .ACFrac
      UD.SM.ActSolFrac = .ActSolFrac
      UD.SM.AddSolFrac = .AddSolFrac
      UD.SM.ActNVFrac = .ActNVFrac
      UD.SM.AddNVFrac = .AddNVFrac
      UD.SM.FlowRate = .FlowRate
      UD.SM.FlowRateUnits = .FlowRateUnits
      UD.SM.SpecGrav = .SpecGrav
      UD.SM.NonVGrav = .NonVGrav
      UD.SM.EvapRate = .EvapRate
    End If
  End With
  Unload fSM
  Set fSM = Nothing
End Sub
