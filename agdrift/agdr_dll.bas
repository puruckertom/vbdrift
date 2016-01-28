Attribute VB_Name = "basAGDR_DLL"
'$Id: agdr_dll.bas,v 1.10 2008/10/22 17:26:06 tom Exp $
'API for agdrift32.dll

Option Explicit

'
' Globals
'

'AGINIT MAA flags
Public Const AGINIT_NORMAL = 0
Public Const AGINIT_MAA = 1

'AGENDS results flags
Public Const AGENDS_DEPOS = 0      'Deposition
Public Const AGENDS_PID = 1        'Pond-Integrated Deposition
Public Const AGENDS_FLUX = 2       'Vertical Flux
Public Const AGENDS_1HRCON = 3     '1-hour average concentration
Public Const AGENDS_COV = 4        'COV
Public Const AGENDS_MEAN = 5       'COV Mean Deposition
Public Const AGENDS_ALOFT = 6      'Fraction Aloft
Public Const AGENDS_SGLDEP = 7     'Single-Swath Deposition
Public Const AGENDS_SGLHAF = 8     'Single-Swath upwind half-boom deposition
Public Const AGENDS_MULDEP = 9     'Multiple deposition setup
Public Const AGENDS_SBLOCK = 10    'Spray block deposition
Public Const AGENDS_CANOPY = 11    'Canopy deposition
Public Const AGENDS_TAALOFT = 12   'Time accountancy aloft
Public Const AGENDS_TAVAPOR = 13   'Time accountancy vapor
Public Const AGENDS_TACANOPY = 14  'Time accountancy canopy
Public Const AGENDS_TAGROUND = 15  'Time accountancy ground
Public Const AGENDS_HAALOFT = 16   'Height accountancy aloft
Public Const AGENDS_HAVAPOR = 17   'Height accountancy vapor
Public Const AGENDS_HACANOPY = 18  'Height accountancy canopy
Public Const AGENDS_SBLOCKDSD = 19 'Spray block drop size distribution
Public Const AGENDS_DWINDDSD = 20  'Downwind drop size distribution
Public Const AGENDS_FLUXDSD = 21   'Vertical flux drop size distribution
Public Const AGENDS_DAALOFT = 22   'Time accountancy aloft
Public Const AGENDS_DAVAPOR = 23   'Time accountancy vapor
Public Const AGENDS_DACANOPY = 24  'Time accountancy canopy
Public Const AGENDS_DAGROUND = 25  'Time accountancy ground
Public Const AGENDS_SBCOVER = 26   'Spray Block Area Coverage
Public Const AGENDS_CANDSD = 27    'Canopy drop size distribution
Public Const AGENDS_LAYOUT = 28    'Application Layout

'AGGRND results flags
Public Const AGGRND_DEPOS = 0      'Deposition
Public Const AGGRND_PID = 1        'Pond-Integrated Deposition

'
' Function Declarations
'
Declare Sub agarea Lib "agdrift32.dll" Alias "_agarea@40" ( _
  nxpts&, nypts&, xgrdv!, ygrdv!, dgrdv!, _
  nacb&, xacb!, yacb!, area!, cover!)
Declare Sub agaver Lib "agdrift32.dll" Alias "_agaver@16" ( _
  npts&, dv!, dmin!, dav!)
Declare Sub agcov Lib "agdrift32.dll" Alias "_agcov@32" ( _
  ncov&, COVVal!, COVESW!, COVDep!, _
  INTYPE&, COV!, ESW!, dep!)
Declare Sub agdrin Lib "agdrift32.dll" Alias "_agdrin@36" ( _
  nd&, Typ&, X!, Y!, Z!, XN!, YN!, ZN!, Size!)
Declare Sub agdrop Lib "agdrift32.dll" Alias "_agdrop@4" (n&)
Declare Sub agdrot Lib "agdrift32.dll" Alias "_agdrot@4" (dep!)
Declare Sub agdrp Lib "agdrift32.dll" Alias "_agdrp@32" ( _
  UD As UserData, DiamIni!, RelHgt!, DiamFin!, _
  Dist!, TimeImpact!, _
  ByVal cdat$, clen&)
Declare Sub agdsrn Lib "agdrift32.dll" Alias "_agdsrn@40" ( _
  lflg&, nusr&, dkv!, xkv!, VMD!, xrs!, _
  D10!, D90!, F141!, DP!)
Declare Sub agends Lib "agdrift32.dll" Alias "_agends@16" ( _
  iflg&, nv&, yv!, dv!)
Declare Sub agfill Lib "agdrift32.dll" Alias "_agfill@40" ( _
  itype&, nusr&, div!, xiv!, _
  npts&, dv!, xv!, _
  ier&, ByVal chstr$, jchstr&)
Declare Sub aggrnd Lib "agdrift32.dll" Alias "_aggrnd@32" ( _
  ityp&, itier&, xdwnd!, iswth&, IDEP&, _
  np&, yv!, dv!)
Declare Sub aginit Lib "agdrift32.dll" Alias "_aginit@8" ( _
  UD As UserData, maaflg&)
Declare Sub agkick Lib "agdrift32.dll" Alias "_agkick@44" ( _
  DK As DropKickData, iunit&, lfl&, iqual&, _
  npts&, dv!, xv!, _
  ier&, realwd!, ByVal chstr$, jchstr&)
Declare Sub agkirk Lib "agdrift32.dll" Alias "_agkirk@44" ( _
  BK As DropKirkData, iunit&, lfl&, iqual&, _
  npts&, dv!, xv!, _
  ier&, realwd!, ByVal chstr$, jchstr&)
Declare Sub aglibr Lib "agdrift32.dll" Alias "_aglibr@24" ( _
  itier&, IDEP&, dwnd!, nptsd&, ydv!, ddv!)
Declare Sub aglims Lib "agdrift32.dll" Alias "_aglims@12" ( _
  np&, dv!, pv!)
Declare Sub agnozl Lib "agdrift32.dll" Alias "_agnozl@24" ( _
  nvar&, yv!, sspan!, boomwid!, nnew&, av!)
Declare Sub agnums Lib "agdrift32.dll" Alias "_agnums@32" ( _
  xnsd!, xcov!, xSM!, xae!, XDE!, xab!, xev!, xcn!)
Declare Sub agorch Lib "agdrift32.dll" Alias "_agorch@36" ( _
  ityp&, itier&, xdwnd!, ibtrow&, ietrow&, IDEP&, _
  npts&, yv!, dv!)
Declare Sub agparm Lib "agdrift32.dll" Alias "_agparm@32" ( _
  lflg&, lcls&, lsrc&, VMD!, xrs!, nusr&, dkv!, xkv!)
Declare Sub agread Lib "agdrift32.dll" Alias "_agread@28" ( _
  iunits&, stat&, idk&, ityp&, _
  adat!, ByVal cdat$, clen&)
Declare Sub agrot Lib "agdrift32.dll" Alias "_agrot@44" ( _
  HK As HKData, iunit&, lfl&, icls&, npts&, dv!, xv!, _
  ier&, realwd!, ByVal chstr$, jchstr&)
Declare Sub agrtrn Lib "agdrift32.dll" Alias "_agrtrn@8" ( _
  xh!, yv!)
Declare Sub agsbck Lib "agdrift32.dll" Alias "_agsbck@16" ( _
  nbnd&, xbnd!, ybnd!, iflg&)
Declare Sub agsbin Lib "agdrift32.dll" Alias "_agsbin@52" ( _
  UD As UserData, _
  nbnd&, xbnd!, ybnd!, fdir!, ncon&, conv!, _
  lnm&, ldn&, lpt&, npts&, dv!, pv!)
Declare Sub agsblk Lib "agdrift32.dll" Alias "_agsblk@68" ( _
  UD As UserData, _
  nsgl&, SglDist!, SglVal!, SglHalf!, _
  IDEP&, INTYPE&, XLENG!, XDEEP!, _
  XACT!, XAPPL!, XDEPS!, XDEPD!, XCONC!, _
  npts&, sv!, bv!)
Declare Sub agsend Lib "agdrift32.dll" Alias "_agsend@28" ( _
  nxpts&, nypts&, xgrdv!, ygrdv!, dgrdv!, conv!, dmax!)
Declare Sub agsetl Lib "agdrift32.dll" Alias "_agsetl@24" ( _
  xSD As SprayMaterialData, npts&, dtv!, stv!, dnv!, snv!)
Declare Sub agsgrd Lib "agdrift32.dll" Alias "_agsgrd@24" ( _
  nxpts&, nypts&, nsflt&, ysflt!, xsbeg!, xsend!)
Declare Sub agsmck Lib "agdrift32.dll" Alias "_agsmck@44" ( _
  itier&, TEMPA!, RHUMA!, _
  NEVNTS&, NYEARS&, PROB!, NTSPD&, _
  ier&, realwd!, ByVal chstr$, jchstr&)
Declare Sub agsmex Lib "agdrift32.dll" Alias "_agsmex@20" ( _
  nexct&, npts&, yv!, dv!, rar!)
Declare Sub agsmpl Lib "agdrift32.dll" Alias "_agsmpl@40" ( _
  npts&, yv!, dv!, nexam&, _
  ndep&, DepDist!, DepVal!, _
  npid&, PIDDist!, PIDVal!)
Declare Sub agsmti Lib "agdrift32.dll" Alias "_agsmti@16" ( _
  iappl&, ndep&, DepDist!, DepVal!)
Declare Sub agstrm Lib "agdrift32.dll" Alias "_agstrm@140" ( _
  UD As UserData, _
  nsgl&, SglDist!, SglVal!, SglHalf!, _
  ISTYPE&, INTYPE&, XWIDE!, XDEEP!, _
  XACT!, XDIST!, XSRATE!, XSLENG!, XSTURN!, _
  XRIPAR!, XDECAY!, XCHARG!, XINPTS!, _
  iunit&, lfl&, xsngl!, Nauto&, Xauto!, Rauto!, _
  npts&, yv!, cv!, _
  NSBL&, TTV!, XXV!, CCV!, _
  ier&, realwd!, ByVal chstr$, jchstr&)
Declare Sub agterr Lib "agdrift32.dll" Alias "_agterr@76" ( _
  UD As UserData, _
  ndep&, DepDist!, DepVal!, _
  npid&, PIDDist!, PIDVal!, _
  ISTYPE&, INTYPE&, XLENG!, _
  XACT!, XLAND!, XAPPL!, XDEPS!, XDEPD!, XCONC!, _
  nusr&, UsrDist!, UsrVal!)
Declare Sub agtox Lib "agdrift32.dll" Alias "_agtox@80" ( _
  UD As UserData, _
  ndep&, DepDist!, DepVal!, _
  npid&, PIDDist!, PIDVal!, _
  ISTYPE&, INTYPE&, LPOND!, DPOND!, _
  XACT!, XPOND!, XAPPL!, XDEPS!, XDEPD!, XCONC!, _
  nusr&, UsrDist!, UsrVal!)
Declare Sub agtraj Lib "agdrift32.dll" Alias "_agtraj@24" ( _
  UD As UserData, drop!, NTR&, slv!, apv!, japv&)
Declare Sub agtrgo Lib "agdrift32.dll" Alias "_agtrgo@12" ( _
  ntrgo&, apv!, japv&)
Declare Sub agupds Lib "agdrift32.dll" Alias "_agupds@8" ( _
  UD As UserData, nd&)
Declare Sub agwdrs Lib "agdrift32.dll" Alias "_agwdrs@40" ( _
  nfldir&, temp!, rhum!, nxspd&, nfreq%, _
  monb&, mone&, tempg!, rhumg!, PROB!)
Declare Sub agwplt Lib "agdrift32.dll" Alias "_agwplt@32" ( _
  PROB!, npts&, deg!, p10!, p30!, p50!, p70!, p90!)

