C**AGKIRK
C  Continuum Dynamics, Inc.
C  AGDISP Version 8.11 08/08/04
C
      SUBROUTINE AGKIRK(BBK,IUNIT,LFL,ICLS,NPTS,DDV,XXV,IER,
     $                  REALWD,CHSTR,JCHSTR)
!MS$ATTRIBUTES DLLEXPORT,STDCALL :: AGKIRK
!MS$ATTRIBUTES REFERENCE :: BBK
!MS$ATTRIBUTES REFERENCE :: IUNIT
!MS$ATTRIBUTES REFERENCE :: LFL
!MS$ATTRIBUTES REFERENCE :: ICLS
!MS$ATTRIBUTES REFERENCE :: NPTS
!MS$ATTRIBUTES REFERENCE :: DDV
!MS$ATTRIBUTES REFERENCE :: XXV
!MS$ATTRIBUTES REFERENCE :: IER
!MS$ATTRIBUTES REFERENCE :: REALWD
!MS$ATTRIBUTES REFERENCE :: CHSTR
!MS$ATTRIBUTES REFERENCE :: JCHSTR
C
C  AGKIRK runs the USDA ARS analysis, then reconstructs the
C  drop size distribution by calling the appropriate function
C
C  BBK    - BKDATA data structure
C  IUNIT  - Units flag: 0 = English; 1 = metric
C  LFL    - Operations flag: 0 = initialization of calculation
C  ICLS   - Size class flag: -1 = no; 0-10 = class to use
C  NPTS   - Number of points in drop size distribution
C  DDV    - Drop size distribution array
C  XXV    - Volume fraction array
C  IER    - Error flag: 0 = no error -- result acceptable
C                       1 = warning with real data and character string
C                       2 = error with real data and character string
C                       3 = warning with character string only
C                       4 = error with character string only
C                       5 = information with character string
C  REALWD - Real data array (value, minimum, maximum)
C  CHSTR  - Character string
C  JCHSTR - Length of character string
C
      INCLUDE 'AGDSTRUC.INC'
C
      RECORD /BKDATA/ BBK
C
      CHARACTER*40 CHSTR
C
      DIMENSION REALWD(3),DDV(2),XXV(2)
      DIMENSION DV(15,22),OV(15,22),QV(15,22)
      DIMENSION SMNV(22),SMXV(22)
C
      DATA DV /
C 40 Degree Flat Fan (Large Orifice): Fixed Wing
     $  1406.353407,   -3.103646,   -6.520486,    0.523777,   -7.753689,
     $     0.01    ,    0.023333,   -0.000494,   -0.000906,   -0.003422,
     $    -0.000381,    0.011979,    0.026042,   -0.000226,    0.011393,
C 40 Degree Flat Fan (Small Orifice): Fixed Wing
     $   443.498698,   35.554688,    1.021528,   -0.134662,   -0.540799,
     $    -0.976562,   -0.0375  ,   -0.017037,    0.01087 ,   -0.001892,
     $    -0.000308,   -0.085938,   -0.000231,    0.001812,   -0.003038,
C 80 Degree Flat Fan: Fixed Wing
     $   325.817148,   44.61849 ,    1.014583,   -0.48479 ,    0.301179,
     $    -1.651042,   -0.019444,   -0.014774,    0.01721 ,   -0.002295,
     $     0.000142,   -0.097656,    0.000231,    0.002302,   -0.004901,
C CP-03: Fixed Wing
     $  1124.62045 , 1120.527   ,   -9.00689 ,    0.314185,   -4.27365 ,
     $ -5385.67    ,   17.61984 ,    0.031857,    1.71278 ,   -0.00562 ,
     $    -0.00012 ,   -5.87121 ,    0.013314,    0.0     ,    0.004919,
C CP-09: Fixed Wing
     $  1330.57402 , 2319.024   ,   -8.95069 ,    0.844018,   -8.70662 ,
     $ -9421.49    ,   10.54546 ,    0.288222,   -0.6917  ,   -0.00193 ,
     $     0.000322,   -1.42046 ,   -0.00941 ,   -0.00317 ,    0.016547,
C CP-11TT with Straight Stream Tips: Fixed Wing
     $  2381.64298 ,    8.830534,  -13.8717  ,    2.38559 ,  -18.731   ,
     $    -0.19856 ,    0.109206,    0.150833,   -0.01491 ,   -0.00018 ,
     $    -0.0003  ,   -0.00596 ,    0.039583,   -0.00747 ,    0.039804,
C Disc Orifice 46 Core: Fixed Wing
     $   950.045917,   -9.252604,   -4.717593,    0.07654 ,   -3.340676,
     $     0.119792,    0.140278,    0.00465 ,   -0.005435,   -0.001288,
     $    -0.00051 ,    0.023438,    0.011111,    0.001963,    0.000778,
C Disc Orifice 46 Core Ceramic: Fixed Wing
     $   915.44354 ,    9.503906,   -3.487963,    0.049215,   -4.438874,
     $    -0.567708,    0.075   ,    0.001502,    0.005435,   -0.000886,
     $    -0.000077,   -0.02474 ,    0.010185,    0.000302,    0.006908,
C Disc Orifice 56 Core: Fixed Wing
     $  1278.60178 ,  -33.410156,   -8.305324,    0.244263,   -4.688802,
     $     1.09375 ,    0.397222,    0.013272,    0.004529,   -0.002496,
     $    -0.000637,   -0.00651 ,    0.016435,    0.00151 ,    0.003689,
C Disc Orifice Straight Stream: Fixed Wing
     $  1995.949707,    8.654948,   -4.609896,    1.384322,  -13.96018 ,
     $    -1.898438,    0.125   ,    0.03    ,   -0.014493,   -0.006703,
     $    -0.000407,    0.095052,    0.018229,   -0.00268 ,    0.02436 ,
C Lund Straight Stream: Fixed Wing
     $  1753.847512,    9.617187,    2.747396,    1.236828,  -12.36748 ,
     $    -1.020833,   -0.1125  ,    0.100417,   -0.004529,   -0.002899,
     $     0.000225,    0.007813,   -0.014062,   -0.00434 ,    0.024595,
C 40 Degree Flat Fan (Large Orifice): Helicopter
     $   869.722456,   15.790616,    1.97536 ,   -0.988075,   -1.420385,
     $     0.019167,   -0.168333,   -0.041523,    0.031159,    0.003865,
     $    -0.001507,   -0.116922,    0.000677,    0.009748,   -0.013029,
C 40 Degree Flat Fan (Small Orifice): Helicopter
     $   170.396939,  101.040359,    3.558222,   -0.691382,    2.509991,
     $    -3.380208,   -0.168056,   -0.042325,    0.042572,    0.006361,
     $    -0.000608,   -0.358692,   -0.01784 ,    0.004084,   -0.013722,
C 80 Degree Flat Fan: Helicopter
     $   333.624667,   55.818812,    3.729231,   -0.891084,   -0.741711,
     $    -1.442708,   -0.101389,   -0.022634,    0.009511,    0.001691,
     $     0.000613,   -0.157254,   -0.018217,    0.00261 ,    0.003571,
C Accu-Flo Double Row: Helicopter
     $ -1402.0276  ,    154541.0,   57.08259 ,   -0.26556 ,   12.46182 ,
     $   -3163194.0, 1037.574   ,   -4.53501 ,   79.17658 ,   -0.00262 ,
     $    -0.0018  , -388.981   ,    0.141425,    0.000299,   -0.0537  ,
C Accu-Flo Single Row: Helicopter
     $  1692.795   ,     15328.0,    6.499909,    0.551652,  -15.6479  ,
     $    -139106.0, -840.861   ,    0.708405,  -15.058   ,    0.089221,
     $    -0.00044 ,    9.212425,    0.14999 ,   -0.0001  ,    0.026173,
C CP DR High Volume Flat Fan: Helicopter
     $   676.783765,   34.82421 ,   -1.58333 ,    0.27936 ,   -3.4311  ,
     $    -0.75    ,    0.06    ,    0.03875 ,   -0.03152 ,    0.007609,
     $     0.000144,    0.022345,    0.000795,    0.000797,   -0.00528 ,
C CP High Volume Flat Fan: Helicopter
     $   236.94974 ,   25.7901  ,   -2.1807  ,   -0.8573  ,    3.54164 ,
     $    -0.5421  ,    0.2225  ,   -0.0146  ,    0.00978 ,    0.01522 ,
     $    -0.0005  ,   -0.004   ,   -0.0134  ,    0.00669 ,   -0.0289  ,
C CP-03: Helicopter
     $    43.394159, 6927.85    ,   -1.22663 ,   -0.72402 ,    4.787058,
     $ -7324.39    ,  -19.5471  ,    0.031048,   -1.41634 ,    0.002147,
     $     0.000156,  -16.3291  ,   -0.02495 ,    0.004736,   -0.02066 ,
C Disc Orifice 46 Core: Helicopter
     $   702.764863,   14.55458 ,   -0.93185 ,   -0.43282 ,   -1.63986 ,
     $     0.427083,   -0.0375  ,   -0.00323 ,   -0.00045 ,    0.004227,
     $    -0.00078 ,   -0.10167 ,   -0.01529 ,    0.004886,    0.000215,
C Disc Orifice Straight Stream: Helicopter
     $   856.562144,   29.71657 ,    0.429309,   -3.11032 ,    6.344401,
     $    -0.8099  ,    0.0     ,    0.034167,    0.09692 ,   -0.00942 ,
     $     0.00163 ,   -0.33772 ,   -0.00712 ,    0.011611,   -0.03804 ,
C Raindrop RD: Helicopter
     $   322.370897,  250.1327  ,    1.865764,   -0.92638 ,   -5.07512 ,
     $   -11.6927  ,   -0.19375 ,   -0.00208 ,    0.045743,    0.0     ,
     $    0.0000678,   -0.54499 ,   -0.012   ,    0.004793,    0.013863/
C
      DATA OV /
C 40 Degree Flat Fan (Large Orifice): Fixed Wing
     $    -5.797369,   -0.909292,    0.211183,    0.020895,    0.088818,
     $     0.0575  ,   -0.010633,    0.000434,   -0.000053,    0.000025,
     $    -0.000025,   -0.007125,    0.000095,   -0.000032,    0.000218,
C 40 Degree Flat Fan (Small Orifice): Fixed Wing
     $    44.951753,    0.338893,   -0.58312 ,    0.044041,   -0.524545,
     $     0.133854,   -0.005069,    0.00178 ,   -0.000276,    0.000496,
     $    0.0000186,   -0.010352,    0.002593,   -0.000349,    0.001866,
C 80 Degree Flat Fan: Fixed Wing
     $    66.474074,   -2.966992,   -0.444416,    0.038658,   -0.634493,
     $     0.199922,   -0.006083,    0.0017  ,   0.0000498,    0.000227,
     $    0.0000076,   -0.001445,    0.002212,   -0.000279,    0.002051,
C CP-03: Fixed Wing
     $    87.353161,  135.4538  ,   -0.88393 ,    0.069071,   -0.94604 ,
     $   295.4545  ,   -2.66174 ,    0.005301,   -0.11792 ,    0.000437,
     $   0.00000302,   -0.15246 ,    0.002527,   -0.00042 ,    0.002837,
C CP-09: Fixed Wing
     $    38.771321, -146.588   ,    0.107187,    0.031674,   -0.40026 ,
     $   447.5207  ,   -0.58182 ,   -0.00589 ,   -0.02734 ,    0.000114,
     $    -0.000013,    0.293561,    0.000646,   -0.00014 ,    0.001207,
C CP-11TT with Straight Stream Tips: Fixed Wing
     $     0.965896,   -0.00456 ,    0.025662,    0.001441,   -0.01481 ,
     $     0.000114,   -0.0016  ,   0.0000758, 0.000000959,   0.0000181,
     $  0.000000818,    0.000109,   -0.000036,   -0.000012,   0.0000507,
C Disc Orifice 46 Core: Fixed Wing
     $    44.779209,    0.597786,   -0.21876 ,    0.048405,   -0.587944,
     $     0.082474,   -0.018597,    0.000564,    0.002255,    0.000086,
     $     0.000083,   -0.010391,    0.002251,   -0.000585,    0.002195,
C Disc Orifice 46 Core Ceramic: Fixed Wing
     $     6.342202,   -0.027227,   -0.044772,    0.003868,   -0.067457,
     $    -0.00263 ,    0.000264,    0.000287,   -0.000127,    0.00003 ,
     $     0.000006,    0.000456,    0.000152,   -0.00004 ,    0.000197,
C Disc Orifice 56 Core: Fixed Wing
     $    23.387121,    1.073008,   -0.191289,    0.051968,   -0.386102,
     $     0.021615,   -0.016778,    0.000644,   -0.000063,    0.000122,
     $     0.000057,   -0.005716,    0.0019  ,   -0.000452,    0.001504,
C Disc Orifice Straight Stream: Fixed Wing
     $    15.736694,   -0.038242,   -0.108667,    0.020872,   -0.219267,
     $     0.048099,    0.00225 ,   -0.000029,   -0.000462,    0.000629,
     $     0.00003 ,   -0.003424,   -0.000167,   -0.000244,    0.000999,
C Lund Straight Stream: Fixed Wing
     $    18.76852 ,    0.213646,   -0.340042,    0.023953,   -0.256653,
     $    -0.024167,    0.039875,    0.002021,    0.000525,    0.000647,
     $     0.000015,   -0.001979,   -0.000875,   -0.000257,    0.001142,
C 40 Degree Flat Fan (Large Orifice): Helicopter
     $     1.858916,    0.060585,   -0.027956,   -0.01048 ,   -0.023231,
     $    -0.001812,   0.0000444,    0.000235,   -0.000058, -0.00000322,
     $    0.0000218,    0.000331,    0.000245,-0.000000717,   0.0000784,
C 40 Degree Flat Fan (Small Orifice): Helicopter
     $     2.221465,   -0.113568,   -0.039841,    0.001702,   -0.03456 ,
     $     0.006198,   -0.000861,    0.000248,   0.0000498,  0.00000242,
     $  -0.00000103,    0.000201,    0.000364,  -0.0000131,    0.000142,
C 80 Degree Flat Fan: Helicopter
     $     1.920302,    0.633599,   -0.09265 ,    0.029997,   -0.151448,
     $    -0.019193,    0.001389, 0.000059465,   -0.000367,-0.000000805,
     $ -0.000052094,   -0.00467 ,    0.00096 , 0.000027708,    0.000902,
C Accu-Flo Double Row: Helicopter
     $     3.745326, -323.093   ,    0.002176,    0.005567,   -0.03859 ,
     $  5603.299   ,    4.60859 ,    0.012022,   -0.06167 ,   -0.000059,
     $   -0.0000045,    0.848229,   -0.00335 ,   -0.00002 ,    0.000378,
C Accu-Flo Single Row: Helicopter
     $     1.605911,  -38.3924  ,    0.161377,    0.001961,   -0.04885 ,
     $   109.7274  ,    0.964143,   -0.00828 ,   -0.00599 ,   -0.00012 ,
     $  0.000000331,    0.441192,   -0.00064 ,   -0.000017,    0.000321,
C CP DR High Volume Flat Fan: Helicopter
     $     0.8812  ,   -0.1297  ,    0.047597,   -0.00237 ,    0.005874,
     $     0.003783,   -0.00075 ,   -0.00108 ,    0.000264,   0.0000163,
     $    0.0000169,   -0.00074 ,   -0.00026 ,   -0.00011 ,    0.000316,
C CP High Volume Flat Fan: Helicopter
     $    -5.562786,    0.15509 ,    0.239245,    0.049494,   -0.00219 ,
     $     0.002821,    0.000075,   -0.00664 ,   -0.00053 ,   -0.00046 ,
     $    0.0000391,   -0.00216 ,   -0.00041 ,   -0.0004  ,    0.000686,
C CP-03: Helicopter
     $     3.468235,   -1.40676 ,   -0.07961 ,   -0.00626 ,   -0.01417 ,
     $   150.2469  ,   -0.5424  ,    0.001087,    0.041173,    0.000149,
     $   0.00000293,   -0.17534 ,    0.000208,   -0.000017,    0.000177,
C Disc Orifice 46 Core: Helicopter
     $     5.289912,   -0.53832 ,   -0.08948 ,   -0.02697 ,    0.051605,
     $     0.020417,    0.001194,    0.000277,    0.000657,   0.0000511,
     $    0.0000357,   -0.00049 ,    0.001143,   0.0000185,   -0.00032 ,
C Disc Orifice Straight Stream: Helicopter
     $     1.756453,   -0.10005 ,    0.029742,    0.008732,   -0.03522 ,
     $     0.002708,    0.00025 ,   -0.00265 ,   -0.00049 ,   0.0000833,
     $   -0.0000076,    0.001199,    0.000332,   -0.000013,    0.000157,
C Raindrop RD: Helicopter
     $     6.365973,   -1.3554  ,   -0.00473 ,    0.005745,   -0.04197 ,
     $     0.070599,    0.000438,   -0.0006  ,   -0.00031 ,  0.00000906,
     $   0.00000136,    0.003064,    0.000101,   -0.000056,    0.000273/
C
      DATA QV /
C 40 Degree Flat Fan (Large Orifice): Fixed Wing
     $    26.954164,   -0.775021,   -0.19115 ,   -0.011394,   -0.20974 ,
     $     0.049429,   -0.008794,    0.002812,    0.000196,    0.000356,
     $     0.000081,   -0.007146,    0.00141 ,   -0.000272,    0.001317,
C 40 Degree Flat Fan (Small Orifice): Fixed Wing
     $    89.055169,   -0.639206,   -0.789293,    0.074857,   -1.024751,
     $     0.303724,   -0.010847,    0.004422,   -0.000901,    0.000536,
     $     0.000125,   -0.01694 ,    0.003172,   -0.00082 ,    0.003876,
C 80 Degree Flat Fan: Fixed Wing
     $   139.264855,   -9.005273,   -0.748249,    0.086956,   -1.251807,
     $     0.524557,   -0.001847,    0.004438,   -0.000362,    0.000447,
     $    0.0000574,    0.003086,    0.002591,   -0.000736,    0.004256,
C CP-03: Fixed Wing
     $   126.273214,  -12.5953  ,   -0.88346 ,    0.109752,   -1.435   ,
     $  1247.658   ,   -4.71562 ,    0.005575,   -0.45191 ,    0.000445,
     $    0.0000597,    0.357008,    0.004381,   -0.00064 ,    0.004349,
C CP-09: Fixed Wing
     $    86.602939, -338.829   ,    0.384178,    0.072213,   -0.90086 ,
     $  1052.204   ,   -1.05143 ,   -0.01904 ,   -0.06192 ,  -0.0000029,
     $    -0.000016,    0.660985,    0.001953,   -0.00038 ,    0.002843,
C CP-11TT with Straight Stream Tips: Fixed Wing
     $    20.040108,   -0.51892 ,   -0.03864 ,    0.030003,   -0.24416 ,
     $     0.008485,   -0.00559 ,    0.000667,    0.000107,   0.0000236,
     $    0.0000293,    0.001742,    0.000859,   -0.00028 ,    0.000895,
C Disc Orifice 46 Core: Fixed Wing
     $    79.863025,   -0.136133,   -0.016485,    0.055149,   -1.046752,
     $     0.135391,   -0.034528,    0.000532,    0.002631,    0.000035,
     $     0.000143,   -0.008346,    0.002541,   -0.000846,    0.003901,
C Disc Orifice 46 Core Ceramic: Fixed Wing
     $    29.533065,   -0.088659,   -0.251662,    0.047475,   -0.422754,
     $     0.167578,   -0.023319,    0.001852,   -0.000829,   -0.000012,
     $     0.000057,   -0.006654,    0.002282,   -0.000406,    0.001668,
C Disc Orifice 56 Core: Fixed Wing
     $    41.645922,    0.420755,   -0.012455,    0.106532,   -0.754806,
     $     0.071615,   -0.035458,    0.000662,   -0.000856,    0.000093,
     $     0.00012 ,   -0.001172,    0.002297,   -0.000931,    0.003074,
C Disc Orifice Straight Stream: Fixed Wing
     $    39.980481,   -1.648854,   -0.316177,    0.052116,   -0.480452,
     $     0.156146,    0.007937,   -0.000217,   -0.001236,    0.000879,
     $     0.000077,   -0.003021,    0.00051 ,   -0.000588,    0.002147,
C Lund Straight Stream: Fixed Wing
     $    47.034529,   -1.776198,   -0.800807,    0.073721,   -0.543231,
     $     0.011667,    0.075125,   -0.000646,   -0.000408,    0.000658,
     $     0.000029,    0.004948,    0.000224,   -0.000574,    0.00222 ,
C 40 Degree Flat Fan (Large Orifice): Helicopter
     $    10.925165,    0.262928,   -0.174395,   -0.037241,   -0.176133,
     $    -0.007542,    0.000933,    0.001199,   -0.000543,   -0.000156,
     $     0.000114,    0.001907,    0.002079,   -0.000126,    0.000857,
C 40 Degree Flat Fan (Small Orifice): Helicopter
     $    15.682473,   -1.254278,   -0.266666,    0.034122,   -0.257573,
     $     0.154766,   -0.011403,    0.001662,   -0.001042,   -0.000145,
     $   -0.0000204,   -0.007311,    0.003429,  -0.0000931,    0.001366,
C 80 Degree Flat Fan: Helicopter
     $    16.524193,   -1.722317,   -0.36483 ,    0.073179,   -0.225148,
     $     0.146589,    0.005931,    0.001915,   -0.003274,-0.000016103,
     $ -0.000083141,   -0.004961,    0.001945, 0.000030358,    0.001217,
C Accu-Flo Double Row: Helicopter
     $    11.405503, -859.24    ,    0.095434,    0.011258,   -0.13254 ,
     $ 16072.0     ,    1.838474,    0.020774,   -0.2616  ,   -0.000091,
     $    -0.000003,    3.148989,   -0.00576 ,   -0.000053,    0.0009  ,
C Accu-Flo Single Row: Helicopter
     $     2.40485 ,  -75.7485  ,    0.243402,    0.008437,   -0.09361 ,
     $   298.9505  ,    2.390061,   -0.00294 ,   -0.00489 ,   -0.00049 ,
     $   -0.0000041,    0.95628 ,   -0.00284 ,   -0.000044,    0.000745,
C CP DR High Volume Flat Fan: Helicopter
     $    -1.150146,    0.228227,    0.093437,   -0.00929 ,    0.029175,
     $    -0.00557 ,   -0.0012  ,   -0.00162 ,    0.000652,   -0.000013,
     $     0.000031,   -0.00338 ,   -0.00063 ,   -0.00019 ,    0.000689,
C CP High Volume Flat Fan: Helicopter
     $     3.358808,   -0.31737 ,    0.226687,    0.076403,   -0.07196 ,
     $     0.013379,    0.00085 ,   -0.00671 ,   -0.00091 ,   -0.00082 ,
     $     0.000039,   -0.00212 ,   -0.00023 ,   -0.00052 ,    0.001172,
C CP-03: Helicopter
     $     7.542536,  -30.4804  ,   -0.11007 ,    0.010938,   -0.04765 ,
     $   455.3065  ,   -0.9343  ,    0.001715,    0.009223,    0.000211,
     $   -0.0000064,   -0.54729 ,    0.000475,   -0.000078,    0.000622,
C Disc Orifice 46 Core: Helicopter
     $    11.955678,   -0.9381  ,   -0.19409 ,   -0.03258 ,    0.014496,
     $     0.026849,    0.001653,    0.000641,    0.001091,    0.000106,
     $    0.0000444,   -0.00012 ,    0.002252, -0.00000031,   -0.00015 ,
C Disc Orifice Straight Stream: Helicopter
     $     4.194969,   -0.40947 ,    0.007556,    0.019413,   -0.07618 ,
     $     0.017708,   -0.00206 ,   -0.00152 ,   -0.00096 ,    0.000183,
     $   -0.0000097,    0.002606,    0.000592,   -0.000075,    0.000427,
C Raindrop RD: Helicopter
     $    16.358272,    -3.19978,   -0.02587 ,    0.00777 ,   -0.0988  ,
     $     0.165104,    0.002625,   -0.00077 ,   -0.0004  ,  0.00000725,
     $   0.00000481,    0.005725,    0.000197,   -0.00011 ,    0.00064 /
C
      DATA SMNV / 11*42.5, 2*13.9, 9*12.5 /
      DATA SMXV / 11*75.0, 2*33.3, 9*47.0 /
C
C  Set all of the necessary parameters for USDA ARS regressions
C
      IER=0
C
      N=BBK.NOZTYPE+1
      O=BBK.ORIFICE
      A=BBK.NOZANGLE
C
      S=BBK.SPEED
      IF (LFL.EQ.0) THEN
        LFL=1
        FAC=1.0
        IF (IUNIT.EQ.0) FAC=1.0/0.447
        CALL AGCHK(S,SMNV(N),SMXV(N),3,SMNV(N),SMXV(N),IER,FAC,REALWD)
        IF (IER.GT.0) THEN
          IF (IUNIT.EQ.0) THEN
            CHSTR='Air Speed (mph)'
          ELSE
            CHSTR='Air Speed (m/s)'
          ENDIF
          JCHSTR=15
          RETURN
        ENDIF
      ENDIF
      S=3.6*S
C
      P=BBK.PRESSURE
      IF (LFL.EQ.1) THEN
        LFL=2
        FAC=1.0
        IF (IUNIT.EQ.0) FAC=14.5
        CALL AGCHK(P,1.37,4.15,3,1.37,4.15,IER,FAC,REALWD)
        IF (IER.GT.0) THEN
          IF (IUNIT.EQ.0) THEN
            CHSTR='Pressure (psig)'
            JCHSTR=15
          ELSE
            CHSTR='Pressure (bar)'
            JCHSTR=14
          ENDIF
          RETURN
        ENDIF
      ENDIF
      P=100.0*P
C
C  Compute the USDA ARS parameters and construct the distribution
C
      IF (LFL.EQ.2) THEN
        IF (N.EQ.1.OR.N.EQ.2.OR.N.EQ.12.OR.N.EQ.13) O=O-4000.0
        IF (N.EQ.3.OR.N.EQ.14.OR.N.EQ.17.OR.N.EQ.18) O=O-8000.0
        VMDNEW=DV(1,N)+DV(2,N)*O+DV(3,N)*A
     $        +DV(4,N)*P+DV(5,N)*S+DV(6,N)*O*O+DV(7,N)*O*A
     $        +DV(8,N)*A*A+DV(9,N)*O*P+DV(10,N)*A*P+DV(11,N)*P*P
     $        +DV(12,N)*O*S+DV(13,N)*A*S+DV(14,N)*P*S+DV(15,N)*S*S
        V1NEW=OV(1,N)+OV(2,N)*O+OV(3,N)*A
     $        +OV(4,N)*P+OV(5,N)*S+OV(6,N)*O*O+OV(7,N)*O*A
     $        +OV(8,N)*A*A+OV(9,N)*O*P+OV(10,N)*A*P+OV(11,N)*P*P
     $        +OV(12,N)*O*S+OV(13,N)*A*S+OV(14,N)*P*S+OV(15,N)*S*S
        V2NEW=QV(1,N)+QV(2,N)*O+QV(3,N)*A
     $        +QV(4,N)*P+QV(5,N)*S+QV(6,N)*O*O+QV(7,N)*O*A
     $        +QV(8,N)*A*A+QV(9,N)*O*P+QV(10,N)*A*P+QV(11,N)*P*P
     $        +QV(12,N)*O*S+QV(13,N)*A*S+QV(14,N)*P*S+QV(15,N)*S*S
C
C  Compute with %V<100 and %V<200, along with DV0.5
C
        Y1=SQRT(100.0/VMDNEW)
        X1=FX(0.01*V1NEW)
        Y2=SQRT(200.0/VMDNEW)
        X2=FX(0.01*V2NEW)
        NP=3
        SUMX=X1+X2
        SUMXY=X1*Y1+X2*Y2
        SUMXX=X1*X1+X2*X2
        SUMY=1.0+Y1+Y2
        SL=(SUMXY-SUMX*SUMY/NP)/(SUMXX-SUMX*SUMX/NP)
        RSNEW=5.126915*SL
C
        LTEM=BBK.SPRTYPE
        ITEM=BBK.SPECSRC
        CALL AGPARX(LTEM,ICLS,ITEM,VMDNEW,RSNEW,NPTS,DDV,XXV)
      ENDIF
      RETURN
      END