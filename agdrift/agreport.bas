Attribute VB_Name = "basAGREPORT"
'agreport.bas - report-generation formatting routines
'$Id: agreport.bas,v 1.3 2001/08/13 17:40:00 tom Exp $
Option Explicit

'The general form of the report output consists of three columns;
'the first generally contains a descriptive name, the second contains
'the current value of the quantity displayed, while the third contains
'the default value for reference if the current value differs from it.
'For tables, where more than one quantity is to be displayed on a line,
'"subcolumns" are defined to contain multiple values in a single main
'column.
Const c1wid = 26 'number of columns to allot for column 1
Const c2wid = 28 'number of columns to allot for column 2
Const c3wid = 28 'number of columns to allot for column 3
Private c1fmt As String  'format string for column 1
Private c2fmt As String  'format string for column 2
Private c3fmt As String  'format string for column 3

Public Function GenReportText() As String
'Generate the text of printed reports for:
'  Input Summary
'  Print Preview
'  Print Report
'The report lists the user-defined inputs in a three-column format.
'Column 1 describes the value and its units, if applicable.
'Column 2 lists the user-entered value
'Column 3 lists the default value if the user-entered value is different
  
  Dim grt As String    'temporary storage for report text
  Dim s As String      'workspace string
  Dim s1 As String, s2 As String
  Dim xUD As UserData  'default user data values
  Dim nloop As Integer 'local loop counter
  Dim start As Integer
  Dim i As Integer
  Dim iDSD As Integer  'DSD index
  Dim Col2Hdr As String, Col3Hdr As String

  'get a set of default data for comparisons
  UserDataDefault xUD

  'Set up the formats for the three columns and subcolumns
  'Try to keep the total under 80 columns for printing
  c1fmt = "!" & String$(c1wid, "@") 'left-justified
  c2fmt = " " & String$(c2wid, "@") '1 space, right-justified
  c3fmt = " " & String$(c3wid, "@") '1 space, right-justified
  'the following produces this: ---Current--- ---Default---
  s = "Current": i = (c2wid - Len(s)) / 2
  Col2Hdr = String(i, "-") + s + String(c2wid - i - Len(s), "-")
  s = "Default": i = (c2wid - Len(s)) / 2
  Col3Hdr = String(i, "-") + s + String(c2wid - i - Len(s), "-")
  
  grt = "" 'start with a blank string
    
  'Report title
  AppendStr grt, "AgDRIFT® Input Data Summary", True
  AppendStr grt, "", True
  
  'General data for all tiers
  AppendStr grt, "--General--", True
  AppendStr grt, "Tier: " & String$(UD.Tier, "I"), True
  AppendStr grt, "Title: " & UD.Title, True
  AppendStr grt, "Notes: ", True
  If Len(UD.Notes) > 0 Then
    start = 1
    Do
      AppendStr grt, " " & LineFromString(UD.Notes, start), True
    Loop While start > 0
  End If
  AppendStr grt, "", True
  
  'Calculation status
  If UD.Tier > TIER_1 Then
    AppendStr grt, "Calculations Done: " & Format$(UC.Valid, "YES/NO"), True
    AppendStr grt, "Run ID: " & GetRunID(), True
    AppendStr grt, "", True
  End If

  'Column headings for all Tiers
  AppendStr grt, "Default values appear when they differ from the Current values.", True
  AppendStr grt, "", True

  'Tier I
  If UD.Tier = TIER_1 Then
    AppendReportLineS grt, " ", Col2Hdr, Col3Hdr
    AppendReportLineS grt, "Application Method", GetBasicNameAM(UD.ApplMethod), GetBasicNameAM(xUD.ApplMethod)
    Select Case UD.ApplMethod
    Case AM_AERIAL 'aerial
      AppendReportLineS grt, "Application Selection", _
        GetBasicNameDSD(UD.DSD(0).BasicType), GetBasicNameDSD(xUD.DSD(0).BasicType)
    Case AM_GROUND 'ground
      AppendReportLineS grt, "Application Selection", _
        GetBasicNameGA(UD.GA.BasicType), GetBasicNameDSD(xUD.GA.BasicType)
'tbc this relies on form controls, change UD ta accomodate extended settings
      If frmTier1Gnd!chkExtended.Value = 1 Then
        AppendReportLineN grt, "Number of Swaths", UN_NONE, UD.GA.NumSwaths, xUD.GA.NumSwaths
      End If
    Case AM_ORCHARD 'orchard
      AppendReportLineS grt, "Application Selection", _
        GetBasicNameOA(UD.OA.BasicType), GetBasicNameDSD(xUD.OA.BasicType)
'tbc this relies on form controls, change UD ta accomodate extended settings
      If frmTier1orc!chkExtended.Value = 1 Then
        AppendReportLineN grt, "Starting Tree Row", UN_NONE, UD.OA.BegTrow, xUD.OA.BegTrow
        AppendReportLineN grt, "Ending Tree Row", UN_NONE, UD.OA.EndTrow, xUD.OA.EndTrow
      End If
    End Select
    AppendStr grt, "", True
  End If
    
  'Aircraft
  Select Case UD.Tier
  Case TIER_2
    AppendReportLineS grt, "--Aircraft--", Col2Hdr, Col3Hdr
    AppendReportLineS grt, "Name", UD.AC.Name, xUD.AC.Name
    AppendReportLineS grt, "Type", GetBasicNameAC2(UD.AC.BasicType), GetBasicNameAC2(xUD.AC.BasicType)
    AppendReportLineN grt, "Boom Length (%)", UN_NONE, UD.NZ.BoomWidth, xUD.NZ.BoomWidth
    AppendReportLineN grt, "Boom Height", UN_LENGTH, UD.CTL.Height, xUD.CTL.Height
    AppendReportLineN grt, "Flight Lines", UN_NONE, UD.CTL.NumLines, xUD.CTL.NumLines
    AppendStr grt, "", True
  Case TIER_3
    AppendReportLineS grt, "--Aircraft--", Col2Hdr, Col3Hdr
    AppendReportLineS grt, "Name", UD.AC.Name, xUD.AC.Name
    AppendReportLineS grt, "Type", GetTypeNameAC(UD.AC.Type), GetTypeNameAC(xUD.AC.Type)
    AppendReportLineN grt, "Boom Height", UN_LENGTH, UD.CTL.Height, xUD.CTL.Height
    AppendReportLineN grt, "Flight Lines", UN_NONE, UD.CTL.NumLines, xUD.CTL.NumLines
    AppendReportLineS grt, "Wing Type", GetTypeNameACWing(UD.AC.WingType), GetTypeNameACWing(xUD.AC.WingType)
    Select Case UD.AC.WingType
    Case 3  'fixed-wing
      AppendReportLineN grt, "Semispan", UN_LENGTH, UD.AC.SemiSpan, xUD.AC.SemiSpan
      AppendReportLineN grt, "Typical Speed", UN_SPEED, UD.AC.TypSpeed, xUD.AC.TypSpeed
      AppendReportLineN grt, "Biplane Separation", UN_LENGTH, UD.AC.BiplSep, xUD.AC.BiplSep
      AppendReportLineN grt, "Weight", UN_MASS, UD.AC.Weight, xUD.AC.Weight
      AppendReportLineN grt, "Planform Area", UN_AREA, UD.AC.PlanArea, xUD.AC.PlanArea
      AppendReportLineN grt, "Propeller RPM", UN_NONE, UD.AC.PropRPM, xUD.AC.PropRPM
      AppendReportLineN grt, "Propeller Radius", UN_LENGTH, UD.AC.PropRad, xUD.AC.PropRad
      AppendReportLineN grt, "Engine Vert Distance", UN_LENGTH, UD.AC.EngVert, xUD.AC.EngVert
      AppendReportLineN grt, "Engine Fwd Distance", UN_LENGTH, UD.AC.EngFwd, xUD.AC.EngFwd
    Case 4  'helicopter
      AppendReportLineN grt, "Rotor Radius", UN_LENGTH, UD.AC.SemiSpan, xUD.AC.SemiSpan
      AppendReportLineN grt, "Typical Speed", UN_SPEED, UD.AC.TypSpeed, xUD.AC.TypSpeed
      AppendReportLineN grt, "Weight", UN_MASS, UD.AC.Weight, xUD.AC.Weight
      AppendReportLineN grt, "Rotor RPM", UN_NONE, UD.AC.PropRPM, xUD.AC.PropRPM
    End Select
    AppendStr grt, "", True
  End Select

  'Drop Size Distributions
  If UD.Tier > TIER_1 Then
    For iDSD = 0 To MAX_DSD - 1
      If DSDIsUsed(UD, iDSD) Then
        s = ""
        Select Case UD.Tier
        Case TIER_2: s = "--Drop Size Distribution--"
        Case TIER_3: s = "-Drop Size Distribution " + Format(iDSD + 1) + "-"
        End Select
        AppendReportLineS grt, s, Col2Hdr, Col3Hdr
        AppendReportLineS grt, "Name", UD.DSD(iDSD).Name, xUD.DSD(iDSD).Name
        AppendReportLineS grt, "Type", GetTypeNameDSD(UD.DSD(iDSD).Type), GetTypeNameDSD(xUD.DSD(iDSD).Type)
        
        Select Case UD.DSD(iDSD).Type
        Case 1 'DropKick
          AppendReportLineS grt, "Nozzle Type", GetTypeNameDKNoz(UD.DK(iDSD).NozType), GetTypeNameDKNoz(xUD.DK(iDSD).NozType)
          AppendReportLineS grt, "Nozzle Name", UD.DK(iDSD).NameNoz, xUD.DK(iDSD).NameNoz
          AppendReportLineN grt, "Dv0.5 (µm)", UN_NONE, UD.DK(iDSD).VMD, xUD.DK(iDSD).VMD
          AppendReportLineN grt, "Relative Span", UN_NONE, UD.DK(iDSD).RelSpan, xUD.DK(iDSD).RelSpan
          AppendReportLineN grt, "Effective Diameter (cm)", UN_NONE, UD.DK(iDSD).EffDiam, xUD.DK(iDSD).EffDiam
          AppendReportLineS grt, "Material Type", GetTypeNameDKMat(UD.DK(iDSD).MatType), GetTypeNameDKMat(xUD.DK(iDSD).MatType)
          AppendReportLineS grt, "Material Name", UD.DK(iDSD).NameMat, xUD.DK(iDSD).NameMat
          AppendStr grt, Format$("Dynamic Surface", c1fmt), True
          AppendReportLineN grt, "  Tension (dynes/cm)", UN_NONE, UD.DK(iDSD).DynSurfTens, xUD.DK(iDSD).DynSurfTens
          AppendReportLineN grt, "Shear Viscosity (cp)", UN_NONE, UD.DK(iDSD).ShearVisc, xUD.DK(iDSD).ShearVisc
          AppendStr grt, Format$("Elongational", c1fmt), True
          AppendReportLineN grt, "  Viscosity (cp)", UN_NONE, UD.DK(iDSD).ElongVisc, xUD.DK(iDSD).ElongVisc
          AppendReportLineN grt, "Density (gm/cm³)", UN_NONE, UD.DK(iDSD).Density, xUD.DK(iDSD).Density
          AppendReportLineN grt, "Air Speed", UN_SPEED, UD.DK(iDSD).Speed, xUD.DK(iDSD).Speed
          AppendReportLineN grt, "Nozzle Orientation (deg)", UN_NONE, UD.DK(iDSD).NozAngle, xUD.DK(iDSD).NozAngle
          AppendReportLineN grt, "Pressure", UN_PRESSURE, UD.DK(iDSD).Pressure, xUD.DK(iDSD).Pressure
          AppendReportLineN grt, "Flow Rate", UN_FLOWRATE, UD.DK(iDSD).flow, xUD.DK(iDSD).flow
          AppendReportLineN grt, "Spray Type", UN_NONE, UD.DK(iDSD).SprayType, xUD.DK(iDSD).SprayType
        Case 5 'ARS (DropKirk)
          AppendReportLineS grt, "Nozzle Type", GetTypeNameDKNoz(UD.BK(iDSD).NozType), GetTypeNameDKNoz(xUD.BK(iDSD).NozType)
          AppendReportLineS grt, "Nozzle Name", UD.BK(iDSD).NameNoz, xUD.BK(iDSD).NameNoz
          AppendReportLineN grt, "Orifice (in or #)", UN_NONE, UD.BK(iDSD).Orifice, xUD.BK(iDSD).Orifice
          AppendReportLineN grt, "Air Speed", UN_SPEED, UD.BK(iDSD).Speed, xUD.BK(iDSD).Speed
          AppendReportLineN grt, "Nozzle Angle (deg)", UN_NONE, UD.BK(iDSD).NozAngle, xUD.BK(iDSD).NozAngle
          AppendReportLineN grt, "Pressure", UN_PRESSURE, UD.BK(iDSD).Pressure, xUD.BK(iDSD).Pressure
          AppendReportLineN grt, "Spray Type", UN_NONE, UD.BK(iDSD).SprayType, xUD.BK(iDSD).SprayType
        End Select
        AppendReportLineCH grt, "Drop Categories", UN_NONE, "Diam (um)", UN_NONE, "Frac"
        For i = 0 To UD.DSD(iDSD).NumDrop - 1
          AppendReportLineC grt, i + 1, UN_NONE, "0.00", UD.DSD(iDSD).Diam(i), xUD.DSD(iDSD).Diam(i), _
                                        UN_NONE, "0.0000", UD.DSD(iDSD).MassFrac(i), xUD.DSD(iDSD).MassFrac(i)
        Next
        AppendStr grt, "", True
      End If
    Next
  End If

  'Nozzles
  If UD.Tier > TIER_2 Then
    AppendReportLineS grt, "--Nozzle Distribution--", Col2Hdr, Col3Hdr
'    AppendReportLineS grt, "Name", UD.NZ.Name, xUD.NZ.Name
'    AppendReportLineS grt, "Type", GetTypeNameNZ(UD.NZ.Type), GetTypeNameNZ(xUD.NZ.Type)
'    AppendReportLineN grt, "Horiz Distance Limit (%)", UN_NONE, UD.NZ.PosHorizLimit, xUD.NZ.PosHorizLimit
    AppendReportLineN grt, "Boom Length (%)", UN_NONE, UD.NZ.BoomWidth, xUD.NZ.BoomWidth
    AppendReportLineCH grt, "Nozzle DSD & Locations", _
                            UN_NONE, "DSD", UN_LENGTH, "H", UN_LENGTH, "V", UN_LENGTH, "F"
    For i = 0 To UD.NZ.NumNoz - 1
      AppendReportLineC grt, i + 1, UN_NONE, , UD.NZ.NozType(i) + 1, xUD.NZ.NozType(i) + 1, _
                                    UN_LENGTH, , UD.NZ.PosHoriz(i), xUD.NZ.PosHoriz(i), _
                                    UN_LENGTH, , UD.NZ.PosVert(i), xUD.NZ.PosVert(i), _
                                    UN_LENGTH, , UD.NZ.PosFwd(i), xUD.NZ.PosFwd(i)
    Next
    AppendStr grt, "", True
  End If
  
  'Swath
  If UD.Tier > TIER_1 Then
    AppendReportLineS grt, "--Swath--", Col2Hdr, Col3Hdr
    s1 = "": s2 = ""
    Select Case UD.CTL.SwathWidthType
    Case 0:    s1 = AGFormat$(UnitsDisplay(UD.CTL.SwathWidth, UN_LENGTH)) + " " & UnitsName(UN_LENGTH)
    Case 1, 2: s1 = AGFormat$(UD.CTL.SwathWidth) + " x Wingspan"
    End Select
    Select Case xUD.CTL.SwathWidthType
    Case 0:    s2 = AGFormat$(UnitsDisplay(xUD.CTL.SwathWidth, UN_LENGTH)) + " " & UnitsName(UN_LENGTH)
    Case 1, 2: s2 = AGFormat$(xUD.CTL.SwathWidth) + " x Wingspan"
    End Select
    AppendReportLineS grt, "Swath Width", s1, s2
    
    s1 = "": s2 = ""
    Select Case UD.CTL.SwathDispType
    Case 0: s1 = AGFormat$(UD.CTL.SwathDisp) + " x Swath Width"
    Case 1: s1 = AGFormat$(UD.CTL.SwathDisp) + " x Application Rate"
    Case 2: s1 = AGFormat$(UnitsDisplay(UD.CTL.SwathDisp, UN_LENGTH)) + " " & UnitsName(UN_LENGTH)
    Case 3: s1 = "Aircraft Centerline"
    End Select
    Select Case xUD.CTL.SwathDispType
    Case 0: s2 = AGFormat$(xUD.CTL.SwathDisp) + " x Swath Width"
    Case 1: s2 = AGFormat$(xUD.CTL.SwathDisp) + " x Application Rate"
    Case 2: s2 = AGFormat$(UnitsDisplay(xUD.CTL.SwathDisp, UN_LENGTH)) + " " & UnitsName(UN_LENGTH)
    Case 3: s2 = "Aircraft Centerline"
    End Select
    AppendReportLineS grt, "Swath Displacement", s1, s2
    If UD.Tier > TIER_2 Then
      AppendReportLineS grt, "Half Boom", Format(UD.CTL.HalfBoom, "YES/NO"), Format(xUD.CTL.HalfBoom, "YES/NO")
    End If
    AppendStr grt, "", True
  End If
  
  'Spray Material
  Select Case UD.Tier
  Case TIER_2
    AppendReportLineS grt, "--Spray Material--", Col2Hdr, Col3Hdr
    If UD.Smokey = 0 Then
      AppendReportLineN grt, "Nonvolatile Rate", UN_RATEMASS, UD.SM.NVfrac * UD.SM.FlowRate * UD.SM.NonVGrav, _
                                                               xUD.SM.NVfrac * xUD.SM.FlowRate * xUD.SM.NonVGrav
      AppendReportLineN grt, "Active Rate", UN_RATEMASS, UD.SM.ACfrac * UD.SM.FlowRate * UD.SM.NonVGrav, _
                                                          xUD.SM.ACfrac * xUD.SM.FlowRate * xUD.SM.NonVGrav
    Else
      AppendReportLineN grt, "Nonvolatile Fraction", UN_NONE, UD.SM.NVfrac, xUD.SM.NVfrac
      AppendReportLineN grt, "Active Fraction", UN_NONE, UD.SM.ACfrac, xUD.SM.ACfrac
    End If
    AppendStr grt, Format$("Spray Volume", c1fmt), True
    AppendReportLineN grt, "  Rate", UN_RATEVOL, UD.SM.FlowRate, xUD.SM.FlowRate
    AppendReportLineS grt, "Carrier Type", GetBasicNameSM(UD.SM.BasicType), GetBasicNameSM(xUD.SM.BasicType)
    AppendStr grt, "", True
  Case TIER_3
    AppendReportLineS grt, "--Spray Material--", Col2Hdr, Col3Hdr
    AppendReportLineS grt, "Name", UD.SM.Name, xUD.SM.Name
    AppendReportLineS grt, "Type", GetTypeNameSM(UD.SM.Type), GetTypeNameSM(xUD.SM.Type)
    If UD.Smokey = 0 Then
      AppendReportLineN grt, "Nonvolatile Rate", UN_RATEMASS, UD.SM.NVfrac * UD.SM.FlowRate * UD.SM.NonVGrav, _
                                                               xUD.SM.NVfrac * xUD.SM.FlowRate * xUD.SM.NonVGrav
      AppendReportLineN grt, "Active Rate", UN_RATEMASS, UD.SM.ACfrac * UD.SM.FlowRate * UD.SM.NonVGrav, _
                                                          xUD.SM.ACfrac * xUD.SM.FlowRate * xUD.SM.NonVGrav
    Else
      AppendReportLineN grt, "Nonvolatile Fraction", UN_NONE, UD.SM.NVfrac, xUD.SM.NVfrac
      AppendReportLineN grt, "Active Fraction", UN_NONE, UD.SM.ACfrac, xUD.SM.ACfrac
    End If
    AppendStr grt, Format$("Spray Volume", c1fmt), True
    AppendReportLineN grt, "  Rate", UN_RATEVOL, UD.SM.FlowRate, xUD.SM.FlowRate
    AppendReportLineN grt, "Specific Gravity", UN_NONE, UD.SM.SpecGrav, xUD.SM.SpecGrav
    AppendStr grt, Format$("Evaporation", c1fmt), True
    AppendReportLineN grt, "  Rate (µm²/deg C/sec)", UN_NONE, UD.SM.EvapRate, xUD.SM.EvapRate
    AppendStr grt, "", True
  End Select
  
  'Meteorology
  If UD.Tier > TIER_1 Then
    AppendReportLineS grt, "--Meteorology--", Col2Hdr, Col3Hdr
    AppendReportLineN grt, "Wind Speed", UN_SPEED, UD.MET.WS, xUD.MET.WS
    If UD.Tier > TIER_2 Then
      AppendReportLineN grt, "Wind Direction (deg)", UN_NONE, UD.MET.WD, xUD.MET.WD
    End If
    AppendReportLineN grt, "Temperature", UN_TEMP, UD.MET.temp, xUD.MET.temp
    AppendReportLineN grt, "Relative Humidity (%)", UN_NONE, UD.MET.Humidity, xUD.MET.Humidity
    AppendStr grt, "", True
  End If
  
  'Transport
  If UD.Tier > TIER_1 Then
    AppendReportLineS grt, "--Transport--", Col2Hdr, Col3Hdr
    AppendReportLineN grt, "Flux Plane", UN_LENGTH, UD.CTL.FluxPlane, xUD.CTL.FluxPlane
    AppendStr grt, "", True
  End If
  
  'Canopy
  If UD.Tier > TIER_1 And UD.Smokey = AUD_FS Then
    AppendReportLineS grt, "--Canopy--", Col2Hdr, Col3Hdr
    Select Case UD.Tier
    Case TIER_2
      AppendReportLineN grt, "Height", UN_LENGTH, UD.CAN.Height, xUD.CAN.Height
      AppendReportLineN grt, "Canopy Roughness", UN_LENGTH, UD.CAN.NDRuff * UD.CAN.Height, xUD.CAN.NDRuff * xUD.CAN.Height
      AppendReportLineN grt, "Canopy Displacement", UN_LENGTH, UD.CAN.NDDisp * UD.CAN.Height, xUD.CAN.NDDisp * xUD.CAN.Height
    Case TIER_3
      If UD.CAN.Type <> 0 Then
        AppendReportLineS grt, "Name", UD.CAN.Name, xUD.CAN.Name
      End If
      AppendReportLineS grt, "Type", GetTypeNameCN(UD.CAN.Type), GetTypeNameCN(xUD.CAN.Type)
      Select Case UD.CAN.Type
      Case 0 'none
        'no canopy, nothing else to print!
      Case 1 'Story
        AppendReportLineN grt, "Element Size", UN_LENGTH, UD.CAN.EleSiz, xUD.CAN.EleSiz
        AppendReportLineN grt, "Temperature", UN_TEMP, UD.CAN.temp, xUD.CAN.temp
        AppendReportLineN grt, "Humidity", UN_PERCENT, UD.CAN.Humidity, xUD.CAN.Humidity
        AppendStr grt, Format$("Stand", c1fmt), True
        AppendReportLineN grt, " Density", UN_STANDDENSITY, UD.CAN.StanDen, xUD.CAN.StanDen
        AppendReportLineN grt, "Canopy Roughness", UN_LENGTH, UD.CAN.NDRuff * UD.CAN.Height, xUD.CAN.NDRuff * xUD.CAN.Height
        AppendReportLineN grt, "Canopy Displacement", UN_LENGTH, UD.CAN.NDDisp * UD.CAN.Height, xUD.CAN.NDDisp * xUD.CAN.Height
        AppendReportLineCH grt, "Tree Envelope", UN_LENGTH, "Hgt", UN_LENGTH, "Dia", UN_NONE, "PoP"
        For i = 0 To UD.CAN.NumEnv - 1
          AppendReportLineC grt, i + 1, UN_LENGTH, , UD.CAN.EnvHgt(i), xUD.CAN.EnvHgt(i), _
                                        UN_LENGTH, , UD.CAN.EnvDiam(i), xUD.CAN.EnvDiam(i), _
                                        UN_NONE, , UD.CAN.EnvPop(i), xUD.CAN.EnvPop(i)
        Next
      Case 2 'Optical
        AppendReportLineN grt, "Element Size", UN_LENGTH, UD.CAN.EleSiz, xUD.CAN.EleSiz
        AppendReportLineN grt, "Temperature", UN_TEMP, UD.CAN.temp, xUD.CAN.temp
        AppendReportLineN grt, "Humidity", UN_PERCENT, UD.CAN.Humidity, xUD.CAN.Humidity
        AppendReportLineN grt, "Canopy Roughness", UN_LENGTH, UD.CAN.NDRuff * UD.CAN.Height, xUD.CAN.NDRuff * xUD.CAN.Height
        AppendReportLineN grt, "Canopy  Displacement", UN_LENGTH, UD.CAN.NDDisp * UD.CAN.Height, xUD.CAN.NDDisp * xUD.CAN.Height
        AppendReportLineS grt, "Optical Canopy Type", GetTypeNameOP(UD.CAN.optType), GetTypeNameOP(xUD.CAN.optType)
        Select Case UD.CAN.optType
        Case 1 'User-Defined
          AppendReportLineCH grt, "LAI Envelope", UN_LENGTH, "Hgt", UN_NONE, "LAI"
          For i = 0 To UD.CAN.NumLAI - 1
            AppendReportLineC grt, i + 1, UN_LENGTH, , UD.CAN.LAIHgt(i), xUD.CAN.LAIHgt(i), _
                                          UN_NONE, , UD.CAN.LAICum(i), xUD.CAN.LAICum(i)
          Next
        Case 2 'Library
          AppendReportLineN grt, "Height", UN_LENGTH, UD.CAN.LibHgt, xUD.CAN.LibHgt
          AppendReportLineN grt, "LAI", UN_NONE, UD.CAN.LibLAI, xUD.CAN.LibLAI
        End Select
      Case 3 'Basic
        AppendReportLineN grt, "Height", UN_LENGTH, UD.CAN.Height, xUD.CAN.Height
        AppendReportLineN grt, "Canopy Roughness", UN_LENGTH, UD.CAN.NDRuff * UD.CAN.Height, xUD.CAN.NDRuff * xUD.CAN.Height
        AppendReportLineN grt, "Canopy Displacement", UN_LENGTH, UD.CAN.NDDisp * UD.CAN.Height, xUD.CAN.NDDisp * xUD.CAN.Height
      End Select
    End Select
    AppendStr grt, "", True
  End If
  
  'Terrain
  If UD.Tier > TIER_2 Then
    AppendReportLineS grt, "--Terrain--", Col2Hdr, Col3Hdr
    If UD.CAN.Type = 0 Then 'print only if no canopy
      AppendReportLineN grt, "Surface Roughness", UN_LENGTH, UD.MET.SurfRough, xUD.MET.SurfRough
    End If
    If UD.Smokey = AUD_FS Then
      AppendReportLineN grt, "Upslope Angle (deg)", UN_NONE, UD.TRN.Upslope, xUD.TRN.Upslope
      AppendReportLineN grt, "Sideslope Angle (deg)", UN_NONE, UD.TRN.Sideslope, xUD.TRN.Sideslope
    End If
    AppendStr grt, "", True
  End If

  'Advanced
  If UD.Tier > TIER_2 Then
    AppendReportLineS grt, "--Advanced--", Col2Hdr, Col3Hdr
    AppendReportLineN grt, "Wind Speed Height", UN_LENGTH, UD.MET.WindHeight, xUD.MET.WindHeight
    AppendReportLineN grt, "Max Compute Time (sec)", UN_NONE, UD.CTL.MaxComputeTime, xUD.CTL.MaxComputeTime
    AppendReportLineN grt, "Max Downwind Dist", UN_LENGTH, UD.CTL.MaxDownwindDist, xUD.CTL.MaxDownwindDist
    AppendReportLineN grt, "Vortex Decay Rate", UN_SPEED, UD.MET.VortexDecay, xUD.MET.VortexDecay
    AppendReportLineN grt, "Aircraft Drag Coeff", UN_NONE, UD.AC.DragCoeff, xUD.AC.DragCoeff
    AppendReportLineN grt, "Propeller Efficiency", UN_NONE, UD.AC.PropEff, xUD.AC.PropEff
    AppendReportLineN grt, "Ambient Pressure", UN_AIRPRESSURE, UD.MET.Pressure, xUD.MET.Pressure
    AppendStr grt, "", True
  End If

  GenReportText = grt
End Function

Public Sub AppendReportLineN(s As String, Name As String, _
                            Units As Integer, CurVal, DefVal)
'Standard report line: single numbers (floating or integer)
  
  'Column 1 - the name
  If Units = UN_NONE Then
    AppendStr s, Format$(Name, c1fmt), False
  Else
    AppendStr s, Format$(Name & " (" & UnitsName(Units) & ")", c1fmt), False
  End If
  
  'Column 2 - the current value
  AppendStr s, Format$(AGFormat$(UnitsDisplay(CurVal, Units)), c2fmt), False
  
  'Column 3 - the default value
  If (CurVal <> DefVal) Then
    AppendStr s, Format$(AGFormat$(UnitsDisplay(DefVal, Units)), c3fmt), False
  End If
  
  'end with a CR
  AppendStr s, "", True
End Sub

Public Sub AppendReportLineS(s As String, _
                             Name As String, CurVal, DefVal)
'Standard report line: strings
  
  'Column 1 - the name
  AppendStr s, Format$(Name, c1fmt), False
  
  'Column 2 - the current value
  AppendStr s, Format$(ClipStr$(CurVal, c2wid), c2fmt), False
  
  'Column 3 - the default value
  If (CurVal <> DefVal) Then
    AppendStr s, Format$(ClipStr$(DefVal, c3wid), c3fmt), False
  End If
  
  'end with a CR
  AppendStr s, "", True
End Sub

Public Sub AppendReportLineC(s As String, Index As Integer, _
                             Optional Units1, Optional Fmt1, Optional CurVal1, Optional DefVal1, _
                             Optional Units2, Optional Fmt2, Optional CurVal2, Optional DefVal2, _
                             Optional Units3, Optional Fmt3, Optional CurVal3, Optional DefVal3, _
                             Optional Units4, Optional Fmt4, Optional CurVal4, Optional DefVal4)
'Standard report line: 1- to 4-column table entries
'If Fmtx is missing, use AGFormat

  Dim s0 As String, s1 As String
  Dim subwid As Integer
  Dim subfmt As String
  Dim ncol As Integer
  Dim ShowDefaults As Boolean
  
  'Compute subcolumn width and define format
  ncol = 0
  If Not IsMissing(CurVal1) Then ncol = 1
  If Not IsMissing(CurVal2) Then ncol = 2
  If Not IsMissing(CurVal3) Then ncol = 3
  If Not IsMissing(CurVal4) Then ncol = 4
  subwid = (c2wid - ncol - 1) / ncol
  subfmt = String$(subwid, "@")
  
  'Column 1 - the Index, right justified
  AppendStr s, Format$(CStr(Index), String$(c1wid, "@")), False
  
  'Column 2 - the current values
  s0 = "": s1 = ""
  If Not IsMissing(CurVal1) Then
    If Not IsMissing(Fmt1) Then
      s1 = Format(UnitsDisplay(CurVal1, Units1), Fmt1)
    Else
      s1 = AGFormat$(UnitsDisplay(CurVal1, Units1))
    End If
    AppendStr s0, Format(s1, subfmt), False
  End If
  If Not IsMissing(CurVal2) Then
    If Not IsMissing(Fmt2) Then
      s1 = Format(UnitsDisplay(CurVal2, Units2), Fmt2)
    Else
      s1 = AGFormat$(UnitsDisplay(CurVal2, Units2))
    End If
    AppendStr s0, " " + Format(s1, subfmt), False
  End If
  If Not IsMissing(CurVal3) Then
    If Not IsMissing(Fmt3) Then
      s1 = Format(UnitsDisplay(CurVal3, Units3), Fmt3)
    Else
      s1 = AGFormat$(UnitsDisplay(CurVal3, Units3))
    End If
    AppendStr s0, " " + Format(s1, subfmt), False
  End If
  If Not IsMissing(CurVal4) Then
    If Not IsMissing(Fmt4) Then
      s1 = Format(UnitsDisplay(CurVal4, Units4), Fmt4)
    Else
      s1 = AGFormat$(UnitsDisplay(CurVal4, Units4))
    End If
    AppendStr s0, " " + Format(s1, subfmt), False
  End If
  AppendStr s, Format$(s0, c2fmt), False
  
  'Column 3 - the default values
  ShowDefaults = False
  If Not IsMissing(CurVal1) Then
    If (CurVal1 <> DefVal1) Then ShowDefaults = True
  End If
  If Not IsMissing(CurVal2) Then
    If (CurVal2 <> DefVal2) Then ShowDefaults = True
  End If
  If Not IsMissing(CurVal3) Then
    If (CurVal3 <> DefVal3) Then ShowDefaults = True
  End If
  If Not IsMissing(CurVal4) Then
    If (CurVal4 <> DefVal4) Then ShowDefaults = True
  End If
  If ShowDefaults Then
    s0 = "": s1 = ""
    If Not IsMissing(DefVal1) Then
      If Not IsMissing(Fmt1) Then
        s1 = Format(UnitsDisplay(DefVal1, Units1), Fmt1)
      Else
        s1 = AGFormat$(UnitsDisplay(DefVal1, Units1))
      End If
      AppendStr s0, Format(s1, subfmt), False
    End If
    If Not IsMissing(DefVal2) Then
      If Not IsMissing(Fmt2) Then
        s1 = Format(UnitsDisplay(DefVal2, Units2), Fmt2)
      Else
        s1 = AGFormat$(UnitsDisplay(DefVal2, Units2))
      End If
      AppendStr s0, " " + Format(s1, subfmt), False
    End If
    If Not IsMissing(DefVal3) Then
      If Not IsMissing(Fmt3) Then
        s1 = Format(UnitsDisplay(DefVal3, Units3), Fmt3)
      Else
        s1 = AGFormat$(UnitsDisplay(DefVal3, Units3))
      End If
      AppendStr s0, " " + Format(s1, subfmt), False
    End If
    If Not IsMissing(DefVal4) Then
      If Not IsMissing(Fmt4) Then
        s1 = Format(UnitsDisplay(DefVal4, Units4), Fmt4)
      Else
        s1 = AGFormat$(UnitsDisplay(DefVal4, Units4))
      End If
      AppendStr s0, " " + Format(s1, subfmt), False
    End If
    AppendStr s, Format$(s0, c3fmt), False
  End If
  
  'end with a CR
  AppendStr s, "", True
End Sub

Public Sub AppendReportLineCH(s As String, Name As String, _
                              Optional Units1, Optional SubCol1, _
                              Optional Units2, Optional SubCol2, _
                              Optional Units3, Optional SubCol3, _
                              Optional Units4, Optional SubCol4)
'Standard report line: 1- to 4-column table headers
  Dim s0 As String, s1 As String, s2 As String, s3 As String, s4 As String
  Dim subwid As Integer
  Dim subfmt As String
  Dim ncol As Integer
  
  'Compute subcolumn width and define format
  ncol = 0
  If Not IsMissing(SubCol1) Then ncol = 1
  If Not IsMissing(SubCol2) Then ncol = 2
  If Not IsMissing(SubCol3) Then ncol = 3
  If Not IsMissing(SubCol4) Then ncol = 4
  subwid = (c2wid - ncol - 1) / ncol
  subfmt = String$(subwid, "@")
  
  'Column 1 - the Name, left justified, + the Index Header, right justified
  AppendStr s, Format$(Name, "!" + String$(c1wid - 1, "@")) + "#", False
  
  'Construct the subcolumns
  If Not IsMissing(SubCol1) Then
    s1 = SubCol1
    If Units1 <> UN_NONE Then AppendStr s1, "(" + UnitsName(Units1) + ")", False
  End If
  If Not IsMissing(SubCol2) Then
    s2 = SubCol2
    If Units2 <> UN_NONE Then AppendStr s2, "(" + UnitsName(Units2) + ")", False
  End If
  If Not IsMissing(SubCol3) Then
    s3 = SubCol3
    If Units3 <> UN_NONE Then AppendStr s3, "(" + UnitsName(Units3) + ")", False
  End If
  If Not IsMissing(SubCol4) Then
    s4 = SubCol4
    If Units4 <> UN_NONE Then AppendStr s4, "(" + UnitsName(Units4) + ")", False
  End If
  
  'Assemble subcolumns
  s0 = ""
  If Not IsMissing(SubCol1) Then AppendStr s0, Format$(ClipStr$(s1, subwid), subfmt), False
  If Not IsMissing(SubCol2) Then AppendStr s0, " " + Format$(ClipStr$(s2, subwid), subfmt), False
  If Not IsMissing(SubCol3) Then AppendStr s0, " " + Format$(ClipStr$(s3, subwid), subfmt), False
  If Not IsMissing(SubCol4) Then AppendStr s0, " " + Format$(ClipStr$(s4, subwid), subfmt), False
  
  'Column 2 - the column headers
  AppendStr s, Format$(s0, c2fmt), False
  
  'Column 3 - the column headers again
  AppendStr s, Format$(s0, c3fmt), False
  
  'end with a CR
  AppendStr s, "", True
End Sub

