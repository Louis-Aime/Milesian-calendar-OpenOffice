REM  *****  BASIC  *****

Attribute VB_Name = "MeanMoonPhase"
'MeanMoonPhase: Find a moon phase near a date.
'Copyright Miletus SARL 2017-2018. www.calendriermilesien.org
'Extended after the same module in VBA
' -> MacOS epoch not taken into account (no way to reach the parameter)
'Tested under LibreOffice Calc V5.0 and 6.0
'No warranty.
'If transmitted, even with changes, present header shall be maintained in full.
'This package uses no other element.
'Note: the moon computed here is a mean moon, with no Terrestrial Time correction.
'The error on mean moon is less than one day for three milliniums before and after year 2000.
'Functions:
'   MOON_PHASE_LAST : last time the mean moon reached a given phase
'   MOON_PHASE_NEXT : next time where the mean moon reaches a given phase
'Parameters
'   FromDate: a date and time UTC expressed in Excel
'   MoonPhase: 0 or omitted: new moon; 1: first quarter; 2: full moon; 3: last quarter; <0 or >3: error.
'Version M2018-09-01
'Version M2018-12-25 : error handling 

Const MeanSynodicMoon As Double = 29.53058883 'Mean duration of a synodic month in decimal days = 29d 12h 44mn 2s 7/8s
Const MeanNewMoon2000 As Double = 36531.59773 'G2000-01-06T14-20-44 TT, conventional date of a mean new moon
Const InvArgMsg = "Err (Milesian) "	'Error message displayed in place of function result, for non-typed functions

Function MOON_PHASE_LAST(FromDate, Optional MoonPhase As Integer)
Attribute MOON_PHASE_LAST.VB_Description = "Date of last mean moon phase. Phase 0 or omitted: New Moon, 1: First Q, 2: Full Moon, 3: Last Q."
'Date of last mean moon phase before FromDate (in UTC).
'MoonPhase: 0 or omitted, new moon; 1: first quarter; 2: full moon; 3: last quarter; else: error.
On Error Goto ErrorHandler
Dim Phase As Double, Target
If IsMissing(MoonPhase) Then MoonPhase = 0
If MoonPhase < 0 Or MoonPhase > 3 Then Goto ErrorHandler
'If Not IsDate(FromDate) Then Goto ErrorHandler
'Target = 		'Force to extract Date element of FromDate
Phase = FromDate    	'Force conversion to Double in order to avoid Date type controls
Phase = Phase - MeanNewMoon2000 - (MoonPhase / 4) * MeanSynodicMoon
 While Phase < 0
    Phase = Phase + MeanSynodicMoon
 Wend
 While Phase >= MeanSynodicMoon
    Phase = Phase - MeanSynodicMoon
 Wend
MOON_PHASE_LAST = FromDate - Phase 'DateAdd("s",-Phase*86400, Target)
Exit Function
ErrorHandler: 
 MOON_PHASE_LAST = InvArgMsg
End Function

Function MOON_PHASE_NEXT(FromDate, Optional MoonPhase As Integer)
Attribute MOON_PHASE_NEXT.VB_Description = "Date of next mean moon phase. Phase 0 or omitted: New Moon, 1: First Q, 2: Full Moon, 3: Last Q."
'Date of next mean moon phase after FromDate (in UTC).
'MoonPhase: 0 or omitted, new moon; 1: first quarter; 2: full moon; 3: last quarter; else: error.
On Error Goto ErrorHandler
Dim Phase As Double, Target
If IsMissing(MoonPhase) Then MoonPhase = 0
If MoonPhase < 0 Or MoonPhase > 3 Then Goto ErrorHandler
'If Not IsDate(FromDate) Then Goto ErrorHandler
'Target = DateAdd ("d", 0, FromDate)		'Force to extract Date element of FromDate
Phase = FromDate       	'Force conversion to Double in order to avoid Date type controls
Phase = MeanNewMoon2000 - Phase + (MoonPhase / 4) * MeanSynodicMoon
 While Phase < 0
    Phase = Phase + MeanSynodicMoon
 Wend
 While Phase >= MeanSynodicMoon
    Phase = Phase - MeanSynodicMoon
 Wend
MOON_PHASE_NEXT = FromDate - (-Phase) 'Because a simple "+" may generate a string concatenation.
Exit Function
ErrorHandler: 
 MOON_PHASE_NEXT = InvArgMsg
End Function
