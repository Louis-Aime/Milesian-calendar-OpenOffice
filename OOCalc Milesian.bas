REM  *****  BASIC  *****

'Milesian Calendar: Enter and display dates in OOCalc followin Milesian calendar conventions
'Copyright Miletus SARL 2018. www.calendriermilesien.org
'For use as a Basic module.
'Developped following a VBA package
' -> MacOS epoch not taken into account
'Tested under LibreOffice Calc V5.0.
'No warranty.
'May be used for personal or professional purposes.
'If transmitted or integrated, even with changes, present header shall be maintained in full.
'Functions are aimed at extending Date & Time functions, and use similar parameters syntax in English
'Version V1.0 M2018 - enhancements expected
'
Const MStoPresentEra As Long = 986163 'Offset between 1/1m/-800 epoch and Microsoft origin (1899-12-31T00 is 1)
Const InvArgMsg = "Err :502"
Const InvalidNumber = 2147483647 'A very high number to be set for wrong values of entry parameters

'Const DayOffsetMacOS As Long = 1462  'Difference between Windows epoch and MacOS epoch - not used in Libre Office

'#Part 1: internal procedures

Sub Milesian_IntegDiv(ByVal Dividend As Long, ByVal Divisor As Long, Cycle As Long, Phase As Long)
'Quotient and modulo in the same operation. Divisor shall by positive.
'Cycle (i.e. Quotient) is same sign as Dividend. 0 <= Phase (i.e. modulo) < Divisor.
Cycle = 0
Phase = Dividend
If Divisor > 0 Then
 While Phase < 0
   Phase = Phase + Divisor
   Cycle = Cycle - 1
   Wend
 While Phase >= Divisor
    Phase = Phase - Divisor
    Cycle = Cycle + 1
    Wend
Else
	Cycle = InvalidNumber
	Phase = InvalidNumber
End If
End Sub

Sub Milesian_IntegDivCeiling(ByVal Dividend As Long, ByVal Divisor As Long, ByVal ceiling As Integer, Cycle As Long, Phase As Long)
'Quotient and modulo in the same operation. By exception, remainder may be = divisor if quotient = ceiling
'Cycle (i.e. Quotient) is same sign as Dividend. 0 <= Phase (i.e. modulo) <= Divisor.
Cycle = 0
Phase = Dividend
If Divisor > 0 And Dividend >= 0 And Dividend <= ceiling * Divisor + 1 Then
 ceiling = ceiling - 1 'Decrease ceiling by 1 in order to simplify test in the next loop
 While (Phase >= Divisor) And Cycle < ceiling
 Phase = Phase - Divisor
 Cycle = Cycle + 1
 Wend
Else
	Cycle = InvalidNumber
	Phase = InvalidNumber
End If
End Sub

Function PosDiv (ByVal A, D) 'The Integer division with positive remainder
PosDiv = 0
  If D <= 0 Then 
	PosDiv = InvArgMsg
  Else
	While (A < 0)
		A = A + D
		PosDiv = PosDiv - 1
	Wend
	While (A >= D)
		A = A - D
		PosDiv = PosDiv + 1
	Wend
  End If
End Function

'#Part 2: a function used internally, but available to user

Function MILESIAN_IS_LONG_YEAR(ByVal Year As Long) As Boolean
'Is year Year a 366 days year, i.e. a year just before a bissextile year following the Milesian rule.
'Search for any value of year, provided that year are on a continuous degree
'On Error Goto ErrorHandler
	If IsNumeric(Year) Then	'Beware : no test at all is performed !!
		Year = Year + 1
		MILESIAN_IS_LONG_YEAR = ((Year Mod 4) = 0) And ((Year Mod 100) <> 0) Or (((Year Mod 400) = 0) And ((Year + 800) Mod 3200) <> 0)
	Else
		MILESIAN_IS_LONG_YEAR = InvArgMsg
	End If
End Function

'#Part 3: Compute date from milesian parameters

Function MILESIAN_DATE(Year As Long, Month As Integer, DayInMonth As Integer) As Long
'Date number from a Milesian date given as year (positive or negative), month, daynumber in month
'Result is forced to a long number, elsewhise an ambiguous string is created.
  On Error Goto ErrorHandler
  Dim A As Integer 'Intermediate computations as non-long integer
  Dim B As Long   'Bimester number, for intermediate computations
  Dim M1 As Long  'Month rank
  Dim D As Long   'Days expressed in long integer
'Check that Milesian date is OK
  If Month > 0 And Month < 13 And DayInMonth > 0 And DayInMonth < 32 Then 'Basic filter
	M1 = Month - 1 'Count month rank, i.e. 0..11
	Milesian_IntegDiv M1, 2, B, M1 'B = full bimesters, M1 = 1 if a full month added, else 0
	If DayInMonth < 31 Or (M1 = 1 And (B < 5 Or MILESIAN_IS_LONG_YEAR(Year))) Then
	  Year = Year + 800    'Set Epoch to the year -800
	  A = PosDiv (Year, 4) - PosDiv(Year, 100) + PosDiv(Year, 400) - PosDiv(Year, 3200) 'Sum non-long terms: leap days
	  D = Year            'Force long-integer conversion
	  D = D * 365      'Begin computation of days in long-integer;
	  D = D - MStoPresentEra - 1 + B * 61 + M1 * 30 + A + DayInMonth 'Computations in long-integer first
' Date1904 specific parts cancelled
' If ActiveWorkbook.Date1904 Then
' D = D - DayOffsetMacOS
' If (D < 0) Then Error 1  'With Mac Calendar, Day is not authorised to be < 0
' End If
	  MILESIAN_DATE = D
	Else	' Cases where date elements do not build a correct milesian date
	  MILESIAN_DATE = InvalidNumber	'A very big number
	End If
  Else
	MILESIAN_DATE = InvalidNumber
  End If
Exit Function
ErrorHandler:
  MILESIAN_DATE = InvalidNumber
End Function

Function MILESIAN_YEAR_BASE(Year As Long) As Long 'The Year base or Doomsday of a year i.e. the date just before the 1 1m of the year
On Error Goto ErrorHandler
Dim A As Integer, D As Long   'Force long integer
' If Year < 1 Then Error 1 'No specific control, as date before year 1 may be handled.
Year = Year + 800    'Set Epoch to the year -800
D = Year            'Force long-interger conversion
D = D * 365      'Begin computation of days in long-integer;
A = PosDiv (Year, 4) - PosDiv(Year, 100) + PosDiv(Year, 400) - PosDiv(Year, 3200)
D = D - MStoPresentEra + A - 1           'Computations in long-integer first
' Date1904 specific parts cancelled
'If ActiveWorkbook.Date1904 Then
'D = D - DayOffsetMacOS
'If (D < 0) Then Error 1   'With Mac Calendar, Day is not authorised to be < 0
'End If
MILESIAN_YEAR_BASE = D
Exit Function
ErrorHandler: 
MILESIAN_YEAR_BASE = InvalidNumber
End Function

'#Part 4: Extract Milesian elements from Date number

Sub Milesian_DateElement(ByVal EXD As Date, Y As Long, M As Integer, Q As Integer)
' From EXD, a Day Number under Excel, compute the milesian date Q / M / Y (day in month, month, year)
' Y is year in common era. From an Excel date, Y is >= 100
' M is milesian month number, 1 to 12
' Q is number of day in month, 1 to 31
' This is an internal subroutine. Corresponding functions come after.
' Note: Excel under MS and with Date1904 set passes EXD as the MS number, not as the Date1904, except when EXD < 1 !!!
Dim Cycle As Long, Day As Long      'Cycle is used serveral times with a different meaning each time
Day = Int(EXD)                      'Initiate Day as integer part of Excel date, suppress control on EXD
' Date1904 specific parts cancelled
'If ActiveWorkbook.Date1904 And Day = 0 Then Day = Day + DayOffsetMacOS 'Strange behavior under MS, to be checked under Mac OS
Day = Day + MStoPresentEra
Milesian_IntegDiv Day, 1168775, Cycle, Day   'Day is day rank in Milesian era (starting from 1/1m/-800), Cycle is era (0 begins 1/1/-800)
Y = -800 + Cycle * 3200
Milesian_IntegDiv Day, 146097, Cycle, Day    'Day is day rank in 400 years period, Cycle is quadrisaeculum
Y = Y + Cycle * 400
Milesian_IntegDivCeiling Day, 36524, 4, Cycle, Day   'Day is day rank in century, Cycle is rank of century
Y = Y + Cycle * 100
Milesian_IntegDiv Day, 1461, Cycle, Day              'Day rank in quadriannum
Y = Y + Cycle * 4
Milesian_IntegDivCeiling Day, 365, 4, Cycle, Day     'Day rank in year
Y = Y + Cycle
Milesian_IntegDiv Day, 61, Cycle, Day             'Day rank in bimester
M = 2 * Cycle
Milesian_IntegDivCeiling Day, 30, 2, Cycle, Day  'Day: day rank within month, Cycle = month rank in bimester
M = M + Cycle + 1                       'M: month number, 1 to 12
Q = Day + 1                             'Q: day number within month, 1 to 31
End Sub

Function MILESIAN_YEAR(GDate)  'The milesian year (common era) for a given date (a series number or a string)
On Error Goto ErrorHandler
Dim Y As Long, M As Integer, Q As Integer, NumDate as Date
NumDate = GDate 'Force conversion or raise error
Milesian_DateElement NumDate, Y, M, Q   'Compute the 3 figures of the milesian date
MILESIAN_YEAR = Y
Exit Function
ErrorHandler: 
MILESIAN_YEAR = InvArgMsg
End Function

Function MILESIAN_MONTH(GDate)   'The milesian month number (1-12) for a given Excel date
On Error Goto ErrorHandler
Dim Y As Long, M As Integer, Q As Integer, NumDate as Date
NumDate = GDate 'Force conversion or raise error
Milesian_DateElement NumDate, Y, M, Q   'Compute the 3 figures of the milesian date
MILESIAN_MONTH = M
Exit Function
ErrorHandler:  
MILESIAN_MONTH = InvArgMsg
End Function

Function MILESIAN_DAY(GDate)  'The day number in the milesian month for a given Excel date
On Error Goto ErrorHandler
Dim Y As Long, M As Integer, Q As Integer, NumDate as Date
NumDate = GDate 'Force conversion or raise error
Milesian_DateElement NumDate, Y, M, Q   'Compute the 3 figures of the milesian date
MILESIAN_DAY = Q
Exit Function
ErrorHandler:  
MILESIAN_DAY = InvArgMsg
End Function

Function MILESIAN_DISPLAY(GDate) As String
'Milesian date as a string, from an expression that we hope to be a date
On Error Goto ErrorHandler
Dim Y As Long, M As Integer, Q As Integer, NumDate as Date
NumDate = GDate 'Force conversion or raise error
Milesian_DateElement NumDate, Y, M, Q   'Compute the 3 figures of the milesian date
MILESIAN_DISPLAY = Q & " " & M & "m " & Y
Exit Function
ErrorHandler: 
MILESIAN_DISPLAY = InvArgMsg
End Function

'#Part 5: Computations on milesian months

Function MILESIAN_MONTH_SHIFT(GDate, MonthShift As Long) As Long 'Same date several (milesian) months later of earlier
On Error Goto ErrorHandler 'Error comes from wrong parameter
Dim Y As Long, M As Integer, Q As Integer, NumDate as Date
Dim M1 As Long, Cycle As Long, Phase As Long
NumDate = GDate 'Force conversion or raise error
'Compute begin milesian date
Milesian_DateElement NumDate, Y, M, Q
'Compute month rank from 1m of year 0
M1 = Y                     ' Force computation of month in Long
M1 = (M1 * 12) + MonthShift + M - 1 'In this order, Long shall be before simple Integer
'Compute year and month rank
Milesian_IntegDiv M1, 12, Cycle, Phase
Y = Cycle
M = Phase + 1
'If Q was 31, set to end of month, else use same day number
If (Q = 31) And (((M Mod 2) = 1) Or ((M = 12) And Not MILESIAN_IS_LONG_YEAR(Y))) Then Q = 30
MILESIAN_MONTH_SHIFT = MILESIAN_DATE(Y, M, Q)
Exit function
ErrorHandler:
MILESIAN_MONTH_SHIFT = InvalidNumber
End Function

Function MILESIAN_MONTH_END(GDate, MonthShift As Long) As Long 'End of month several (milesian) months later of earlier
On Error Goto ErrorHandler 'Error comes from wrong parameter
Dim Y As Long, M As Integer, Q As Integer, NumDate As Date
Dim M1 As Long, Cycle As Long, Phase As Long
NumDate = GDate 'Force conversion or raise error
'Compute begin milesian date
Milesian_DateElement NumDate, Y, M, Q
'Compute month rank from 1m of year 0
M1 = Y                     ' Force computation of month in Long
M1 = (M1 * 12) + MonthShift + M - 1 'In this order, Long shall be before simple Integer
'Compute year and month rank
Milesian_IntegDiv M1, 12, Cycle, Phase
Y = Cycle
M = Phase + 1
'If Q was 31, set to end of month, else use same day number
If (((M Mod 2) = 1) Or ((M = 12) And Not MILESIAN_IS_LONG_YEAR(Y))) Then
Q = 30
Else: Q = 31
End If
MILESIAN_MONTH_END = MILESIAN_DATE(Y, M, Q)
Exit function
ErrorHandler:
MILESIAN_MONTH_END = InvalidNumber
End Function
