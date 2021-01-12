REM  *****  BASIC  *****

Attribute VB_Name = "MilesianCalendar"
'Milesian Calendar: Enter and display dates in Open Office Calc following Milesian calendar conventions
'Copyright Miletus SARL 2018. www.calendriermilesien.org
'For use as a Basic module.
'Extended after the same module in VBA
' -> MacOS epoch not taken into account (no way to reach the parameter)
'Tested under LibreOffice Calc V5.0 and 6.0
'No warranty.
'May be used for personal or professional purposes.
'If transmitted or integrated, even with changes, present header shall be maintained in full.
'Functions are aimed at extending Date & Time functions, and use similar parameters syntax in English
'Version M2021-01-22: 
' MILESIAN_YEAR_BASE yields a day at 0:00, not 7:30
' GREGORIAN_EPACT and MILESIAN_EPACT added
'M2019-06-30: set MILESIAN_YEAR_BASE time to 7:30 UTC for computation of Epact
'M2019-01-15: solar intercalation rule is same as Gregorian
' changed type of Y to Long in Milesian_Date_Element
' modified error handling and error message, mostly with untyped function
' defined highest a lowest dates handled with Calc, for MILESIAN_DATE; milesian figures and display without limitation
'M2018-08-30: modified MILESIAN_DISPLAY, year always with at least 3-digits
Const MStoPresentEra As Long = 693969 'Offset between 1/1m/000 epoch and Microsoft origin (1899-12-30T00 is 0)
Const MStoJulianMinus05 As Long = 2415018 'Offset between julian day epoch and Microsoft origin, minus 0.5
Const HighDate = 11274306 'Highest date that is properly handled in Calc, high limit for MILESIAN_DATE
Const LowDate = -12661859 'Lowest date that is properly handled in Calc, low limit for MILESIAN_DATE
Const HighYear = 32768 'Highest Milesian year handled
Const LowYear = -32768	'Lowest Milesian year handled
Const InvArgMsg = "Err (Milesian) "	'Error message displayed in place of function result, for non-typed functions

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
	'Else part deleted. Should be error raising.
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
		' Else part deleted. Should be error raising.
	End If
End Sub

Private Function PosDiv (ByVal A, D) 'The Integer division with positive remainder
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

Private Function PosMod(ByVal A, D)  'The always positive modulo, even if A<0
    If D <= 0 Then
        PosMod = InvArgMsg
    Else
        While (A < 0)
            A = A + D
        Wend
        While (A >= D)
            A = A - D
        Wend
    PosMod = A
    End If
End Function

'#Part 2: a function used internally, but available to user

Function MILESIAN_IS_LONG_YEAR(ByVal Year) As Boolean
'Is year Year a 366 days year, i.e. a year just before a bissextile year following the Milesian rule.
'Search for any value of year, provided that year are on a continuous degree
On Error Goto ErrorHandler
  If Year <> Int(Year) Or Year < LowYear Or Year > HighYear Then Goto ErrorHandler 'Check that we have an integer numeric value
  Year = Year + 1
  MILESIAN_IS_LONG_YEAR = PosMod (Year,4) = 0 And (PosMod (Year,100) <> 0 Or PosMod(Year, 400) = 0)
  Exit Function
ErrorHandler:
  'MILESIAN_IS_LONG_YEAR = InvArgMsg
  MILESIAN_IS_LONG_YEAR = False
End Function

'#Part 3: Compute date from milesian parameters

Function MILESIAN_DATE(Year, Month, DayInMonth) 'Date set as a long integer, without time element
'Date number from a Milesian date given as year (positive or negative), month, daynumber in month
'Result is forced to a Date objet
'Date must be in the domain, else an ambiguous string is created.
  On Error Goto ErrorHandler
  Dim Result As Date 'Intermediate result forced as Date
  Dim A As Integer 'Intermediate computations as non-long integer
  Dim B As Long   'Bimester number, for intermediate computations
  Dim M1 As Long  'Month rank
  Dim D As Long   'Days expressed in long integer
'Check that Milesian date is OK
  If Year <> Int(Year) Or Month <> Int(Month) Or DayInMonth <> Int (DayInMonth) Then Goto ErrorHandler
  If Year >= LowYear And Year <= HighYear And Month > 0 And Month < 13 And DayInMonth > 0 And DayInMonth < 32 Then 'Basic filter
	M1 = Month - 1 'Count month rank, i.e. 0..11
	Milesian_IntegDiv M1, 2, B, M1 'B = full bimesters, M1 = 1 if a full month added, else 0
	If DayInMonth < 31 Or (M1 = 1 And (B < 5 Or MILESIAN_IS_LONG_YEAR(Year))) Then
	  A = PosDiv (Year, 4) - PosDiv(Year, 100) + PosDiv(Year, 400) 'Sum non-long terms: leap days
	  D = Year            'Force long-integer conversion
	  D = D * 365      'Begin computation of days in long-integer;
	  D = D - MStoPresentEra - 1 + B * 61 + M1 * 30 + A + DayInMonth 'Computations in long-integer first
	  Result = D
	Else	' Case where date elements do not build a correct milesian date
	  Goto ErrorHandler	
	End If
  Else		' Case where the date elements are outside basic values
	Goto ErrorHandler
  End If
  If Result > HighDate or Result < LowDate Then Goto ErrorHandler
  MILESIAN_DATE = Result
  Exit Function
  ErrorHandler:
	MILESIAN_DATE = InvArgMsg
End Function

Function MILESIAN_YEAR_BASE(ByVal Year As Long) 'The Year base or Doomsday of a year i.e. the date just before the 1 1m of the year
	On Error Goto ErrorHandler
	If Year <> Int(Year) Or Year < LowYear Or Year > HighYear Then Goto ErrorHandler
	Dim A As Integer, D As Long ', Result As Date ' Result is not used
	D = Year            'Force long-integer conversion
	D = D * 365      'Begin computation of days in long-integer;
	A = PosDiv (Year, 4) - PosDiv(Year, 100) + PosDiv(Year, 400)
	D = D - MStoPresentEra + A - 1           'Computations in long-integer first
	MILESIAN_YEAR_BASE = D ' used to be at 7:30 by adding + 0.3125
	Exit Function
	ErrorHandler: 
	 MILESIAN_YEAR_BASE = InvArgMsg
End Function

'#Part 4: Extract Milesian elements from Date number

Sub Milesian_DateElement(DNum As Date, Y As Long, M As Integer, Q As Integer, T As Variant)
' From DNum, a Date object, compute the milesian date Q / M / Y (day in month, month, year)
' Y is year in common era, relative value (may be 0 or negative)
' M is milesian month number, 1 to 12
' Q is number of day in month, 1 to 31
' T is positive decimal part: the time.
' This is an internal subroutine. Corresponding functions come after.
	Dim Cycle As Long, Day As Long      'Cycle is used serveral times with a different meaning each time
	Day = Int(DNum)                      'Initiate Day as highest integer lower or equal to DNum (force Dnum to its numeric expression)
	T = DNum - Day		' Time part, 0 <= T < 1
	Day = Day + MStoPresentEra
	Milesian_IntegDiv Day, 146097, Cycle, Day    'Day is day rank in 400 years period, Cycle is quadrisaeculum
	Y = Cycle * 400
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

Function MILESIAN_YEAR(TheDate)  'The milesian year (common era) for a Date argument (a series number or a string)
	On Error Goto ErrorHandler
	Dim Y As Long, M As Integer, Q As Integer, T As Variant
	'Dim NumDate as Date
	'NumDate = TheDate 'Force conversion or raise error
	Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
	MILESIAN_YEAR = Y
	Exit Function
	ErrorHandler: 
	  MILESIAN_YEAR = InvArgMsg
End Function

Function MILESIAN_MONTH(TheDate)   'The milesian month number (1-12) for a Date argument
	On Error Goto ErrorHandler
	If TheDate = InvArgMsg Then Goto ErrorHandler 'If TheDate was obtained from an erroneous computation, issue error
	Dim Y As Long, M As Integer, Q As Integer, T As Variant
	Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
	MILESIAN_MONTH = M
	Exit Function
	ErrorHandler:  
	  MILESIAN_MONTH = InvArgMsg
End Function

Function MILESIAN_DAY(TheDate)  'The day number in the milesian month for a Date argument
	On Error Goto ErrorHandler
	If TheDate = InvArgMsg Then Goto ErrorHandler 'If TheDate was obtained from an erroneous computation, issue error
	Dim Y As Long, M As Integer, Q As Integer, T As Variant
	Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
	MILESIAN_DAY = Q
	Exit Function
	ErrorHandler:  
	  MILESIAN_DAY = InvArgMsg
End Function

Function MILESIAN_TIME(TheDate)
	On Error Goto ErrorHandler
	If TheDate = InvArgMsg Then Goto ErrorHandler 'If TheDate was obtained from an erroneous computation, issue error
	Dim Y As Long, M As Integer, Q As Integer, T As Variant
	Milesian_DateElement TheDate, Y, M, Q, T   'Compute the figures of the milesian date
	MILESIAN_TIME = T
	Exit Function
	ErrorHandler:  
	  MILESIAN_TIME = InvArgMsg
End Function

Function MILESIAN_DISPLAY(TheDate, Optional Wtime) As String 'Default =False does not work
'Milesian date as a string, for a Date argument. Year is always three digits.
	On Error Goto ErrorHandler
	If TheDate = InvArgMsg Then Goto ErrorHandler 'If TheDate was obtained from an erroneous computation, issue error
	Dim Y As Long, M As Integer, Q As Integer, T As Variant, YS As String
	'Dim NumDate as Date
	'NumDate = TheDate 'Force conversion or raise error
	Milesian_DateElement TheDate, Y, M, Q, T  'Compute the figures of the milesian date
	YS = Format (Y, "000")
		'Fill year element with zeroes - if former function is not able.
		'YS = Y
		'If Y >= 0 And Y < 10 Then YS = "00" & Y
		'If Y >= 10 And Y < 100 Then YS = "0" & Y
		'If Y > -10 And Y < 0 Then YS = "-00" & (-Y)
		'If Y > -100 And Y <= -10 Then YS = "-0" & (-Y)
	MILESIAN_DISPLAY = Q & " " & M & "m " & YS
	If IsMissing(Wtime) Then Wtime = ( T <> 0 )
	If Wtime Then MILESIAN_DISPLAY = MILESIAN_DISPLAY & " " & T
	Exit Function
	ErrorHandler: 
	  MILESIAN_DISPLAY = InvArgMsg
End Function

'#Part 5: Computations on milesian months

Function MILESIAN_MONTH_SHIFT(TheDate, MonthShift As Long) 'Same date several (milesian) months later of earlier
	On Error Goto ErrorHandler 'Error comes from wrong parameter
	Dim Y As Long, M As Integer, Q As Integer, T As Variant, NumDate as Date
	Dim M1 As Long, Cycle As Long, Phase As Long
	NumDate = TheDate 'Force conversion or raise error
	Milesian_DateElement NumDate, Y, M, Q, D   'Compute the figures of the milesian date
	'Compute month rank from 1m of year 0
	M1 = (Y * 12) + MonthShift + M - 1 'In this order, Y is Long and shall be before simple Integer
	'Compute year and month rank
	Milesian_IntegDiv M1, 12, Cycle, Phase
	Y = Cycle
	M = Phase + 1
	'If Q was 31, set to end of month, else use same day number
	If (Q = 31) And (((M Mod 2) = 1) Or ((M = 12) And Not MILESIAN_IS_LONG_YEAR(Y))) Then Q = 30
	MILESIAN_MONTH_SHIFT = MILESIAN_DATE(Y, M, Q)
	Exit function
	ErrorHandler:
	  MILESIAN_MONTH_SHIFT = InvArgMsg
End Function

Function MILESIAN_MONTH_END(TheDate, MonthShift As Long) 'End of month several (milesian) months later of earlier
	On Error Goto ErrorHandler 'Error comes from wrong parameter
	Dim Y As Long, M As Integer, Q As Integer, T As Variant, NumDate as Date
	Dim M1 As Long, Cycle As Long, Phase As Long
	NumDate = TheDate 'Force conversion or raise error
	Milesian_DateElement NumDate, Y, M, Q, D   'Compute the figures of the milesian date
	'Compute month rank from 1m of year 0
	M1 = (Y * 12) + MonthShift + M - 1 'In this order, Y is Long and shall be before simple Integer
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
	  MILESIAN_MONTH_END = InvArgMsg
End Function

'#Part 6: Julian Day conversion functions

Function JULIAN_EPOCH_COUNT(TheDate)
'A dared strategy: compute directly Julian count as if a date, but convert into a double before returning
	Dim Result As Double	
	Result = TheDate + MStoJulianMinus05 + 0.5
    JULIAN_EPOCH_COUNT = Result 'TimePart + IntDate
End Function

Function JULIAN_EPOCH_DATE(Julian_Count)
'The strategy of computing directly to a Date object does not work: no Date computation inside a routine
    Dim IntDate As Long, TimePart As Double
    IntDate = Int(Julian_Count)       'Integer part of Julian Day
    TimePart = Julian_Count - IntDate 'Decimal part, i.e. time after noon
    TimePart = TimePart + 0.5 'Add, not substract, a half day
    IntDate = IntDate - MStoJulianMinus05 - 1 'Compensate full day added from above
    JULIAN_EPOCH_DATE = TimePart + IntDate
End Function

Function DAYOFWEEK_Ext(TheDate As Date, Optional DispType As Integer) 'As Integer 'Milesian way: Sunday = 0, Monday = 1, up to Saturday = 6
    Dim IntDate As Long, Start As Integer, Phase As Integer
    
    '1. Compute Start and Phase from DispType
    If IsMissing(DispType) Then DispType = 0    'This option value is not used with standard DOW routines
    Phase = 6   'The most common case: cycle starts with Sunday
    Select Case DispType
        Case 0          'The Milesian, John Conway, the simpliest to memorize
            Start = 0
        Case 1          'The Spreadsheets' standard
            Start = 1
        Case 2
            Start = 1
            Phase = Phase - 1
        Case 3
            Start = 0
            Phase = Phase - 1
        Case 11 To 17
            Start = 1
            Phase = Phase - (DispType - 10)
        Case Else
            DAYOFWEEK_Ext = InvArgMsg
            Exit function
        End Select
    
    '2. Extract Date element and compute
    IntDate = Int(TheDate)  'Convert date-time to hold date component only
    DAYOFWEEK_Ext = PosMod(IntDate + Phase, 7) + Start

End Function

Function GREGORIAN_EPACT(ByVal Year) 'Gregorian epact computed after the Milesian method www.calendriermilesien.org
Attribute GREGORIAN_EPACT.VB_Description = "Epact in the Gregorian sense for the given year"
    On Error Goto ErrorHandler
    Dim S As Integer 'Components of year 'N As Long
    If Year <> Int(Year) Or Year < LowYear Or Year > HighYear Then Goto ErrorHandler
    S = PosDiv(Year, 100)  'Milesian_IntegDiv Year, 100, S, N   'Decompose Year in centuries (S) + years in century (N)
    GREGORIAN_EPACT = PosMod((8 + 11 * PosMod(Year, 19) - S + S \ 4 + (8 * S + 13) \ 25), 30)  'Epact.
    Exit function
	ErrorHandler: 
		GREGORIAN_EPACT = InvArgMsg
End Function

Function MILESIAN_EPACT(ByVal Year) 'The Gregorian epact shifted to begin of Milesian year
Attribute MILESIAN_EPACT.VB_Description = "The moon age computed from the Gregorian epact, one day before Milesian new year"
    On Error Goto ErrorHandler
    MILESIAN_EPACT = PosMod(GREGORIAN_EPACT(Year) - 11, 30)
    Exit function
	ErrorHandler: 
		MILESIAN_EPACT = InvArgMsg   
End Function

