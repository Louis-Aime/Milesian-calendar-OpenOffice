REM  *****  BASIC  *****

Attribute VB_Name = "DateParse"
'DateParse.
'Copyright Miletus SARL 2018. www.calendriermilesien.org
'For use as an Open Office BASIC module.
'After a VBA module
'Tested under Libre Office 6.0 under MS Windows 10
'No warranty of conformance to business objectives
'Neither resale, neither transmission, neither compilation nor integration in other modules
'are permitted without authorisation
'DATE_PARSE, applied to any Date-Time expression as a string, returns the corresponding Date series number.
'This module uses MilesianCalendar
'Version of this module: M2018-12-24: adapted to date capacities of Open Office - LIbre Office version 6

Const HighDate = 11274306 'Highest date that is properly handled in Calc, high limit for DATE_PARSE
Const LowDate = -12661859 'Lowest date that is properly handled in Calc, low limit fo DATE_PARSE
Const InvArgMsg = "Err (Milesian) "	'Error message displayed in place of function result, for non-typed functions

Function DATE_PARSE(TheCell As String) 'A String is converted into a Numeric that corresponds to a Date
Attribute DATE_PARSE.VB_Description = "Extract date and time value from string. Standard or Milesian expression. Negative years possible."
'TheCell holds a valid string for a date, either in milesian, either in a specific calendar,
'or in the default calendar. In OpenOffice, the Default calendar is Julian until 4 Oct. 1582, then Gregorian.
'TheCell is supposed a numeric date string. It shall not work with months expressed by their names.
'Summary:
'   1. Prepare, convert all in uppercase, trim
'   2. Extract calendar indication when available.
'   3. Extract Time part, store in T
'   4. Split Date part into minimum 2 elements
'   5. Find Month element after its pattern ("M" month marker)
'   6. Find Year element, or else provide present year
'   7. Find Day element, or else provide "1" by default, return
'   Note: authorised delimiters are "./-". Comma followed with blank is dropped. "m" of Milesian month is recognised. Other lead to error.
'   ":" delimiter is specific to time part. May only appear in last word of string.
'   If as first character: canvas is yyy/mm/dd or similar, even M-yyy-mm-dd is possible.
'   If not as first, find and extract "m", compute d, m and year elements, call MILESIAN_DATE.
On Error Goto ErrorHandler

Dim Elem, I    'Elem: the array of elements of TheCell when splitted, I: a standard index
Dim T As Double, Y, M, D    'The date elements
Dim Yindex, Mindex, Dindex, YSign    'Index of year, month and day in Elem, -1 means "unknown"
Dim Delimiters      'The list of possible delimiters (outside " ")
Delimiters = Array("/", ".", "-")
Dim DelimNumber, BlankNumber
DelimNumbers = Array(0, 0, 0) 'Later will count how many non-blank delimiters

'1. Prepare: Convert string to Uppercase, drop blanks in excess, remove ",", "E" as they are misinterpreted
Y = 0	'We need Y to be initialized
YSign = 1	'First assumption: the sign of year should not be changed
TheCell = Trim(TheCell) 'Suppress all leading and trailing blank
TheCell = UCase(TheCell) 'All uppercase
'In Open Office: Drop initial "'"
If Left(TheCell, 1) = "'" Then TheCell = Right(TheCell, Len(TheCell)-1)
TheCell = Replace(TheCell, ", ", " ")    'Drop comma followed by blank (which is authorised)
If InStr(TheCell, ",") > 0 Then Goto ErrorHandler 'No other comma allowed in date string
If InStr(TheCell, "E") > 0 Then Goto ErrorHandler 'No "E" of a possible wrong numeric value
Elem = Split(TheCell)   'Split with " "
TheCell = ""            'Reconstruct TheCell
For I = LBound(Elem) To UBound(Elem)
    If Len(Elem(I)) > 0 Then TheCell = TheCell & Switch(Len(TheCell) = 0, "", Len(TheCell) > 0, " ") & Elem(I)
    Next I

'2. Did the user specify the calendar with the first character ?
'Implementation note: it is possible to add other tags here, e.g. "J" for julian calendar.

Dim K As String     'Calendar code, First letter of TheCell if it is a letter
Select Case Left(TheCell, 1)
    Case "M"
        K = "M"     'At least, year in first place, month without "m" in second place
        Yindex = 0
        Mindex = 1
        Dindex = -1 'Day may be not present, to be checked later
    Case "0" To "9"   'First character is one figure. Minus sign cannot be the first character if no letter is encountered.
        K = "D" 'Not specified, default, order is not known
        Yindex = -1
        Mindex = -1
        Dindex = -1
    'Other calendar could come here as other case tags
    Case Else
        K = "U" 'Specified but unknown
    End Select
If K <> "D" Then TheCell = Right(TheCell, Len(TheCell) - 1) 'Drop the first character from TheCell

'3. Extract Time element. This is necessary here since DateValue ignores the time part.
Elem = Split(TheCell)   'Split again, knowing that there is no empty element
'Check whether last element contains ":", if yes it is the "time" part, else no time part.
If InStr(Elem(UBound(Elem)), ":") > 0 Then    'Last elements holds a time string
    T = TimeValue(Elem(UBound(Elem)))
    'Drop Time value from TheCell and re-compute Elem
    TheCell = ""
    For I = LBound(Elem) To UBound(Elem) - 1
        If Len(Elem(I)) > 0 Then TheCell = TheCell & Switch(Len(TheCell) = 0, "", Len(TheCell) > 0, " ") & Elem(I)
        Next I
    Elem = Split(TheCell, " ", 4)
Else
    T = TimeValue("00:00:00")
End If
If UBound(Elem) > 2 Then Goto ErrorHandler ' Not more than 3 elements, excluding time part
'Here T holds the time part, TheCell holds the date part, Elem holds a possible detailed decomposition with at most 3 elements
    
'4. Extract Date elements, whichever the separator is
'4.1 To begin with, solve case where significant cell begins with "-" and consider this first part as year.
If Left(TheCell,1) = "-" Then
	If K="D" Then Goto ErrorHandler	'No "free" date string shall begin with a minus sign !! 
	' From here we know that there was a letter before the minus sign, so this sign is for the year only.
	Yindex = 0	'Year found (already)
	YSign = -1	'As we drop initial "-" sign, Year sign should be changed later
	TheCell = Right(TheCell, Len(TheCell) - 1)	'Drop initial "-" which is not a separator
End If
'4.2 Extract remaining part, by searching the "most successfull" separator
DelimNumber = -2 ' Means: we have not found which delimiter is used
If UBound(Elem) > LBound (Elem) Then DelimNumber = -1 'Means: the space is a possible delimiter, see the other ones
For I = LBound(Delimiters) To UBound(Delimiters)
	Elem = Split (TheCell, Delimiters(I))
	If Ubound(Elem) > LBound(Elem) Then 'We have a possible delimiter
		If DelimNumber = -2 Then 
			DelimNumber = I 'We did not have one before, now we have found
		Else 'Another delimiter seems to work. If not "-" (which is the last one), let us drop.
			If I < UBound(Delimiters) Then Goto ErrorHandler 'The string is not properly written
		End If 
	End If
	Next I
If DelimNumber = -2 Then 
	Goto ErrorHandler 'We did not find how to split
Else
	If DelimNumber = -1 Then
		Elem = Split (TheCell)
	Else
		Elem = Split (TheCell, Delimiters(DelimNumber))
	End If
End If
If UBound(Elem) > 2 Then Goto ErrorHandler	'Not more than 3 elements
If Yindex = 0 Then Elem(Yindex) = YSign * Elem(Yindex)	'Meanwhile we found the year, a negative value
	
'5. Check whether one element, and only one, is a milesian month notation:
'Search elements looking like "1M" to "12M", as long as there is no other indication
For I = LBound(Elem) To UBound(Elem)    'Examine each element
  If Right(Elem(I), 1) = "M" And ((Val(Elem(I)) > 0 And Val(Elem(I)) < 10 And Len(Elem(I)) = 2) Or ((Val(Elem(I)) >= 10 And Val(Elem(I)) <= 12 And Len(Elem(I)) = 3))) Then 'A Milesian month
    If K = "D" Then 'Calendar still not defined
        K = "M" 'Milesian calendar
        Mindex = I  'This is the month's index
        Elem(I) = Left(Elem(I), Len(Elem(I)) - 1) 'Set this element to a pure number
       Else
        Goto ErrorHandler     'Only one month indication authorised
       End If
     End If
  Next I

'6. Search for year element: negative, or stricly positive with at least three (numeric) character, whether first or last element, or non-existent
'Note: this part is not valid if we authorise month names, for other calendars
If Yindex > -1 Then	'We may already know which element is the year, let us check it
	If Not(IsNumeric(Elem(Yindex)) And (Len(Elem(Yindex)) >= 3 Or Val(Elem(Yindex)) <= 0)) Then Goto ErrorHandler
Else 'Year still not found
  For I = LBound(Elem) To UBound(Elem)    'Examine each element
    If IsNumeric(Elem(I)) And (Len(Elem(I)) >= 3 Or Val(Elem(I)) < 0) Then    'This can represent a year
        If Yindex = I Or Yindex = -1 Then   'Year field recognised
            Yindex = I
          Else                              'Only one year field authorised
            Goto ErrorHandler
          End If
      End If
    Next I
End If
If K = "M" And Yindex = -1 And UBound(Elem) = 2 Then Goto ErrorHandler 'No positive 2-char year authorised in Milesian notation

'7. Find whether there is a Day indication, make last computations and return
Select Case K
    Case "D"	'If default format date is specified with a negative year, register as such
   		Dim Olymp as Long, Y1 As Long, D1 As Long 'New variables for this section
    	If Yindex > -1 and Val(Elem(Yindex)) < 4 Then 'Since there is an error with 01/01/0001
    		Select Case UBound(Elem)	'How many elements ? 
    			Case 1	'Two elements, year is known, other element is month
					Mindex = (Yindex+1) Mod 2
				Case 2	'Three elements, order is either Y-M-D or D-M-Y, else error (this is a non-US version)
					Mindex = 1	'In the middle. Or to be determined from the Locale.
					Dindex = Switch (Yindex=0, 2, Yindex=2, 0)
				End Select
			If Dindex > -1 Then
            	If IsNumeric(Elem(Dindex)) Then
           			D = Val(Elem(Dindex))
	            Else
	                Goto ErrorHandler
	            End If
			Else
				D = 1
			End If
    		Milesian_IntegDiv Val(Elem(Yindex))-4, 4, Olymp, Y1
    		Y1 = Y1 + 4
    		D1 = DateValue(D & "/" & Elem(Mindex) & "/" & "00" & Y1)	'Computed a shifted date as a long integer
    		D1 = D1 + 1461*Olymp
    	Else
    		D1 = DateValue(TheCell) 'Force conversion of the date to a number
        End If
        If D1 < LowDate Or D1 > HighDate Then Goto ErrorHandler
       	DATE_PARSE = D1 + T    ' Return always a Double
	
    Case "M"    'At this level, Mindex is known  (>-1). Find and check other elements.
    	Dim ComputedDate
        If Mindex > 1 Then Goto ErrorHandler  'Month may never be indicated as 3rd element
        M = Val(Elem(Mindex))
        'If M <> Int(M) Then Goto ErrorHandler
        If Yindex = -1 Then 'Year is not specified, provide with today's date
            Y = MILESIAN_YEAR(Date) 'Today's date
        Else	'Place of Year is known, test whether it is a valid year
        	If IsNumeric(Elem(Yindex)) Then
            	Y = Val(Elem(Yindex)) 'The Val function ignores the comma !
            	'If Y <> Int(Y) Then Goto ErrorHandler
            Else
            	Goto ErrorHandler
            End If
        End If
        'Find place of day or set default day
        I = LBound(Elem)
        Do While Dindex = -1 And I <= UBound(Elem)
            If I <> Yindex And I <> Mindex Then Dindex = I  'Found
            I = I + 1
        Loop
        If Dindex > -1 Then 'D found
            If IsNumeric(Elem(Dindex)) Then
                D = Val(Elem(Dindex))
                'If D <> Int(D) Then Goto ErrorHandler
            Else
                Goto ErrorHandler
            End If
        Else    'D was not specified
            D = 1
        End If
        ComputedDate = MILESIAN_DATE(Y, M, D)
        If ComputedDate = InvArgMsg Then Goto ErrorHandler
        DATE_PARSE = ComputedDate + T
    Case Else
        Goto ErrorHandler
    End Select
  Exit Function
ErrorHandler:
	DATE_PARSE = InvArgMsg
End Function
