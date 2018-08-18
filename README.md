# Milesian-calendar-OpenOffice
Basic functions for the Milesian calendar for OpenOffice (LibreOffice) Calc

Copyright (c) Miletus, Louis-Aimé de Fouquières, 2018

MIT licence applies

## Installation
1. Put all contents (.xba and .xlb files) on a dedicated directory (named "CalendarFunctions") of your file system.
1. Open or create an OpenOffice Calc file.
1. Menu: Tools/Macro/Manage/Basic.
1. In small window: "Manage" button.
1. In new small window, "Library" nailthumb (upper line)
1. In "location", you may choose either one of your open Calc files, or "my macros".
1. "Import" button
1. Select the dedicated directory of (1), then choose either .xlb file.
1. In next window, you should select "replace library", in particular if you update. The other option is up to you.
1. From now on, the "CalendarFunctions" library is included (or referred) in your file, or in your personal macros. 
You can see the functions and even edit them.
## Options
* OpenOffice / Security: Authorize access on request.
* Calc / Compute : Date field set to 30.12.1899 (default value). Sorry, other values are not handled.

## Using the functions
* Choose a cell and hit "insert function" near the input bar.
* You may enter the name of any of the functions. Be sure to put the right parameters in your formula.

## MilesianCalendar
Compute a date time stamp with Milesian date elements, or retrieve Milesian date elements from a time stamp.

### General considerations
* All "date" results are given as a long integer. 
Choose the format of the cell as a "date" format with a long year to see the result as an ordinary (i.e. julio-gregorian) date.
* Some control of Basic are presently not very effective, so we cannot check for non integer values of parameters.
* OpenOffice Calc displays BC years the same way as AD, in an ambiguous way. 
This is the reason why we chose to give date results as a long integer, representing the series number.
* Note that OpenOffice switches from Julian to Gregorian calendar at the earliest, i.e. on 15 Oct 1582.
* You cannot raise errors from Basic. A date computed with wrong parameters shall appear as 20/12/-3741.
* There is no range control on dates. Calc displays properly from 1 Jan 32767 B.C. up to 31 Dec 32767 A.D.

### Open Office similar functions of this module 
They work like the standard date-time functions of OpenOffice (and of Excel or other sheets BTW)

* MILESIAN_YEAR, MILESIAN_MONTH, MILESIAN_DAY: the Milesian date elements of an Excel date-time stamp.
* MILESIAN_DATE (Year, Month, Day_in_month): the time stamp (at 00:00) of a Milesian date given by its elements.
* MILESIAN_TIME: the time part of a time stamp; works with dates prior to 30/12/1899. 
* MILESIAN_DISPLAY (Date, Wtime) : a string that expresses a date in Milesian. 
If optional Wtime is 1 or missing, time part is added to string.
* MILESIAN_MONTH_END : works like MONTH.END.
* MILESIAN_MONTH_SHIFT : works like MONTH.SHIFT.

### Private functions
Milesian_IntegDiv, Milesian_IntegDivCeiling, PosDiv, PosMod, Milesian_DateElement, 
are private functions and procedures, not described here.

### MILESIAN_IS_LONG_YEAR (Year)
Boolean, whether the year is long (366 days) or not. 
* Year, the year in question. 

A long Milesian year is just before a leap year, e.g. 2015 is a long year because 2016 is a leap year. 
The Milesian calendar applies the Gregorian rules all the time, i.e. even before 1582, 
and applies a proposed supplementary rule: 
years -4001, -801, 2399, 5599 etc. every 3200 years, are *not* long. 
e.g. 1899 is *not* long whereas 1999 is.

### MILESIAN_YEAR_BASE (Year) 
Date of the day before the 1 1m of year Y, i.e. the "doomsday".

### JULIAN_EPOCH_COUNT (Date)
Decimal Julian Day from time stamp, deemed UTC date. 
* Date: the date to convert.

### JULIAN_EPOCH_DATE (Count)
Time stamp (Date type) representing the UTC Date from a fractional Julian Day.
* Count: fractional Julian Day to convert.

### DAYOFWEEK_Ext (Date, Option)
The day of the week for the Date, with another default option.
* Date: the date whose day of week is computed
* Option: a number; default or 0 means 0 = Sunday, Monday = 1, etc., Saturday = 6; 
1 is DAYOFWEEK's default option meaning 1 = Sunday, 2 = Monday, etc., Saturday = 7;
2, 3, are the same as OO Calc DAYOFWEEK's options.

## MilesianMoonPhase
Next or last mean moon. Error is +/- 6 hours for +/- 3000 years from year 2000.
### LastMoonPhase (FromDate, Moonphase)
Date of last new moon, or of other specified moon phase. Result is in Terrestrial Time.
* FromDate: Base date (deemed UTC);
* MoonPhase (0 by default): 0 for new moon, 1 for 1st quarter, 2 for full moon, 3 for last quarter.
### NextMoonPhase (FromDate, Moonphase)
Similar, but computes next moon phase.

## DateParse
This module has only a string parser, that converts a (numeric) Gregorian or Milesian date or date-time expression 
into an OO time stamp. 
### DATE_PARSE (String)
Date (time stamp) corresponding to a date expression
* String: holds the date expression. 
This parser recognises a Gregorian or Milesian date expression. 
The string is a Milesian date expression if either the month number ends with "m" (and without leading 0), 
or if the complete string begins with "M", in which case elements must be in the order year, month, date.
Date must hold three digit, and may be negative. BC years are counted in relative figures i.e. year 2 B.C. is year -1. 
Separators between date elements must be the same (comma is accepted with spaces). 
It is possible to specify only 2 date elements, but this must include the month. 
If specified, the year is 3-digit, otherwise it is considered "current year". 
If day of month is not specified, it is set to 1. 
