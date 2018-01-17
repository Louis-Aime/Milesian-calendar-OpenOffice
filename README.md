# Milesian-calendar-OpenOffice
Basic functions for the Milesian calendar for OpenOffice (LibreOffice) Calc

Copyright (c) Miletus, Louis-Aimé de Fouquières, 2018

MIT licence applies

## Installation
1. Create a new OpenOffice file, and save in ODT format.
1. Tools/Macro/Manage/Basic 
1. Set pointer to your file, open with the + sign, point "Standard" and hit "new"
1. Name the new module "Milesian" or whatever you wish
1. Delete the contents (an empty "Main" sub) and replace it with the contents of "OOCalc Milesian.bas"
1. You may edit the file to see the contents.

## Options
* OpenOffice / Security: Authorize access on request.
* Calc / Compute : Date field set to 30.12.1899 (default value). Sorry, other values are not handled.

## Using the functions
* Choose en cell and hit "insert function" near the input bar.
* You may enter the name of any of the functions. Be sure to put the right parameters in your formula.
* NB: Functions are sensitive to "1904 Calendar" (by default on MacOS in old versions of Excel)

## MilesianCalendar
Compute a system date with Milesian date elements, or retrieve Milesian date elements from a system date.

### General considerations
* All "date" results are given as a long integer. 
Choose the format of the cell as a "date" format with a long year to see the result as an ordinary date.
* Some control of Basic are presently not very effective, so we cannot check for non integer values of parameters.
* OpenOffice Calc displays BC years the same way as AD, in an ambiguous way. 
This is the reason why a chose to give date results as a long integer, representing the series number.
* Note that OpenOffice swithes from Julian to Gregorian calendar at the earliest, i.e. on 25 10m 1582 (15 Oct. 1582).
* You cannot raise errors from Basic. A date computed with wrong parameters shall appear as 20/12/-3741.
* There is no range control on dates. Calc displays properly from 1 Jan. 32767 B.C. up to 31 Dec. 32767 A.D.

### MILESIAN_IS_LONG_YEAR (Year)
Boolean, whether the year is long (366 days) or not. 
* Year, the year in question. 

A long Milesian year is just before a leap year, e.g. 2015 is a long year because 2016 is a leap year. 
With the Milesian calendar, a proposed rule is this:
years -4001, -801, 2399, 5599 etc. are *not* long. Elsewise the Gregorian rules are applied, 
e.g. 1899 is *not* long whereas 1999 is.

### MILESIAN_YEAR_BASE (Year) 
Date of the day before the 1 1m of year Y, i.e. the "doomsday".

### Other functions of this module 
They work like the standard date-time functions of OpenOffice. 

* MILESIAN_YEAR, MILESIAN_MONTH, MILESIAN_DAY: the Milesian date elements of an Excel date-time stamp.
* MILESIAN_DISPLAY (D) : a string that expresses a date in Milesian.
* MILESIAN_MONTH_END : works like MONTH.END.
* MILESIAN_MONTH_SHIFT : works like MONTH.SHIFT.
