import * as vscode from 'vscode';
import { collectAllSymbols } from './includeProvider';
import { isCursorInHtmlFileLinkAttribute } from '../utils/htmlLinkUtils';
import { COM_MEMBER_DOCS } from '../constants/comObjects';
import { getZone, isInsideAspBlock } from '../utils/aspUtils';
import * as path from 'path';

// ─────────────────────────────────────────────────────────────────────────────
// VBScript keyword docs for hover
// ─────────────────────────────────────────────────────────────────────────────
const KEYWORD_DOCS: Record<string, string> = {

    // ── Declarations ──────────────────────────────────────────────────────────
    'dim':     '**Dim** — Declares one or more variables.\n\nExample: `Dim name, age`',
    'redim':   '**ReDim** — Resizes a dynamic array.\n\nExample: `ReDim arr(10)`',
    'set':     '**Set** — Assigns an object reference to a variable.\n\nExample: `Set rs = Server.CreateObject("ADODB.Recordset")`',
    'const':   '**Const** — Declares a constant value that cannot change.\n\nExample: `Const MAX = 100`',

    // ── Conditionals ──────────────────────────────────────────────────────────
    'if':          '**If** — Opens a conditional block.\n\nExample: `If x > 0 Then ... ElseIf ... Else ... End If`',
    'then':        '**Then** — Follows the condition in an `If` or `ElseIf` statement.\n\nExample: `If x > 0 Then`',
    'elseif':      '**ElseIf** — Additional condition branch inside an `If` block.\n\nExample: `ElseIf x = 0 Then`',
    'else':        '**Else** — Fallback branch when no `If` or `ElseIf` condition matched.\n\nExample: `Else\n    x = 0\nEnd If`',
    'end if':      '**End If** — Closes an `If` block.',
    'select case': '**Select Case** — Multi-branch conditional on a single expression.\n\nExample: `Select Case x\n    Case 1 ...\n    Case Else ...\nEnd Select`',
    'end select':  '**End Select** — Closes a `Select Case` block.',
    'case':        '**Case** — A branch inside a `Select Case` block.\n\nExample: `Case 1\n    ...`',
    'case else':   '**Case Else** — The fallback branch inside a `Select Case` block, matching anything not caught by other `Case` values.\n\nExample: `Case Else\n    label = "Unknown"`',

    // ── For loops ─────────────────────────────────────────────────────────────
    'for':      '**For** — Counter-based loop.\n\nExample: `For i = 1 To 10 Step 1 ... Next`',
    'for each': '**For Each** — Iterates over every item in a collection or array.\n\nExample: `For Each item In collection ... Next`',
    'to':       '**To** — Defines the upper bound in a `For` loop.\n\nExample: `For i = 1 To 10`',
    'step':     '**Step** — Defines the increment in a `For` loop.\n\nExample: `For i = 10 To 1 Step -1`',
    'next':     '**Next** — Closes a `For` or `For Each` loop.\n\nExample: `For i = 1 To 10 ... Next`',
    'each':     '**Each** — Used in `For Each` to iterate a collection.\n\nExample: `For Each item In collection`',
    'in':       '**In** — Separates the loop variable from the collection in `For Each`.\n\nExample: `For Each item In collection`',
    'exit for': '**Exit For** — Exits a `For` or `For Each` loop immediately.\n\nExample: `If done Then Exit For`',

    // ── Do / Loop ─────────────────────────────────────────────────────────────
    'do':           '**Do** — Opens a `Do` loop. Can have a pre- or post-condition, or loop forever.\n\nExample: `Do While condition ... Loop`',
    'do while':     '**Do While** — Loops while a condition is true (pre-condition check).\n\nExample: `Do While Not rs.EOF ... Loop`',
    'do until':     '**Do Until** — Loops until a condition becomes true (pre-condition check).\n\nExample: `Do Until cursor = 10 ... Loop`',
    'loop':         '**Loop** — Closes a `Do` block.\n\nExample: `Do ... Loop`',
    'loop while':   '**Loop While** — Closes a `Do` block and repeats while condition is true (post-condition check).\n\nExample: `Do ... Loop While x > 0`',
    'loop until':   '**Loop Until** — Closes a `Do` block and repeats until condition becomes true (post-condition check).\n\nExample: `Do ... Loop Until x >= 5`',
    'exit do':      '**Exit Do** — Exits a `Do` loop immediately.\n\nExample: `If done Then Exit Do`',

    // ── While / Wend ──────────────────────────────────────────────────────────
    'while': '**While** — Condition-based loop. Prefer `Do While` for new code.\n\nExample: `While condition ... Wend`',
    'wend':  '**Wend** — Closes a `While` loop.\n\nExample: `While condition ... Wend`',

    // ── Functions and Subs ────────────────────────────────────────────────────
    'function':     '**Function** — Declares a function that returns a value.\n\nExample: `Function GetName(id) ... End Function`',
    'sub':          '**Sub** — Declares a subroutine that does not return a value.\n\nExample: `Sub ConnectDb() ... End Sub`',
    'end function': '**End Function** — Closes a `Function` block.',
    'end sub':      '**End Sub** — Closes a `Sub` block.',
    'exit function':'**Exit Function** — Exits a `Function` early.\n\nExample: `If b = 0 Then Exit Function`',
    'exit sub':     '**Exit Sub** — Exits a `Sub` early.\n\nExample: `If Not flag Then Exit Sub`',

    // ── With ──────────────────────────────────────────────────────────────────
    'with':      '**With** — Shorthand for repeated access to an object\'s members.\n\nExample: `With rs\n    .MoveNext\nEnd With`',
    'end with':  '**End With** — Closes a `With` block.',

    // ── Class ─────────────────────────────────────────────────────────────────
    'class':     '**Class** — Declares a VBScript class.\n\nExample: `Class MyClass ... End Class`',
    'end class': '**End Class** — Closes a `Class` block.',

    // ── Error handling ────────────────────────────────────────────────────────
    'on error resume next': '**On Error Resume Next** — Suppresses runtime errors and continues execution. Always check `Err.Number` after suspicious calls.\n\nExample: `On Error Resume Next\nconn.Open ...\nIf Err.Number <> 0 Then ...`',
    'on error goto':        '**On Error GoTo 0** — Re-enables normal error handling after `On Error Resume Next`.\n\nExample: `On Error GoTo 0`',
    'goto 0':               '**On Error GoTo 0** — Re-enables normal error handling after `On Error Resume Next`.\n\nExample: `On Error GoTo 0`',
    // Middle/tail words of the 3-word compounds — resolved to the full compound doc
    'error resume':  '**On Error Resume Next** — Suppresses runtime errors and continues execution. Always check `Err.Number` after suspicious calls.',
    'resume next':   '**On Error Resume Next** — Suppresses runtime errors and continues execution. Always check `Err.Number` after suspicious calls.',
    'error goto':    '**On Error GoTo 0** — Re-enables normal error handling after `On Error Resume Next`.',

    // ── Option ────────────────────────────────────────────────────────────────
    'option explicit': '**Option Explicit** — Forces all variables to be declared with `Dim`. Recommended to prevent typo bugs.',
};

// ─────────────────────────────────────────────────────────────────────────────
// Built-in VBScript function hover docs
// ─────────────────────────────────────────────────────────────────────────────
const BUILTIN_FUNCTION_DOCS: Record<string, string> = {
    'abs':             '**Abs(number)** — Returns the absolute value of a number.',
    'array':           '**Array(arglist)**\n\nReturns a Variant containing an array.\n\n**Parameters:**\n- **arglist** — Comma-delimited list of values to assign to the elements of the array. If omitted, an empty array is created.',
    'asc':             '**Asc(string)** — Returns the ANSI character code of the first character in a string.',
    'atn':             '**Atn(number)** — Returns the arctangent of a number (in radians).',
    'cbool':           '**CBool(expression)** — Converts an expression to a Boolean.',
    'cbyte':           '**CByte(expression)** — Converts an expression to a Byte.',
    'ccur':            '**CCur(expression)** — Converts an expression to Currency.',
    'cdate':           '**CDate(expression)** — Converts an expression to a Date.',
    'cdbl':            '**CDbl(expression)** — Converts an expression to a Double.',
    'chr':             '**Chr(charcode)** — Returns the character associated with an ANSI character code.',
    'cint':            '**CInt(expression)** — Converts an expression to an Integer.',
    'clng':            '**CLng(expression)** — Converts an expression to a Long.',
    'cos':             '**Cos(number)** — Returns the cosine of an angle (in radians).',
    'createobject':    '**CreateObject(servername.typename)** — Creates and returns a reference to an Automation object.',
    'csng':            '**CSng(expression)** — Converts an expression to a Single.',
    'cstr':            '**CStr(expression)** — Converts an expression to a String.',
    'date':            '**Date()** — Returns the current system date.',
    'dateadd':         '**DateAdd(interval, number, date)**\n\nReturns a date to which a specified time interval has been added.\n\n**Parameters:**\n- **interval** — The interval to add. `"yyyy"` year, `"q"` quarter, `"m"` month, `"y"` day of year, `"d"` day, `"w"` weekday, `"ww"` week, `"h"` hour, `"n"` minute, `"s"` second\n- **number** — Number of intervals to add. Positive = future, negative = past\n- **date** — The starting date',
    'datediff':        '**DateDiff(interval, date1, date2[, firstdayofweek[, firstweekofyear]])**\n\nReturns the number of intervals between two dates.\n\n**Parameters:**\n- **interval** — The interval to measure. Same values as `DateAdd`\n- **date1** — The earlier date\n- **date2** — The later date\n- **firstdayofweek** — `1` Sunday (default), `2` Monday, `3` Tuesday, `4` Wednesday, `5` Thursday, `6` Friday, `7` Saturday\n- **firstweekofyear** — `1` week containing Jan 1 (default), `2` first week with 4+ days, `3` first full week',
    'datepart':        '**DatePart(interval, date[, firstdayofweek[, firstweekofyear]])**\n\nReturns the specified part of a given date.\n\n**Parameters:**\n- **interval** — The part to return. Same values as `DateAdd`\n- **date** — The date to evaluate\n- **firstdayofweek** — `1` Sunday (default), `2` Monday, `3` Tuesday, `4` Wednesday, `5` Thursday, `6` Friday, `7` Saturday\n- **firstweekofyear** — `1` week containing Jan 1 (default), `2` first week with 4+ days, `3` first full week',
    'dateserial':      '**DateSerial(year, month, day)**\n\nReturns a Variant of subtype Date for a specified year, month, and day.\n\n**Parameters:**\n- **year** — Full four-digit year e.g. `2024`. Values 0–99 are interpreted as 1900–1999\n- **month** — Month as a number `1`–`12`. Values outside this range roll over e.g. `13` = January of next year\n- **day** — Day as a number `1`–`31`. Values outside this range roll over e.g. `32` = first day of next month',
    'datevalue':       '**DateValue(date)** — Returns a Variant of subtype Date.',
    'day':             '**Day(date)** — Returns a whole number between 1 and 31 representing the day of the month.',
    'exp':             '**Exp(number)** — Returns e (the base of natural logarithms) raised to a power.',
    'filter':          '**Filter(InputStrings, Value[, Include[, Compare]])**\n\nReturns a zero-based array containing a subset of a string array.\n\n**Parameters:**\n- **InputStrings** — One-dimensional array of strings to search\n- **Value** — The string to search for\n- **Include** — `True` (default) return elements that contain Value, `False` return elements that do not contain Value\n- **Compare** — `0` (vbBinaryCompare) case-sensitive, `1` (vbTextCompare) case-insensitive',
    'fix':             '**Fix(number)** — Returns the integer portion of a number (truncates toward zero).',
    'formatcurrency':  '**FormatCurrency(Expression[, NumDigitsAfterDecimal[, IncludeLeadingDigit[, UseParensForNegativeNumbers[, GroupDigits]]]])**\n\nReturns an expression formatted as a currency value using the system currency symbol.\n\n**Parameters:**\n- **Expression** — The value to format\n- **NumDigitsAfterDecimal** — Number of decimal places. `-1` = use system default\n- **IncludeLeadingDigit** — `-1` (True) show leading zero e.g. `$0.50`, `0` (False) omit it e.g. `$.50`, `-2` = use system default\n- **UseParensForNegativeNumbers** — `-1` (True) wrap negatives in parentheses e.g. `($1.00)`, `0` (False) use minus sign e.g. `-$1.00`, `-2` = use system default\n- **GroupDigits** — `-1` (True) group digits e.g. `$1,000.00`, `0` (False) no grouping e.g. `$1000.00`, `-2` = use system default',
    'formatdatetime':  '**FormatDateTime(Date[, NamedFormat])**\n\nReturns an expression formatted as a date or time.\n\n**Parameters:**\n- **Date** — The date/time value to format\n- **NamedFormat** — `0` (vbGeneralDate, default) date and/or time, `1` (vbLongDate) long date e.g. `Monday, 1 January 2024`, `2` (vbShortDate) short date e.g. `01/01/2024`, `3` (vbLongTime) long time e.g. `12:00:00 AM`, `4` (vbShortTime) short time e.g. `12:00`',
    'formatnumber':    '**FormatNumber(Expression[, NumDigitsAfterDecimal[, IncludeLeadingDigit[, UseParensForNegativeNumbers[, GroupDigits]]]])**\n\nReturns an expression formatted as a number.\n\n**Parameters:**\n- **Expression** — The value to format\n- **NumDigitsAfterDecimal** — Number of decimal places. `-1` = use system default\n- **IncludeLeadingDigit** — `-1` (True) show leading zero e.g. `0.5`, `0` (False) omit it e.g. `.5`, `-2` = use system default\n- **UseParensForNegativeNumbers** — `-1` (True) wrap negatives in parentheses e.g. `(1.00)`, `0` (False) use minus sign e.g. `-1.00`, `-2` = use system default\n- **GroupDigits** — `-1` (True) group digits e.g. `1,000.00`, `0` (False) no grouping e.g. `1000.00`, `-2` = use system default',
    'formatpercent':   '**FormatPercent(Expression[, NumDigitsAfterDecimal[, IncludeLeadingDigit[, UseParensForNegativeNumbers[, GroupDigits]]]])**\n\nReturns an expression formatted as a percentage (multiplied by 100).\n\n**Parameters:**\n- **Expression** — The value to format e.g. `0.5` becomes `50%`\n- **NumDigitsAfterDecimal** — Number of decimal places. `-1` = use system default\n- **IncludeLeadingDigit** — `-1` (True) show leading zero e.g. `0.50%`, `0` (False) omit it e.g. `.50%`, `-2` = use system default\n- **UseParensForNegativeNumbers** — `-1` (True) wrap negatives in parentheses e.g. `(50.00%)`, `0` (False) use minus sign e.g. `-50.00%`, `-2` = use system default\n- **GroupDigits** — `-1` (True) group digits e.g. `1,000.00%`, `0` (False) no grouping, `-2` = use system default',
    'getobject':       '**GetObject([pathname[, class]])**\n\nReturns a reference to an object provided by an Automation component.\n\n**Parameters:**\n- **pathname** — Full path of the file containing the object. Omit to use the class argument alone\n- **class** — The class of the object in the form `appname.objecttype` e.g. `"Excel.Sheet"`. Required if pathname is omitted',
    'hex':             '**Hex(number)** — Returns a string representing the hexadecimal value of a number.',
    'hour':            '**Hour(time)** — Returns a whole number between 0 and 23 representing the hour of the day.',
    'inputbox':        '**InputBox(prompt[, title[, default[, xpos[, ypos]]]])**\n\nDisplays a prompt in a dialog box and returns the text the user entered.\n\n**Parameters:**\n- **prompt** — The message shown to the user\n- **title** — Title bar text. Defaults to the application name if omitted\n- **default** — Default value pre-filled in the input field. Empty if omitted\n- **xpos** — Horizontal position of the dialog in twips from the left edge of the screen\n- **ypos** — Vertical position of the dialog in twips from the top edge of the screen',
    'instr':           '**InStr([start, ]string1, string2[, compare])**\n\nReturns the position of the first occurrence of string2 within string1, or 0 if not found.\n\n**Parameters:**\n- **start** — Starting position for the search. Defaults to `1`. Required if compare is specified\n- **string1** — The string to search in\n- **string2** — The string to search for\n- **compare** — `0` (vbBinaryCompare) case-sensitive, `1` (vbTextCompare) case-insensitive',
    'instrrev':        '**InStrRev(string1, string2[, start[, compare]])**\n\nReturns the position of the last occurrence of string2 within string1 (searches from the end), or 0 if not found.\n\n**Parameters:**\n- **string1** — The string to search in\n- **string2** — The string to search for\n- **start** — Starting position for the search, counting from the left. `-1` (default) = start from the last character\n- **compare** — `0` (vbBinaryCompare) case-sensitive, `1` (vbTextCompare) case-insensitive',
    'int':             '**Int(number)** — Returns the integer portion of a number (rounds down).',
    'isarray':         '**IsArray(varname)** — Returns True if the variable is an array.',
    'isdate':          '**IsDate(expression)** — Returns True if the expression can be converted to a date.',
    'isempty':         '**IsEmpty(expression)** — Returns True if the variable is uninitialized.',
    'isnull':          '**IsNull(expression)** — Returns True if the expression is Null.',
    'isnumeric':       '**IsNumeric(expression)** — Returns True if the expression can be evaluated as a number.',
    'isobject':        '**IsObject(expression)** — Returns True if the expression references a valid object.',
    'join':            '**Join(list[, delimiter])**\n\nReturns a string created by joining the elements of an array.\n\n**Parameters:**\n- **list** — One-dimensional array whose elements will be joined\n- **delimiter** — String used to separate elements in the result. `" "` (single space) by default. Use `""` for no separator',
    'lbound':          '**LBound(arrayname[, dimension])**\n\nReturns the smallest available subscript for the specified dimension of an array.\n\n**Parameters:**\n- **arrayname** — Name of the array\n- **dimension** — Which dimension to check. `1` (default) = first dimension, `2` = second dimension, etc.',
    'lcase':           '**LCase(string)** — Returns a string converted to lowercase.',
    'left':            '**Left(string, length)** — Returns a specified number of characters from the left side of a string.',
    'len':             '**Len(string | varname)** — Returns the number of characters in a string or bytes required to store a variable.',
    'log':             '**Log(number)** — Returns the natural logarithm of a number.',
    'ltrim':           '**LTrim(string)** — Returns a copy of a string without leading spaces.',
    'mid':             '**Mid(string, start[, length])**\n\nReturns a specified number of characters from a string.\n\n**Parameters:**\n- **string** — The source string\n- **start** — Position of the first character to return. `1` = first character\n- **length** — Number of characters to return. If omitted or longer than the string, returns all characters from start to the end',
    'minute':          '**Minute(time)** — Returns a whole number between 0 and 59 representing the minute of the hour.',
    'month':           '**Month(date)** — Returns a whole number between 1 and 12 representing the month of the year.',
    'monthname':       '**MonthName(month[, abbreviate])**\n\nReturns a string indicating the specified month.\n\n**Parameters:**\n- **month** — Month number `1`–`12`\n- **abbreviate** — `True` return abbreviated name e.g. `"Jan"`, `False` (default) return full name e.g. `"January"`',
    'msgbox':          '**MsgBox(prompt[, buttons[, title]])**\n\nDisplays a message in a dialog box and returns a value indicating which button was clicked.\n\n**Parameters:**\n- **prompt** — The message to display\n- **buttons** — Controls which buttons and icon appear. Common values: `0` OK only, `1` OK + Cancel, `2` Abort + Retry + Ignore, `3` Yes + No + Cancel, `4` Yes + No, `5` Retry + Cancel. Add `16` for critical icon, `32` for question icon, `48` for warning icon, `64` for info icon\n- **title** — Title bar text. Defaults to the application name if omitted\n\n**Return values:** `1` OK, `2` Cancel, `3` Abort, `4` Retry, `5` Ignore, `6` Yes, `7` No',
    'now':             '**Now()** — Returns the current system date and time.',
    'oct':             '**Oct(number)** — Returns a string representing the octal value of a number.',
    'replace':         '**Replace(expression, find, replacewith[, start[, count[, compare]]])**\n\nReturns a string with all (or a limited number of) occurrences of a substring replaced.\n\n**Parameters:**\n- **expression** — The source string\n- **find** — The substring to search for\n- **replacewith** — The substring to replace with\n- **start** — Position in expression to begin searching. `1` (default) = from the beginning. Note: the returned string always begins at this position\n- **count** — Number of replacements to make. `-1` (default) = replace all occurrences\n- **compare** — `0` (vbBinaryCompare) case-sensitive, `1` (vbTextCompare) case-insensitive',
    'rgb':             '**RGB(red, green, blue)** — Returns a whole number representing an RGB colour value.',
    'right':           '**Right(string, length)** — Returns a specified number of characters from the right side of a string.',
    'rnd':             '**Rnd([number])**\n\nReturns a random Single between 0 and 1. Call `Randomize` first for a different sequence each run.\n\n**Parameters:**\n- **number** — `< 0` always returns the same number for the same seed, `> 0` or omitted returns the next random number in the sequence, `= 0` returns the most recently generated number',
    'round':           '**Round(expression[, numdecimalplaces])**\n\nReturns a number rounded to a specified number of decimal places. Uses banker\'s rounding (rounds to even) on .5.\n\n**Parameters:**\n- **expression** — The value to round\n- **numdecimalplaces** — Number of decimal places to keep. `0` (default) = round to whole number',
    'rtrim':           '**RTrim(string)** — Returns a copy of a string without trailing spaces.',
    'second':          '**Second(time)** — Returns a whole number between 0 and 59 representing the second of the minute.',
    'sgn':             '**Sgn(number)** — Returns an integer indicating the sign of a number.',
    'sin':             '**Sin(number)** — Returns the sine of an angle (in radians).',
    'space':           '**Space(number)** — Returns a string consisting of the specified number of spaces.',
    'split':           '**Split(expression[, delimiter[, count[, compare]]])**\n\nReturns a zero-based array of substrings split from a string.\n\n**Parameters:**\n- **expression** — The string to split\n- **delimiter** — String used to identify splits. `" "` (single space) by default\n- **count** — Maximum number of substrings to return. `-1` (default) = return all substrings\n- **compare** — `0` (vbBinaryCompare) case-sensitive, `1` (vbTextCompare) case-insensitive',
    'sqr':             '**Sqr(number)** — Returns the square root of a number.',
    'strcomp':         '**StrComp(string1, string2[, compare])**\n\nReturns a value indicating the result of a string comparison.\n\n**Parameters:**\n- **string1** — First string\n- **string2** — Second string\n- **compare** — `0` (vbBinaryCompare) case-sensitive, `1` (vbTextCompare) case-insensitive\n\n**Return values:** `-1` string1 < string2, `0` string1 = string2, `1` string1 > string2, `Null` if either string is Null',
    'string':          '**String(number, character)**\n\nReturns a string made up of a character repeated a specified number of times.\n\n**Parameters:**\n- **number** — How many times to repeat the character\n- **character** — The character to repeat. Can be a character code (e.g. `42`) or a string — only the first character is used',
    'strreverse':      '**StrReverse(string)** — Returns the reverse of a string.',
    'tan':             '**Tan(number)** — Returns the tangent of an angle (in radians).',
    'time':            '**Time()** — Returns the current system time.',
    'timer':           '**Timer()** — Returns the number of seconds elapsed since midnight.',
    'timeserial':      '**TimeSerial(hour, minute, second)**\n\nReturns a Variant of subtype Date containing the time for a specific hour, minute, and second.\n\n**Parameters:**\n- **hour** — `0`–`23`. Values outside this range roll over e.g. `24` = midnight of next day\n- **minute** — `0`–`59`. Values outside this range roll over\n- **second** — `0`–`59`. Values outside this range roll over',
    'timevalue':       '**TimeValue(time)** — Returns a Variant of subtype Date containing the time.',
    'trim':            '**Trim(string)** — Returns a copy of a string without leading or trailing spaces.',
    'typename':        '**TypeName(varname)** — Returns a string that describes the subtype of a variable.',
    'ubound':          '**UBound(arrayname[, dimension])**\n\nReturns the largest available subscript for the specified dimension of an array.\n\n**Parameters:**\n- **arrayname** — Name of the array\n- **dimension** — Which dimension to check. `1` (default) = first dimension, `2` = second dimension, etc.',
    'ucase':           '**UCase(string)** — Returns a string converted to uppercase.',
    'vartype':         '**VarType(varname)** — Returns a value indicating the subtype of a variable.',
    'weekday':         '**Weekday(date[, firstdayofweek])**\n\nReturns a whole number representing the day of the week.\n\n**Parameters:**\n- **date** — The date to evaluate\n- **firstdayofweek** — Which day is considered day 1. `1` (vbSunday, default), `2` (vbMonday), `3` (vbTuesday), `4` (vbWednesday), `5` (vbThursday), `6` (vbFriday), `7` (vbSaturday)\n\n**Return values:** `1`–`7` depending on firstdayofweek setting',
    'weekdayname':     '**WeekdayName(weekday[, abbreviate[, firstdayofweek]])**\n\nReturns a string indicating the specified day of the week.\n\n**Parameters:**\n- **weekday** — Day number as returned by the `Weekday` function\n- **abbreviate** — `True` return abbreviated name e.g. `"Mon"`, `False` (default) return full name e.g. `"Monday"`\n- **firstdayofweek** — Which day is considered day 1. `1` (vbSunday, default), `2` (vbMonday), through `7` (vbSaturday)',
    'year':            '**Year(date)** — Returns a whole number representing the year.',
};

// 3-word compound keyword sequences — checked before 2-word compounds.
// Each entry maps the lowercase 3-word key to the KEYWORD_DOCS key to use.
const THREE_WORD_COMPOUNDS: Record<string, string> = {
    'on error resume':   'on error resume next',
    'error resume next': 'on error resume next',
    'on error goto':     'on error goto',
    'error goto 0':      'on error goto',
};


// ─────────────────────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────────────────
// Context detection
// Uses the canonical isInsideAspBlock from aspUtils (comment + string aware)
// and a simple script-block scan for the JS zone.
// ─────────────────────────────────────────────────────────────────────────────

type AspContext = 'vbscript' | 'script' | 'html';

function getAspContext(document: vscode.TextDocument, position: vscode.Position): AspContext {
    const fullText = document.getText();
    const offset   = document.offsetAt(position);

    // VBScript block — use the canonical comment/string-aware scanner.
    if (isInsideAspBlock(fullText, offset)) { return 'vbscript'; }

    // JavaScript <script> block — scan for the nearest unclosed <script> tag
    // that is NOT a VBScript block (language="vbscript").
    let searchFrom = 0;
    while (true) {
        const scriptOpen = fullText.indexOf('<script', searchFrom);
        if (scriptOpen === -1 || scriptOpen >= offset) { break; }
        const tagEnd     = fullText.indexOf('>', scriptOpen);
        if (tagEnd === -1) { break; }
        const scriptTag  = fullText.slice(scriptOpen, tagEnd + 1);
        const scriptClose = fullText.indexOf('</script', tagEnd);
        if (
            !/language\s*=\s*["']vbscript["']/i.test(scriptTag) &&
            tagEnd < offset &&
            (scriptClose === -1 || offset <= scriptClose)
        ) {
            return 'script';
        }
        searchFrom = scriptClose === -1 ? fullText.length : scriptClose + 9;
    }

    return 'html';
}

// ─────────────────────────────────────────────────────────────────────────────
// Hover provider
// • VBScript context (<% %>): all hovers — keywords, symbols, COM members.
// • Script context (<script>):  symbol/COM hovers only, no keyword docs.
// • HTML context:               no hovers (except HTML link guard already applied).
// ─────────────────────────────────────────────────────────────────────────────
export class AspHoverProvider implements vscode.HoverProvider {

    provideHover(
        document: vscode.TextDocument,
        position: vscode.Position
    ): vscode.ProviderResult<vscode.Hover> {

        const lineText = document.lineAt(position.line).text;

        // Suppress hover inside HTML file-link attributes (href, src, etc.)
        if (isCursorInHtmlFileLinkAttribute(lineText, position.character)) return null;

        const context = getAspContext(document, position);

        // No hovers inside plain HTML
        if (context === 'html') return null;

        const wordRange = document.getWordRangeAtPosition(position, /\w+/);
        if (!wordRange) return null;

        const word    = document.getText(wordRange);
        const wordKey = word.toLowerCase();

        // ── Suppress hover when cursor is inside a string literal ────────────
        // VBScript strings are delimited by ".  Scan the line up to the cursor,
        // tracking open/close quotes ("" is an escaped quote inside a string).
        // If the cursor lands inside a string the word is a value, not an
        // identifier — so Case "Active", Response.Write "msg", etc. must never
        // show variable/function/keyword hovers.
        {
            let inStr = false;
            const col = position.character;
            for (let ci = 0; ci < col; ci++) {
                if (lineText[ci] === '"') {
                    if (inStr && lineText[ci + 1] === '"') { ci++; continue; } // escaped ""
                    inStr = !inStr;
                }
            }
            if (inStr) return null;
        }

        const allSymbols = collectAllSymbols(document);
        const docText    = document.getText();

        // ── 1. COM member after dot — e.g. rs.EOF, conn.Execute ──────────────
        const charBeforeWord = lineText.charAt(wordRange.start.character - 1);
        if (charBeforeWord === '.') {
            const textBeforeDot = lineText.substring(0, wordRange.start.character - 1);
            const objMatch      = textBeforeDot.match(/\b(\w+)$/);
            if (objMatch) {
                const objName    = objMatch[1].toLowerCase();
                const comVar     = allSymbols.comVariables.find(cv => cv.name.toLowerCase() === objName);
                if (comVar) {
                    const memberDoc = COM_MEMBER_DOCS[`${comVar.progId}.${wordKey}`];
                    if (memberDoc) {
                        return new vscode.Hover(
                            new vscode.MarkdownString(`**${memberDoc.label}**\n\n${memberDoc.doc}`)
                        );
                    }
                }
            }
        }

        // ── 2. User-defined Function or Sub ───────────────────────────────────
        // Skip functions defined inside a <script> (JS) block in this file —
        // they are JavaScript, not VBScript, and should not show a VBScript hover.
        // Functions from #include files are always VBScript so they are shown as-is.
        const fn = allSymbols.functions.find(f => {
            if (f.name.toLowerCase() !== wordKey) return false;
            if (f.filePath === document.uri.fsPath) {
                const fnOffset = document.offsetAt(new vscode.Position(f.line, 0));
                if (getZone(docText, fnOffset) === 'js') return false;
            }
            return true;
        });
        if (fn) {
            const fromInclude = fn.filePath !== document.uri.fsPath;
            const header      = fn.params
                ? `**${fn.kind} ${fn.name}(${fn.params})**`
                : `**${fn.kind} ${fn.name}**`;
            const source      = fromInclude
                ? `\n\n*Defined in \`${path.basename(fn.filePath)}\`*`
                : `\n\n*Defined in this file*`;
            return new vscode.Hover(new vscode.MarkdownString(header + source));
        }

        // ── 3. COM object variable (rs, conn, dict, etc.) ─────────────────────
        const comVar = allSymbols.comVariables.find(cv => cv.name.toLowerCase() === wordKey);
        if (comVar) {
            const fromInclude = comVar.filePath !== document.uri.fsPath;
            const source      = fromInclude
                ? `*Declared in \`${path.basename(comVar.filePath)}\`*`
                : `*Declared in this file*`;
            return new vscode.Hover(
                new vscode.MarkdownString(
                    `**${comVar.name}** — \`${comVar.progId}\`\n\n${source}\n\nType \`${comVar.name}.\` to see available members.`
                )
            );
        }

        // ── 4. User-defined variable ──────────────────────────────────────────
        const variable = allSymbols.variables.find(v => v.name.toLowerCase() === wordKey);
        if (variable) {
            const fromInclude = variable.filePath !== document.uri.fsPath;
            const source      = fromInclude
                ? `*Declared in \`${path.basename(variable.filePath)}\`*`
                : `*Declared in this file*`;
            return new vscode.Hover(new vscode.MarkdownString(`**${variable.name}** — variable\n\n${source}`));
        }

        // ── 5. User-defined constant ──────────────────────────────────────────
        const constant = allSymbols.constants.find(c => c.name.toLowerCase() === wordKey);
        if (constant) {
            const fromInclude = constant.filePath !== document.uri.fsPath;
            const source      = fromInclude
                ? `*Declared in \`${path.basename(constant.filePath)}\`*`
                : `*Declared in this file*`;
            return new vscode.Hover(
                new vscode.MarkdownString(`**${constant.name}** = \`${constant.value}\`\n\n${source}`)
            );
        }

        // ── 6. Built-in VBScript function hover ─────────────────────────────────
        // Show docs for built-in functions like Split(), InStr(), DateDiff(), etc.
        if (BUILTIN_FUNCTION_DOCS[wordKey]) {
            return new vscode.Hover(new vscode.MarkdownString(BUILTIN_FUNCTION_DOCS[wordKey]));
        }

        // ── 7. VBScript keywords — only inside <% %> blocks ──────────────────
        // Keyword docs are VBScript-specific so we suppress them inside <script>.
        if (context !== 'vbscript') return null;

        // Suppress hover inside comments. Strip string literals first so a quote
        // inside a string isn't mistaken for a comment delimiter.
        const strippedForComment = lineText.replace(/"[^"]*"/g, m => ' '.repeat(m.length));
        const commentIdx         = strippedForComment.indexOf("'");
        if (commentIdx !== -1 && position.character > commentIdx) return null;

        // Extract words immediately before and after the hovered word so we can
        // assemble 2-word and 3-word compound keys and return the correct doc
        // with a range that spans the entire compound, not just the hovered word.

        const textBefore      = lineText.substring(0, wordRange.start.character);
        const textAfter       = lineText.substring(wordRange.end.character);
        const wordBeforeMatch = textBefore.match(/\b(\w+)(\s+)$/);
        const wordAfterMatch  = textAfter.match(/^(\s+)(\w+)\b/);
        const twoBeforeMatch  = wordBeforeMatch
            ? textBefore.substring(0, textBefore.length - wordBeforeMatch[0].length).match(/\b(\w+)(\s+)$/)
            : null;
        const twoAfterMatch   = wordAfterMatch
            ? textAfter.substring(wordAfterMatch[0].length).match(/^(\s+)(\w+)\b/)
            : null;

        const wBefore  = wordBeforeMatch?.[1]?.toLowerCase() ?? '';
        const wAfter   = wordAfterMatch?.[2]?.toLowerCase()  ?? '';
        const wwBefore = twoBeforeMatch?.[1]?.toLowerCase()  ?? '';
        const wwAfter  = twoAfterMatch?.[2]?.toLowerCase()   ?? '';

        // Helper: return a Hover with a range spanning from startCol to endCol
        const makeHover = (doc: string, startCol: number, endCol: number) =>
            new vscode.Hover(
                new vscode.MarkdownString(doc),
                new vscode.Range(
                    new vscode.Position(position.line, startCol),
                    new vscode.Position(position.line, endCol)
                )
            );

        const wStart  = wordRange.start.character;
        const wEnd    = wordRange.end.character;
        const bStart  = wordBeforeMatch  ? wStart  - wordBeforeMatch[0].length  : wStart;
        const bbStart = twoBeforeMatch   ? bStart  - twoBeforeMatch[0].length   : bStart;
        const aEnd    = wordAfterMatch   ? wEnd    + wordAfterMatch[0].length    : wEnd;
        const aaEnd   = twoAfterMatch    ? aEnd    + twoAfterMatch[0].length     : aEnd;

        // ── 3-word compounds ─────────────────────────────────────────────────
        // Check word3 first (cursor on last word), then word2, then word1.
        // This ensures the widest possible range is always returned — e.g.
        // hovering "GoTo" should span "On Error GoTo 0", not just "Error GoTo 0".

        // Helper: if a compound ends with 'goto', extend the range to include
        // a trailing '0' (On Error GoTo 0) since 0 is part of the statement.
        // Also extends 'on error resume' to include trailing 'Next'.
        const extendForGoTo = (docKey: string, endCol: number): number => {
            const tail = lineText.substring(endCol);
            if (docKey.endsWith('goto')) {
                const m = tail.match(/^(\s+)(0)\b/);
                return m ? endCol + m[0].length : endCol;
            }
            if (docKey === 'on error resume next') {
                const m = tail.match(/^(\s+)(next)\b/i);
                return m ? endCol + m[0].length : endCol;
            }
            return endCol;
        };

        // Cursor on word 3: prev 2 + hovered  (e.g. hover "Resume" in "On Error Resume")
        // Also try to look one level further back in case this 3-word compound is itself
        // the tail of a longer known compound (e.g. "error goto 0" inside "on error goto 0").
        if (wwBefore && wBefore) {
            const k = `${wwBefore} ${wBefore} ${wordKey}`;
            const docKey = THREE_WORD_COMPOUNDS[k];
            if (docKey && KEYWORD_DOCS[docKey]) {
                let startCol = bbStart;
                // Look one more word back — if it extends to a known 3-word compound
                // that shares the same docKey, widen to include it too.
                const textBeforeBb = lineText.substring(0, wordRange.start.character
                    - wordBeforeMatch![0].length
                    - twoBeforeMatch![0].length);
                const threeBackMatch = textBeforeBb.match(/\b(\w+)(\s+)$/);
                if (threeBackMatch) {
                    const w3 = threeBackMatch[1].toLowerCase();
                    const wider = THREE_WORD_COMPOUNDS[`${w3} ${wwBefore} ${wBefore}`];
                    if (wider === docKey) startCol = bbStart - threeBackMatch[0].length;
                }
                return makeHover(KEYWORD_DOCS[docKey], startCol, extendForGoTo(docKey, wEnd));
            }
        }
        // Cursor on word 2: prev + hovered + next  (e.g. hover "Error" in "On Error GoTo")
        if (wBefore && wAfter) {
            const k = `${wBefore} ${wordKey} ${wAfter}`;
            const docKey = THREE_WORD_COMPOUNDS[k];
            if (docKey && KEYWORD_DOCS[docKey]) return makeHover(KEYWORD_DOCS[docKey], bStart, extendForGoTo(docKey, aEnd));
        }
        // Cursor on word 1: hovered + next 2  (e.g. hover "On" in "On Error GoTo")
        if (wAfter && wwAfter) {
            const k = `${wordKey} ${wAfter} ${wwAfter}`;
            const docKey = THREE_WORD_COMPOUNDS[k];
            if (docKey && KEYWORD_DOCS[docKey]) return makeHover(KEYWORD_DOCS[docKey], wStart, extendForGoTo(docKey, aaEnd));
        }

        // ── 2-word compounds ─────────────────────────────────────────────────
        // Case a: word before + hovered  (e.g. "End Function", "Do While", "GoTo 0")
        // Walk back up to two extra levels to widen to the full compound range.
        if (wBefore) {
            const k = `${wBefore} ${wordKey}`;
            if (KEYWORD_DOCS[k]) {
                let startCol = bStart;
                // One level back: e.g. "error goto 0" → startCol = bbStart (Error)
                if (wwBefore) {
                    const wider1 = THREE_WORD_COMPOUNDS[`${wwBefore} ${wBefore} ${wordKey}`];
                    if (wider1) {
                        startCol = bbStart;
                        // Two levels back: look for a word before wwBefore that forms a 3-word
                        // compound with wwBefore+wBefore — e.g. "on" before "error goto 0"
                        const textBeforeBb = lineText.substring(0, wordRange.start.character - wordBeforeMatch![0].length - twoBeforeMatch![0].length);
                        const threeBeforeMatch = textBeforeBb.match(/\b(\w+)(\s+)$/);
                        if (threeBeforeMatch) {
                            const ww3 = threeBeforeMatch[1].toLowerCase();
                            const wider2 = THREE_WORD_COMPOUNDS[`${ww3} ${wwBefore} ${wBefore}`];
                            if (wider2 === wider1) startCol = bbStart - threeBeforeMatch[0].length;
                        }
                    }
                }
                return makeHover(KEYWORD_DOCS[k], startCol, wEnd);
            }
        }
        // Case b: hovered + word after  (e.g. "For Each", "Loop Until")
        // Special case: GoTo + 0 — extend range to include the 0
        if (wAfter) {
            const k = `${wordKey} ${wAfter}`;
            if (KEYWORD_DOCS[k]) {
                const endCol = extendForGoTo(k, aEnd);
                return makeHover(KEYWORD_DOCS[k], wStart, endCol);
            }
        }

        // ── Single keyword ────────────────────────────────────────────────────
        if (KEYWORD_DOCS[wordKey]) {
            return new vscode.Hover(new vscode.MarkdownString(KEYWORD_DOCS[wordKey]));
        }

        return null;
    }
}