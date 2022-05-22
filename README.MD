# VBA code to add line numbering to code lines only to be used for error handling

The codes add line numbering only to a single line of code. Code lines with _ are counted as one.

The line numbering can be used for error handling with the Erl fuction

` msgbox Err.Number & " " & Err.Description & " occured in line " & Erl()`

The code is inspired by Nepumuk.

# Functions/Subs
## AddLineNumbersToWorkbook
Returns total line numbers added to code in vba project

## AddLineNumbersToSingleCodeObject
returns total line numbers added to code of a single code object identified by the module name

## AddLineNumbersToComponent
Returns total line numbers added to code of a single code object as passed to the function

## LoadVBAReference()
Needed to set refernce to Microsoft Visual Basic for Applications Extensibility 5.3 library if not availbale in Tools - References
