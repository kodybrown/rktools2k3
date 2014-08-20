'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' forms.vbs - form script for Windows .NET Server 2003
'
' Usage:
' forms [-vdla?] [-c server]
'
' Examples:
' forms -a -n NewForm -u inches -h 8.5 -w 11 -t 0 -e 0 -b 7 -r 8
' forms -d -n NewForm
' forms -l -c \\server
' forms -l -v -c \\server
'
'----------------------------------------------------------------------

option explicit

'
' Debugging trace flags, to enable debug output trace message
' change gDebugFlag to true.
'
const kDebugTrace = 1
const kDebugError = 2
dim gDebugFlag

gDebugFlag = false

'
' Messages to be displayed if the scripting host is not cscript
'
const kMessage1 = "Please run this script using CScript."
const kMessage2 = "This can be achieved by"
const kMessage3 = "1. Using ""CScript script.vbs arguments"" or"
const kMessage4 = "2. Changing the default Windows Scripting Host to CScript"
const kMessage5 = "   using ""CScript //H:CScript //S"" and running the script "
const kMessage6 = "   ""script.vbs arguments""."

'
' Operation action values.
'
const kActionUnknown    = 0
const kActionAdd        = 1
const kActionDelete     = 2
const kActionList       = 3

'
' Constants for units
'
const kInches           = 0
const kCentimeters      = 1

const kErrorSuccess     = 0
const kErrorFailure     =1

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strServer
    dim strForm
    dim bVerbose
    dim iUnits
    dim iHeight
    dim iWidth
    dim iTop
    dim iLeft
    dim iBottom
    dim iRight

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(kMessage1 & vbCRLF & kMessage2 & vbCRLF & _
                          kMessage3 & vbCRLF & kMessage4 & vbCRLF & _
                          kMessage5 & vbCRLF & kMessage6 & vbCRLF)

        wscript.quit

    end if

    '
    ' Inches is the default for adding or listing forms
    '
    iUnits = kInches

    iRetval = ParseCommandLine(iAction, strServer, strForm, iUnits, iHeight, _
                               iWidth,  iTop, iLeft, iBottom, iRight, bVerbose)

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionAdd
                iRetval = AddForm(strServer, strForm, iUnits, iHeight, iWidth, _
                                  iTop, iLeft, iBottom, iRight)

            case kActionDelete
                iRetval = DeleteForm(strServer, strForm)

            case kActionList
                iRetval = ListForms(strServer, iUnits, bVerbose)

            case else
                Usage(true)

        end select

    end if

end sub

'
' Add a Form
'
function AddForm(strServer, strForm, iUnits, iHeight, iWidth, iTop, iLeft, iBottom, iRight)

    on error resume next

    DebugPrint kDebugTrace, "In AddForm"
    DebugPrint kDebugTrace, "Server Name " & strServer
    DebugPrint kDebugTrace, "Form Name   " & strForm

    dim oMaster
    dim oForm
    dim iRetval

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    set oForm = CreateForm(strForm, iUnits, iHeight, iWidth, iTop, iLeft, iBottom, iRight)

    oForm.ServerName = strServer

    oMaster.FormAdd oForm

    if Err.Number = kErrorSuccess then

        wscript.echo "Form " & strForm & " was added successfully."

        iRetval = kErrorSuccess

    else

        wscript.echo "Unable to add form " & strForm & ", error code: 0x" & _
                      Hex(Err.Number) & ". " & Err.Description

        iRetval = kErrorFailure

    end if

    AddForm = iRetval

end function

'
' Delete a Form
'
function DeleteForm(strServer, strForm)

    on error resume next

    DebugPrint kDebugTrace, "In DeleteForm"
    DebugPrint kDebugTrace, "Server Name " & strServer
    DebugPrint kDebugTrace, "Form Name   " & strForm

    dim oMaster
    dim oForm
    dim iRetval

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    set oForm = CreateObject("Form.Form.1")

    oForm.Name = strForm

    oForm.ServerName = strServer

    oMaster.FormDel oForm

    if Err.Number = kErrorSuccess then

        wscript.echo "Form " & strForm & " deleted successfully."

        iRetval = kErrorSuccess

    else

        wscript.echo "Unable to delete form " & strForm & ", error code: 0x" _
                     & Hex(Err.Number) & ". " & Err.Description

        iRetval = kErrorFailure

    end if

    DeleteForm = iRetval

end function

'
' List all Forms
'
function ListForms(strServer, iUnits, bVerbose)

    on error resume next

    DebugPrint kDebugTrace, "In ListForms"
    DebugPrint kDebugTrace, "Server Name " & strServer

    dim iResult
    dim oMaster
    dim oForm
    dim iRetval

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    for each oForm in oMaster.Forms(strServer)

        iRetval = DisplayFormInformation(oForm, iUnits, bVerbose)

    next

    iResult = kErrorSuccess

    ListForms = iResult

end function

'
' Disply a form's information.
'
function DisplayFormInformation(oForm, iUnits, bVerbose)

    on error resume next

    dim iHeight
    dim iWidth
    dim iTop
    dim iLeft
    dim iBottom
    dim iRight
    dim strName
    dim strFlags
    dim strPad
    dim iRetval

    '
    ' Get the form name
    '
    strName = oForm.Name

    '
    ' Convert the form flags to human readable
    '
    strFlags = ConvertFormFlagsToString(oForm.Flags)

    '
    ' Get the form's height and width
    '
    oForm.GetSize iHeight, iWidth

    '
    ' Get the form's imageable area expressed as coordinate pairs
    '
    oForm.GetImageableArea iLeft, iTop, iRight, iBottom

    '
    ' Convert the form size to the specified units.
    '
    iHeight = Convert(iUnits, false, iHeight)
    iWidth  = Convert(iUnits, false, iWidth)
    iTop    = Convert(iUnits, false, iTop)
    iLeft   = Convert(iUnits, false, iLeft)
    iBottom = Convert(iUnits, false, iBottom)
    iRight  = Convert(iUnits, false, iRight)

    strPad = String(31 - Len(strName), " ")

    if bVerbose = true then

        wscript.echo "Name: " & strName & strPad & "Type: " & strFlags & " Size: " & iWidth _
                     & " x " & iHeight & " Imageable area: (" & iTop & "," & iLeft & ")(" & _
                     iBottom & "," & iRight & ")"

    else

        wscript.echo "Name: " & strName & strPad & "Size: " & iWidth & " x " & iHeight

    end if

    '
    ' Always return success.
    '
    DisplayFormInformation = kErrorSuccess

end function

'
' Create a form object that can be used to call FormAdd
'
function CreateForm(strForm, iUnits, iHeight, iWidth, iTop, iLeft, iBottom, iRight)

    on error resume next

    DebugPrint kDebugTrace, "In CreateForm"
    DebugPrint kDebugTrace, "iUnits      " & iUnits
    DebugPrint kDebugTrace, "iHeight     " & iHeight
    DebugPrint kDebugTrace, "iWidth      " & iWidth
    DebugPrint kDebugTrace, "iTop        " & iTop
    DebugPrint kDebugTrace, "iLeft       " & iLeft
    DebugPrint kDebugTrace, "iBottom     " & iBottom
    DebugPrint kDebugTrace, "iRight      " & iRight

    dim oForm
    dim temp1
    dim temp2

    '
    ' Validate the coordinates
    '
    if iLeft > iWidth or iRight > iWidth or iLeft > iRight or _
       iTop > iHeight or iBottom > iHeight or iTop > iBottom then

        wscript.echo "Error: Incorrect coordinates. Cannot create form"

        wscript.quit

    else

    set oForm = CreateObject("Form.Form.1")

    oForm.Name = strForm

    oForm.SetSize Convert(iUnits, true, iHeight), Convert(iUnits, true, iWidth)

    oForm.SetImageableArea Convert(iUnits, true, iTop),    Convert(iUnits, true, iLeft), _
                           Convert(iUnits, true, iBottom), Convert(iUnits, true, iRight)

    set CreateForm = oForm

    end if

end function

'
' Convert the form flag to a human readable string.
'
function ConvertFormFlagsToString(Flags)

    on error resume next

    select case Flags

    case 0
        ConvertFormFlagsToString = "User"

    case 1
        ConvertFormFlagsToString = "Built-in"

    case 2
        ConvertFormFlagsToString = "Printer"

    case else
        ConvertFormFlagsToString = "Unknown"

    end select

end function

'
' Convert the value to the specified units.
'
function Convert(iUnits, bSource, dValue)

    on error resume next

    '
    ' If bSource is true the value is from the command
    ' line and expressed either in inches or centimeters.
    '
    if bSource = true then

        select case iUnits

        case kInches
            '
            ' Convert from inches to thousands of millimeters
            '
            Convert = CLng(dValue * 100 * 254)

        case kCentimeters
            '
            ' Convert from centimeters to thousands of millimeters
            '
            Convert = CLng(dValue * 100 * 100)

        end select

    '
    ' If bSource is false the value is from the spooler and it
    ' is expressed in thousandths of millimeters.
    '
    else

        select case iUnits

        case kInches
            '
            ' Convert from thousands of millimeters to inches
            '
            Convert = Round(CDbl(dValue / 25400), 2)

        case kCentimeters
            '
            ' Convert from thousands of millimeters to centimeters
            '
            Convert = Round(CDbl(dValue / 10000), 2)

        end select

    end if

end function

'
' Validate the unit string command argument
'
function ValidateUnits(strUnits, iUnits)

    on error resume next

    DebugPrint kDebugTrace, "In ValidateUnits"

    ValidateUnits = true

    if strUnits <> "" then

        if lcase(strUnits) = "inches" then

            iUnits = kInches

        elseif lcase(strUnits) = "centimeters" then

            iUnits = kCentimeters

        else

            ValidateUnits = false

        end if

    end if

end function

'
' Debug display helper function
'
sub DebugPrint(uFlags, strString)

    if gDebugFlag = true then

        if uFlags = kDebugTrace then

            wscript.echo "Debug: " & strString

        end if

        if uFlags = kDebugError then

            if Err <> 0 then

                wscript.echo "Debug: " & strString & " Failed with " & Hex(Err)

            end if

        end if

    end if

end sub

'
' Parse the command line into it's components
'
function ParseCommandLine(iAction, strServer, strForm, iUnits, iHeight, _
                          iWidth, iTop, iLeft, iBottom, iRight, bVerbose)

    DebugPrint kDebugTrace, "In the ParseCommandLine"

    dim oArgs
    dim strUnits
    dim i

    iAction = kActionUnknown

    set oArgs = wscript.Arguments

    while i < oArgs.Count

        select case oArgs(i)

            case "-a"
                iAction = kActionAdd

            case "-d"
                iAction = kActionDelete

            case "-l"
                iAction = kActionList

            case "-c"
                i = i + 1
                strServer = oArgs(i)

            case "-v"
                bVerbose = true

            case "-n"
                i = i + 1
                strForm = oArgs(i)

            case "-h"
                i = i + 1
                iHeight = CDbl(oArgs(i))

            case "-w"
                i = i + 1
                iWidth = CDbl(oArgs(i))

            case "-t"
                i = i + 1
                iTop = CDbl(oArgs(i))

            case "-e"
                i = i + 1
                iLeft = CDbl(oArgs(i))

            case "-b"
                i = i + 1
                iBottom = CDbl(oArgs(i))

            case "-r"
                i = i + 1
                iRight = CDbl(oArgs(i))

            case "-u"
                i = i + 1
                strUnits = oArgs(i)

            case "-?"
                Usage(true)

            case else
                Usage(true)

        end select

        i = i + 1

    wend

    '
    ' Check if the units specified are valid.
    '
    if ValidateUnits(strUnits, iUnits) = false then

        wscript.echo "Unsupported units"

        ParseCommandLine = kErrorFailure

    else

        ParseCommandLine = kErrorSuccess

    end if

end  function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo "Usage: forms [-vadl?] [-c server] [-n form-name] [-u inches|centimeters]"
    wscript.echo "             [-h height] [-w width] [-t top] [-e left] [-b bottom] [-r right]"
    wscript.echo "Arguments:"
    wscript.echo "-a     - add a form"
    wscript.echo "-d     - delete a form"
    wscript.echo "-l     - list all forms"
    wscript.echo "-v     - used with list to display entire form details"
    wscript.echo "-h     - specifies the height of the form"
    wscript.echo "-w     - specifies the width of the form"
    wscript.echo "-t     - specifies top coordinate of the imageable area"
    wscript.echo "-e     - specifies left coordinate of the imageable area"
    wscript.echo "-b     - specifies bottom coordinate of the imageable area"
    wscript.echo "-r     - specifies right coordinate of the imageable area"
    wscript.echo "-u     - specifies the units of the size arguments"
    wscript.echo "-v     - verbose list forms full information"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "forms -l"
    wscript.echo "forms -l -c \\server -u inches"
    wscript.echo "forms -a -c \\server -n ""Large Paper"""
    wscript.echo "forms -d -c \\server -n ""Tabloid"""
    wscript.echo "forms -a -n NewForm -u inches -h 11 -w 8.5 -t 0 -e 0 -b 7 -r 8"

    if bExit then
        wscript.quit(1)
    end if

end sub

'
' Determines which program is used to run this script.
' Returns true if the script host is cscript.exe
'
function IsHostCscript()

    on error resume next

    dim strFullName
    dim strCommand
    dim i, j
    dim bReturn

    bReturn = false

    strFullName = WScript.FullName

    i = InStr(1, strFullName, ".exe", 1)

    if i <> 0 then

        j = InStrRev(strFullName, "\", i, 1)

        if j <> 0 then

            strCommand = Mid(strFullName, j+1, i-j-1)

            if LCase(strCommand) = "cscript" then

                bReturn = true

            end if

        end if

    end if

    if Err <> 0 then

        call wscript.echo("Error 0x" & hex(Err.Number) & " occurred. " & Err.Description _
                          & ". " & vbCRLF & "The scripting host could not be determined.")

    end if

    IsHostCscript = bReturn

end function

