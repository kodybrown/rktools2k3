'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' Abstract:
'
' defprn.vbs - default printer script for Windows .NET Server 2003
'
' Usage:
' defprn [-sg?] [-n printer]
'
' Example:
' defprn -s -n printer
' defprn -g
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
const kActionSet        = 1
const kActionGet        = 2

const kErrorSuccess     = 0
const KErrorFailure     = 1

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strPrinterName

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(kMessage1 & vbCRLF & kMessage2 & vbCRLF & _
                          kMessage3 & vbCRLF & kMessage4 & vbCRLF & _
                          kMessage5 & vbCRLF & kMessage6 & vbCRLF)

        wscript.quit

    end if

    iRetval = ParseCommandLine(iAction, strPrinterName)

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionSet
                iRetval = SetDefaultPrinter(strPrinterName)

            case kActionGet
                iRetval = GetDefaultPrinter()

            case else
                Usage(True)
                exit sub

        end select

    end if

end sub

'
' Get the name of the default printer if one exists.
'
function GetDefaultPrinter()

    on error resume next

    DebugPrint kDebugTrace, "In the GetDefaultPrinter"

    dim strDefaultPrinter
    dim oPrint
    dim iRetval

    set oPrint = CreateObject("PrintMaster.PrintMaster.1")

    strDefaultPrinter = oPrint.DefaultPrinter

    if Err.Number = kErrorSuccess then

        wscript.echo "The default printer is: """ & strDefaultPrinter & """ "

        iRetval = kErrorSuccess

    else

        wscript.echo "Unable to get the default printer, error: 0x" & Hex(Err.Number) & ". " & Err.Description

        iRetval = kErrorFailure

    end if

    GetDefaultPrinter = iRetval

end function

'
' Set the specified printer name as the default printer.
'
function SetDefaultPrinter(strName)

    on error resume next

    DebugPrint kDebugTrace, "In the SetDefaultPrinter " & strName

    dim oPrint
    dim iRetval

    set oPrint = CreateObject("PrintMaster.PrintMaster.1")

    oPrint.DefaultPrinter = strName

    if Err.Number = kErrorSuccess then

        wscript.echo "The default printer was set to: """ & strName & """ "

        iRetval = kErrorSuccess

    else

        wscript.echo "Unable to set the default printer, error: 0x" & Hex(Err.Number) & ". " & Err.Description

        iRetval = kErrorFailure

    end if

    SetDefaultPrinter = iRetval

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
function ParseCommandLine(iAction, strPrinterName)

    on error resume next

    DebugPrint kDebugTrace, "In the ParseCommandLine"

    dim oArgs
    dim iIndex

    iAction = kActionUnknown
    iIndex = 0

    set oArgs = wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-s"
                iAction = kActionSet

            case "-g"
                iAction = kActionGet

            case "-n"
                iIndex = iIndex + 1
                strPrinterName = oArgs(iIndex)

            case "-?"
                Usage(true)
                exit function

            case else
                Usage(true)
                exit function

        end select

        iIndex = iIndex + 1

    wend

    if Err.Number = kErrorSuccess then

        ParseCommandLine = kErrorSuccess

    else

        wscript.echo "Unable to parse command line, error 0x" & Hex(Err.Number) & ". " & Err.Description

        ParseCommandLine = kErrorFailure

    end if

end  function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo "Usage: defprn [-sg?] [-n printer]"
    wscript.echo "Arguments:"
    wscript.echo "-s     - set the specified printer as the default printer"
    wscript.echo "-g     - get the current default printer"
    wscript.echo "-n     - printer name"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "defprn -s -n printer"
    wscript.echo "defprn -g"

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

