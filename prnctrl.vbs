'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' Abstract:
'
' prnctrl.vbs - printer control script for Windows .NET Server 2003
'
' Usage:
' prnctrl [-prxt?] [-b printer]
'
' Examples:
' prnctrl.vbs -p -b \\server\printer
' prnctrl.vbs -t -b printer
'----------------------------------------------------------------------

option explicit

'
' Debugging trace flags, to enable debug output trace message
' change gDebugFlag to true.
'
const kDebugTrace = 1
const kDebugError = 2
dim   gDebugFlag

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
const kActionPause      = 1
const kActionResume     = 2
const kActionPurge      = 3
const kActionTestPage   = 4

const kErrorSuccess     = 0
const KErrorFailure     = 1

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strPrinter

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(kMessage1 & vbCRLF & kMessage2 & vbCRLF & _
                          kMessage3 & vbCRLF & kMessage4 & vbCRLF & _
                          kMessage5 & vbCRLF & kMessage6 & vbCRLF)

        wscript.quit

    end if

    iRetval = ParseCommandLine(iAction, strPrinter)

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionPause
                 iRetval = PausePrinter(strPrinter)

            case kActionResume
                 iRetval = ResumePrinter(strPrinter)

            case kActionPurge
                 iRetval = PurgePrinter(strPrinter)

            case kActionTestPage
                 iRetval = PrintTestPage(strPrinter)

            case else
                 Usage(true)
                 exit sub

        end select

    end if

end sub

'
' Pause printer
'
function PausePrinter(strPrinter)

    on error resume next

    DebugPrint kDebugTrace, "In PausePrinter"

    dim oMaster
    dim iResult

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    oMaster.PrinterPause "", strPrinter

    if Err.Number = kErrorSuccess then

        wscript.echo "Printer """ & strPrinter & """ was paused"

        iResult = kErrorSuccess

    else

        wscript.echo "Unable to pause printer """ & strPrinter & """, error: 0x" _
                     & Hex(Err.Number) & ". " & Err.Description

        iResult = kErrorFailure

    end if

    PausePrinter = iResult

end function

'
' Resume printer
'
function ResumePrinter(strPrinter)

    on error resume next

    DebugPrint kDebugTrace, "In ResumePrinter"

    dim oMaster
    dim iResult

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    oMaster.PrinterResume "", strPrinter

    if Err.Number = kErrorSuccess then

        wscript.echo "Printer """ & strPrinter & """ was resumed"

        iResult = kErrorSuccess

    else

        wscript.echo "Unable to resume printer """ & strPrinter & """, error: 0x" _
                     & Hex(Err.Number) & ". " & Err.Description

        iResult = kErrorFailure

    end if

    ResumePrinter = iResult

end function

'
' Purge printer
'
function PurgePrinter(strPrinter)

    on error resume next

    DebugPrint kDebugTrace, "In PurgePrinter"

    dim oMaster
    dim iResult

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    oMaster.PrinterPurge "", strPrinter

    if Err.Number = kErrorSuccess then

        wscript.echo "Success: printer """ & strPrinter & """ was purged"

        iResult = kErrorSuccess

    else

        wscript.echo "Unable to purge printer """ & strPrinter & """, error: 0x" _
                     & Hex(Err.Number) & ". " & Err.Description

        iResult = kErrorFailure

    end if

    PurgePrinter = iResult

end function

'
' Print test page
'
function PrintTestPage(strPrinter)

    on error resume next

    DebugPrint kDebugTrace, "In PrintTestPage"

    dim oMaster
    dim iResult

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    oMaster.PrintTestPage "", strPrinter

    if Err = 0 then

        wscript.echo "Test page sent to printer """ & strPrinter & """"

        iResult = kErrorSuccess

    else

        wscript.echo "Unable to send test page to printer """ & strPrinter & _
                     """, error: 0x" & Hex(Err.Number) & ". " & Err.Description

        iResult = kErrorFailure

    end if

    PrintTestPage = iResult

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
function ParseCommandLine(iAction, strPrinter)

    on error resume next

    DebugPrint kDebugTrace, "In the ParseCommandLine"

    dim oArgs
    dim iIndex

    iAction = kActionUnknown
    iIndex = 0

    set oArgs = wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-p"
                iAction = kActionPause

            case "-r"
                iAction = kActionResume

            case "-x"
                iAction = kActionPurge

            case "-t"
                iAction = kActionTestPage

            case "-b"
                iIndex = iIndex + 1
                strPrinter = oArgs(iIndex)

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

        wscript.echo "Unable to parse command line, error 0x" & Hex(Err.Number) _
                     & " " & Err.Description

        ParseCommandLine = kErrorFailure

    end if

    if strPrinter = "" then

        wscript.echo "Please specify a printer name"

        Usage true

    end if

end function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo "Usage: prnctrl [-prxt?] [-b printer]"
    wscript.echo ""
    wscript.echo "Arguments:"
    wscript.echo "-p     - pause the printer"
    wscript.echo "-r     - resume the printer"
    wscript.echo "-x     - purge the printer"
    wscript.echo "-t     - print test page"
    wscript.echo "-?     - display command usage"
    wscript.echo "-b     - printer name"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "prnctrl.vbs -p -b \\server\printer"
    wscript.echo "prnctrl.vbs -t -b printer"

    if bExit <> 0 then

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

