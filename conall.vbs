'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' Abstract:
'
' conall.vbs - connects to all printers on a print server.
'
' Usage:
' conall [-a?] [-c server]
'
' Examples:
' conall -c \\server
' conall -a -c \\server
'
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
const kActionConnect    = 1

const kErrorSuccess     = 0
const KErrorFailure     = 1

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strServer

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(kMessage1 & vbCRLF & kMessage2 & vbCRLF & _
                          kMessage3 & vbCRLF & kMessage4 & vbCRLF & _
                          kMessage5 & vbCRLF & kMessage6 & vbCRLF)

        wscript.quit

    end if

    iRetval = ParseCommandLine(iAction, strServer)

    if iRetval = kErrorSuccess then

        select case iAction

        case kActionConnect

            if strServer = "" then

                wscript.echo "Please specify the server you want to connect to"

                iRetVal = kErrorFailure

            else

                iRetval = AddAllPrinters(strServer)

            end if

        case else

             Usage(true)

        end select

    end if

end sub

'
' Add a printer connection
'
function AddPrinterConnection(strPrinter)

    on error resume next

    DebugPrint kDebugTrace, "In AddPrinterConnection"

    dim oMaster
    dim iRetval

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    oMaster.PrinterConnectionAdd strPrinter

    if Err.Number = kErrorSuccess then

        wscript.echo "Connected to printer """ & strPrinter & """ "

        iRetval = kErrorSuccess

    else

        wscript.echo "Unable to add printer connection """ & strPrinter & _
                     """, error: 0x" & Hex(Err.Number) & ". " & Err.Description

        Err.Clear

        iRetval = kErrorFailure

    end if

    AddPrinterConnection = iRetval

end function

'
' List the printers and make the connections
'
function AddAllPrinters(strServer)

    on error resume next

    DebugPrint kDebugTrace, "In AddAllPrinters"

    dim oMaster
    dim oPrinter
    dim oError
    dim iCount

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    iCount = 0

    for each oPrinter in oMaster.Printers(strServer)

        if Err.Number = kErrorSuccess then

            if AddPrinterConnection(oPrinter.PrinterName) = kErrorSuccess then

                '
                ' Count the number of printer connections that were made
                '
                iCount = iCount + 1

            end if

        else

            wscript.echo "Unable to enumerate printers on server, error: 0x" _
                         & Hex(Err.Number) & ". " & Err.Description

            AddAllPrinters = kErrorFailure


            exit function

        end if

    next

    if iCount = 0 then

        wscript.echo "There were no printers to connect to"

    else

        wscript.echo "Number of connections made " & iCount

    end if

    AddAllPrinters = kErrorSuccess

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
function ParseCommandLine(iAction, strServer)

    on error resume next

    DebugPrint kDebugTrace, "In the ParseCommandLine"

    dim oArgs
    dim strArg
    dim iIndex

    set oArgs = wscript.Arguments

    iAction = kActionUnknown
    iIndex = 0

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-a"
                iAction = kActionConnect

            case "-c"
                iIndex = iIndex + 1
                strServer = oArgs(iIndex)

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

        wscript.echo "Unable to parse command line, error 0x" & _
                     Hex(Err.Number) & ". " & Err.Description

        ParseCommandLine = KErrorFailure

    end if

end  function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo "Usage: conall [-a?][-c server]"
    wscript.echo "Arguments:"
    wscript.echo "-a     - connect to all printers"
    wscript.echo "-c     - server name"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "conall -a -c \\server"

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
