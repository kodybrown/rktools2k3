'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' clean.vbs - delete all printing components from the specified
'             machine, as if the machine were clean installed.
'
' Usage:
' clean [-afpdb?] [-c \\server]
'
' Examples:
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
const kActionAll        = 1
const kActionForms      = 2
const kActionPorts      = 3
const kActionDrivers    = 4
const kActionPrinters   = 5

'
' General error values
'
const kErrorSuccess     = 0
const kErrorFailure     = 1

'
' Constant printer attribute local
'
const kLocalPrinterFlag = 64

'
' Constant local port
'
const kLocal = 3

'
'  Search patterns used in the regular expressions
'  to skip COMx: and LPTx: ports
'
const kComPortPattern   = "^COM\d+:$"
const kLptPortPattern   = "^LPT\d+:$"

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strServer
    dim bVerbose

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(kMessage1 & vbCRLF & kMessage2 & vbCRLF & _
                          kMessage3 & vbCRLF & kMessage4 & vbCRLF & _
                          kMessage5 & vbCRLF & kMessage6 & vbCRLF)

        wscript.quit

    end if

    iRetval = ParseCommandLine(iAction, strServer, bVerbose)

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionAll
                iRetval = CleanAll(strServer, bVerbose)

            case kActionForms
                iRetval = CleanForms(strServer, bVerbose)

            case kActionPorts
                iRetval = CleanPorts(strServer, bVerbose)

            case kActionDrivers
                iRetval = CleanDrivers(strServer, bVerbose)

            case kActionPrinters
                iRetval = CleanPrinters(strServer, bVerbose)

            case else
                Usage(true)

        end select

    end if

end sub

'
' General function for deleting printing objects
'
function CleanAll(strServer, bVerbose)

    on error resume next

    DebugPrint kDebugTrace, "In the CleanAll function"

    dim iRetval

    iRetval = CleanPrinters(strServer, bVerbose)
    iRetval = CleanPorts(strServer, bVerbose)
    iRetval = CleanDrivers(strServer, bVerbose)
    iRetval = CleanForms(strServer, bVerbose)

    CleanAll = kErrorSucces

end function

'
' Clean all forms
'
function CleanForms(strServer, bVerbose)

    on error resume next

    DebugPrint kDebugTrace, "In the CleanForms function"

    dim oMaster
    dim oForm
    dim iCount
    dim iTotal

    '
    ' Number of forms found
    '
    iTotal = 0

    '
    ' Number of forms deleted
    '
    iCount = 0

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    for each oForm in oMaster.Forms(strServer)

        if Err.Number = kErrorSuccess then

            '
            ' Tally the total number of forms.
            '
            iTotal = iTotal + 1

            '
            ' Delete the form
            '
            oMaster.FormDel oForm

            if Err.Number = kErrorSuccess then

                '
                ' Count the number of forms that were deleted
                '
                iCount = iCount + 1

            else

                if bVerbose then

                    wscript.echo "Error deleting form " & oForm.Name & ", error: 0x" _
                                 & Hex(Err.Number) & ". " & Err.Description

                end if

            end if

            '
            ' Clear the previous error code after each iteration
            '
            Err.Clear

        else

            wscript.echo "Unable to enumerate forms, error: 0x" & _
                         Hex(Err.Number) & ". " & Err.Description

            CleanForms = kErrorFailure

            exit function

        end if

    next

    wscript.echo "Number of forms found " & iTotal & ". Forms deleted " & iCount

    CleanForms = kErrorSuccess

end function

'
' Clean all ports
'
function CleanPorts(strServer, bVerbose)

    on error resume next

    DebugPrint kDebugTrace, "In the CleanPorts function"

    dim oMaster
    dim oPort
    dim iCount
    dim iTotal

    '
    ' Number of ports found
    '
    iTotal = 0

    '
    ' Number of ports deleted
    '
    iCount = 0

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    for each oPort in oMaster.Ports(strServer)

        if Err.Number = kErrorSuccess then

            '
            ' Tally the total number of ports.
            '
            iTotal = iTotal + 1

            '
            ' Do not delete LPTx:, COMx:, FILE: or NUL ports
            '
            if  oPort.PortType         =  kLocal                       and _
              ( bFindPattern(kComPortPattern, oPort.PortName) = true    or _
                bFindPattern(kLptPortPattern, oPort.PortName) = true    or _
                oPort.PortName = "FILE:"                                or _
                oPort.PortName = "NUL") then

                if bVerbose then

                    wscript.echo "Skiping port " & oPort.PortName

                end if

            else

                '
                ' Delete the port
                '
                oMaster.PortDel oPort

                if Err.Number = kErrorSuccess then

                    '
                    ' Count the number of ports that were deleted.
                    '
                    iCount = iCount + 1

                    if bVerbose then

                        wscript.echo "Deleted port " & oPort.PortName

                    end if

                else

                    if bVerbose then

                        wscript.echo "Error deleting port " & oPort.PortName & ", error: 0x" _
                                     & Hex(Err.Number) & ". " & Err.Description

                    end if

                end if

            end if

            Err.Clear

        else

            wscript.echo "Unable to enumerate ports, error: 0x" & _
                         Hex(Err.Number) & ". " & Err.Description

            CleanPorts = kErrorFailure

            exit function

        end if

    next

    wscript.echo "Number of ports found " & iTotal & ". Ports deleted " & iCount

    CleanPorts = kErrorSuccess

end function

'
' Clean all drivers
'
function CleanDrivers(strServer, bVerbose)

    on error resume next

    DebugPrint kDebugTrace, "In the CleanDrivers function"

    dim oMaster
    dim oDriver
    dim iCount
    dim iTotal

    '
    ' Number of drivers found
    '
    iTotal = 0

    '
    ' Number of drivers deleted
    '
    iCount = 0

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    for each oDriver in oMaster.Drivers(strServer)

        if Err.Number = kErrorSuccess then

            '
            ' Tally the total number of drivers.
            '
            iTotal = iTotal + 1

            '
            ' Delete the driver
            '
            oMaster.DriverDel oDriver

            if Err.Number = kErrorSuccess then

                '
                ' Count the number of drivers that were deleted.
                '
                iCount = iCount + 1

                if bVerbose then

                    wscript.echo "Deleted driver " & oDriver.ModelName

                end if

            else

                if bVerbose then

                    wscript.echo "Error deleting driver " & oDriver.ModelName & _
                                 ", error: 0x" & Hex(Err.Number) & ". " & Err.Description

                end if

            end if

            Err.Clear

        else

            wscript.echo "Unable to enumerate drivers, error: 0x" & _
                         Hex(Err.Number) & ". " & Err.Description

            CleanDrivers = kErrorFailure

            exit function

        end if

    next

    wscript.echo "Number of drivers found " & iTotal & ". Drivers deleted " & iCount

    CleanDrivers = kErrorSuccess

end function

'
' Clean all printers
'
function CleanPrinters(strServer, bVerbose)

    on error resume next

    DebugPrint kDebugTrace, "In the CleanPrinters function"
    dim oMaster
    dim oPrinter
    dim iCount
    dim iTotal

    '
    ' Number of printers found
    '
    iTotal = 0

    '
    ' Number of printers deleted
    '
    iCount = 0

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    for each oPrinter in oMaster.Printers(strServer)

        if Err.Number = kErrorSuccess then

            '
            ' Tally the total number of printers.
            '
            iTotal = iTotal + 1

            '
            ' When we enumerate printers on a remote machine, they will
            ' all have the attribute set to Network, not to local. When
            ' we enumerate printers on the local machine, connections have
            ' the attribute network. This is the reason why we need to know
            ' whether we enumerate on a  remote machine or on the local one
            ' Remark: if strServer is not empty, no connections will be
            ' enumerated
            '

            if strServer = "" then

                if (oPrinter.Attributes and kLocalPrinterFlag) = kLocalPrinterFlag then

                    DebugPrint kDebugTrace, "Deleting local printer " & oPrinter.PrinterName

                    '
                    ' Purge the printer
                    '
                    oMaster.PrinterPurge strServer, oPrinter.PrinterName

                    '
                    ' Sometimes we can delete a printer, but we cannot purge it
                    ' We need to clear the error in this case
                    '
                    Err.Clear

                    oMaster.PrinterDel oPrinter

                else

                    DebugPrint kDebugTrace, "Deleting printer connection" & oPrinter.PrinterName

                    oMaster.PrinterConnectionDel oPrinter.PrinterName

                end if

            else

                '
                ' Purge the printer, it is a local printer on a remote machine
                '
                oMaster.PrinterPurge strServer, oPrinter.PrinterName

                '
                ' In some cases, we can delete a printer, but not purge it
                ' We need to clear the error
                '
                Err.Clear

                oPrinter.ServerName = strServer

                oMaster.PrinterDel oPrinter

            end if

            if Err.Number = kErrorSuccess then

                '
                ' Count the number of printers that were deleted.
                '
                iCount = iCount + 1

                if bVerbose then

                    wscript.echo "Deleted printer " & oPrinter.PrinterName

                end if

            else

                if bVerbose then

                    wscript.echo "Error deleting printer " & oPrinter.PrinterName & _
                                 ", error: 0x" & Hex(Err.Number) & ". " & Err.Description

                end if

            end if

            Err.Clear

        else

            wscript.echo "Unable to enumerate printers, error: 0x" & _
                         Hex(Err.Number) & ". " & Err.Description

            CleanPrinters = kErrorFailure

            exit function

        end if

    next

    wscript.echo "Number of printers found " & iTotal & ". Printers deleted " & iCount

    CleanPrinters = kErrorSuccess

end function

'
' Resolve the regular expression
'
function bFindPattern(strPattern, strString)

    dim RegEx                            ' Create variable.
    set RegEx = New RegExp               ' Create regular expression.
    RegEx.Pattern = strPattern           ' Set pattern.
    RegEx.IgnoreCase = true              ' Set case insensitivity.
    bFindPattern = RegEx.Test(strString) ' Test if pattern is found.

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
function ParseCommandLine(iAction, strServer, bVerbose)

    on error resume next

    DebugPrint kDebugTrace, "In the ParseCommandLine"

    dim oArgs
    dim i

    iAction = kActionUnknown

    set oArgs = wscript.Arguments

    while i < oArgs.Count

        select case oArgs(i)

            case "-a"
                iAction = kActionAll

            case "-f"
                iAction = kActionForms

            case "-p"
                iAction = kActionPorts

            case "-d"
                iAction = kActionDrivers

            case "-b"
                iAction = kActionPrinters

            case "-c"
                i = i + 1
                strServer = oArgs(i)

            case "-v"
                bVerbose = true

            case "-?"
                Usage(true)

            case else
                Usage(true)

        end select

        i = i + 1

    wend

    if Err.Number = kErrorSuccess then

        ParseCommandLine = kErrorSuccess

    else

        wscript.echo "Unable to parse command line, error: 0x" & _
                     Hex(Err.Number) & ". " & Err.Description

        ParseCommandLine = kErrorFailure

    end if

end  function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo "Usage: clean [-afpdbv?] [-c server]"
    wscript.echo "Arguments:"
    wscript.echo "-a     - clean all, printers, ports, drivers, forms"
    wscript.echo "-f     - clean forms"
    wscript.echo "-p     - clean ports"
    wscript.echo "-b     - clean printers"
    wscript.echo "-d     - clean drivers"
    wscript.echo "-v     - verbose display progress and error information"
    wscript.echo "-c     - server name"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "clean -a"
    wscript.echo "clean -a -c \\server"
    wscript.echo "clean -v -f -c \\server"

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


