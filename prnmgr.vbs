'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' Abstract:
'
' prnmgr.vbs - printer script for Windows .NET Server 2003
'
' Usage:
' prnmgr [-adl?][c] [-c server][-b printer][-m driver]
'                   [-l driver path][-r port][-f file]
' Examples:
' prnmgr -l -c \\server
' prnmgr.vbs -a -b "printer" -m "driver" -r "lpt1:"
' prnmgr.vbs -a -b "printer" -m "driver" -r "lpt1:"
' prnmgr.vbs -a -b "printer" -c \\server
' prnmgr.vbs -ac -b "\\server\printer"
' prnmgr.vbs -dc -b "\\server\printer"
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
const kActionUnknown     = 0
const kActionAdd         = 1
const kActionAddConn     = 2
const kActionDel         = 3
const kActionDelConn     = 4
const kActionList        = 5
const kActionDelAll      = 6
const kActionDelConnAll  = 7

const kErrorSuccess      = 0
const KErrorFailure      = 1

const kPrinterNetwork    = 16
const kLocalPrinterFlag  = 64

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strServer
    dim strPrinter
    dim strDriverPath
    dim strDriver
    dim strPort
    dim strInfFile

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(kMessage1 & vbCRLF & kMessage2 & vbCRLF & _
                          kMessage3 & vbCRLF & kMessage4 & vbCRLF & _
                          kMessage5 & vbCRLF & kMessage6 & vbCRLF)

        wscript.quit

    end if

    iRetval = ParseCommandLine(iAction, strServer, strPrinter, strDriver, strDriverPath, strPort, strInfFile)

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionAdd
                 iRetval = AddPrinter(strServer, strPrinter, strDriver, strDriverPath, strPort, strInfFile)

            case kActionAddConn
                 iRetval = AddPrinterConnection(strPrinter)

            case kActionDel
                 iRetval = DelPrinter(strServer, strPrinter)

            case kActionDelConn
                 iRetval = DelPrinterConnection(strPrinter)

            case kActionList
                 iRetval = ListPrinters(strServer)

            case kActionDelAll
                 iRetval = DelPrinterAll(strServer)

            case kActionDelConnAll
                 iRetval = DelPrinterConnectionAll()

            case else
                 Usage(true)
                 exit sub

        end select

    end if

end sub

'
' Add a printer
'
function AddPrinter(strServer, strPrinter, strDriver, strDriverPath, strPort, strInfFile)

    on error resume next

    DebugPrint kDebugTrace, "In AddPrinter"
    DebugPrint kDebugTrace, "Server      "  & strServer
    DebugPrint kDebugTrace, "Printer     "  & strPrinter
    DebugPrint kDebugTrace, "Driver      "  & strDriver
    DebugPrint kDebugTrace, "DrivatePath "  & strDriverPath
    DebugPrint kDebugTrace, "Port        "  & strPort
    DebugPrint kDebugTrace, "InfFile     "  & strInfFile

    dim oMaster
    dim oPrinter
    dim iRetval

    set oMaster  = CreateObject("PrintMaster.PrintMaster.1")
    set oPrinter = CreateObject("Printer.Printer.1")

    oPrinter.ServerName  = strServer

    oPrinter.DriverPath  = strDriverPath

    oPrinter.DriverName  = strDriver

    oPrinter.PortName    = strPort

    oPrinter.PrinterName = strPrinter

    oPrinter.InfFile     = strInfFile

    oMaster.PrinterAdd oPrinter

    if Err.Number = kErrorSuccess then

        wscript.echo "Printer """ & strPrinter & """ added"

        iRetval = kErrorSuccess

    else

        wscript.echo "Unable to add printer """ & strPrinter & """, error: 0x" _
                     & Hex(Err.Number) & ". " & Err.Description

       iRetval = kErrorFailure

    end if

    AddPrinter = iRetval

end function

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

        wscript.echo "Printer connection """ & strPrinter & """ added"

        iRetval = kErrorSuccess

    else

        wscript.echo "Unable to add printer connection, error: 0x" & Hex(Err.Number) _
                     & ". " & Err.Description

        iRetval = kErrorFailure

    end if

    AddPrinterConnection = iRetval

end function

'
' Delete a printer connection
'
function DelPrinterConnection(strPrinter)

    on error resume next

    DebugPrint kDebugTrace, "In DelPrinterConnection"

    dim oMaster
    dim iRetval

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    oMaster.PrinterConnectionDel strPrinter

    if Err.Number = kErrorSuccess then

        wscript.echo "Deleted printer connection " & strPrinter

        iRetval = kErrorSuccess

    else

        wscript.echo "Unable to delete printer connection " & strPrinter & ", error: 0x"_
                     & Hex(Err.Number) & ". " & Err.Description

        iRetval = kErrorFailure

    end if

    DelPrinterConnection = iRetval

end function

'
' Delete a printer
'
function DelPrinter(strServer, strPrinter)

    on error resume next

    DebugPrint kDebugTrace, "In DelPrinter"

    dim oMaster
    dim oPrinter
    dim iRetval

    set oMaster  = CreateObject("PrintMaster.PrintMaster.1")
    set oPrinter = CreateObject("Printer.Printer.1")

    oPrinter.ServerName  = strServer
    oPrinter.PrinterName = strPrinter

    oMaster.PrinterDel oPrinter

    if Err.Number = kErrorSuccess then

        wscript.echo "Printer """ & strPrinter & """ deleted"

        iRetval = kErrorSuccess

    else

        wscript.echo "Unable to delete printer """ & strPrinter &  """, error: 0x" _
                     & Hex(Err.Number) & ". " & Err.Description

        iRetval = kErrorFailure

    end if

    DelPrinter = iRetval

end function

'
' List the printers
'
function ListPrinters(strServer)

    on error resume next

    DebugPrint kDebugTrace, "In ListPrinter"

    dim oMaster
    dim oPrinter
    dim iRetval

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    for each oPrinter in oMaster.Printers(strServer)

        if Err.Number = kErrorSuccess then

            wscript.echo ""
            wscript.echo "ServerName      : " & oPrinter.ServerName
            wscript.echo "PrinterName     : " & oPrinter.PrinterName
            wscript.echo "ShareName       : " & oPrinter.ShareName
            wscript.echo "DriverName      : " & oPrinter.DriverName
            wscript.echo "PortName        : " & oPrinter.PortName
            wscript.echo "Comment         : " & oPrinter.Comment
            wscript.echo "Location        : " & oPrinter.Location
            wscript.echo "SepFile         : " & oPrinter.SepFile
            wscript.echo "PrintProcesor   : " & oPrinter.PrintProcessor
            wscript.echo "DataType        : " & oPrinter.DataType
            wscript.echo "Parameters      : " & oPrinter.Parameters
            wscript.echo "Attributes      : " & CSTR(oPrinter.Attributes)
            wscript.echo "Priority        : " & CSTR(oPrinter.Priority)
            wscript.echo "DefaultPriority : " & CStr(oPrinter.DefaultPriority)
            wscript.echo "StartTime       : " & CStr(oPrinter.StartTime)
            wscript.echo "UntilTime       : " & CStr(oPrinter.UntilTime)
            wscript.echo "Status          : " & CStr(oPrinter.Status)
            wscript.echo "Jobc Count      : " & CStr(oPrinter.Jobs)
            wscript.echo "AveragePPM      : " & CStr(oPrinter.AveragePPM)

            Err.Clear

        else

            wscript.echo "Unable to list printers, error: 0x" & _
                         Hex(Err.Number) & ". " & Err.Description

            ListPrinters = kErrorFailure

            exit function

        end if

    next

    wscript.echo "Success listing printers"

    ListPrinters = kErrorSuccess

end function

'
' Delete all local printers
'
function DelPrinterAll(strServer)

    on error resume next

    DebugPrint kDebugTrace, "In DelPrinterAll"

    dim oMaster
    dim oPrinter
    dim iTotal
    dim iCount

    '
    ' Number of local printers found
    '
    iTotal = 0

    '
    ' Number of local printes deleted
    '
    iCount = 0

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    for each oPrinter in oMaster.Printers(strServer)

        if Err.Number = kErrorSuccess then

            '
            ' If strServer is not empty, the enumeration will contain
            ' only local printers on that server. The reason is that
            ' connections are per user resources
            '
            if strServer = "" then

                '
                ' Connections are enumerated in this case
                '
                if (oPrinter.Attributes and kLocalPrinterFlag) = kLocalPrinterFlag then

                    iTotal = iTotal + 1

                    oMaster.PrinterPurge strServer, oPrinter.PrinterName

                    '
                    ' In some cases, we can delete a printer, but not purge it
                    ' We need to clear the error
                    '
                    Err.Clear

                    if DelPrinter(strServer, oPrinter.PrinterName) = kErrorSuccess then

                        iCount = iCount + 1

                    end if

                end if

            else

                '
                ' Only local printers on the server are enumerated
                '
                iTotal = iTotal + 1

                oMaster.PrinterPurge strServer, oPrinter.PrinterName

                '
                ' In some cases, we can delete a printer, but not purge it
                ' We need to clear the error
                '
                Err.Clear

                if DelPrinter(strServer, oPrinter.PrinterName) = kErrorSuccess then

                    iCount = iCount + 1

                end if

            end if

            Err.Clear

        else

            wscript.echo "Unable to enumerate printers on server, error: 0x" _
                         & Hex(Err.Number) & ". " & Err.Description

            DelPrinterAll = kErrorFailure

            exit function

        end if

    next

    wscript.echo "Number of local printers found " & iTotal & ". Printers deleted " & iCount

    DelPrinterAll = kErrorSuccess

end function

'
' Delete all printer connections
'
function DelPrinterConnectionAll()

    on error resume next

    DebugPrint kDebugTrace, "In DelPrinterConnectionAll"

    dim oMaster
    dim oPrinter
    dim iTotal
    dim iCount

    '
    ' Total number of connections found
    '
    iTotal = 0

    '
    ' Total number of connections deleted
    '
    iCount = 0

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    for each oPrinter in oMaster.Printers

        if Err.Number = kErrorSuccess then

            '
            ' Test if the printer is not local
            '
            if (oPrinter.Attributes and kLocalPrinterFlag) <> kLocalPrinterFlag then

                iTotal = iTotal + 1

                if DelPrinterConnection(oPrinter.PrinterName) = kErrorSuccess then

                    iCount = iCount + 1

                end if

                Err.Clear

            end if

        else

            wscript.echo "Unable to enumerate printers on server, error: 0x" _
                         & Hex(Err.Number) & ". " & Err.Description

            DelPrinterConnectionAll = kErrorFailure

            exit function

        end if

    next

    wscript.echo "Number of connections found " & iTotal & ". Connections deleted " & iCount

    DelPrinterConnectionAll = kErrorSuccess

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
function ParseCommandLine(iAction, strServer, strPrinter, strDriver, strDriverPath, strPort, strInfFile)

    on error resume next

    DebugPrint kDebugTrace, "In the ParseCommandLine"

    dim oArgs
    dim iIndex

    iAction = kActionUnknown
    iIndex  = 0

    set oArgs = wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-a"
                iAction = kActionAdd

            case "-ac"
                iAction = kActionAddConn

            case "-d"
                iAction = kActionDel

            case "-dc"
                iAction = kActionDelConn

            case "-l"
                iAction = kActionList

            case "-x"
                iAction = kActionDelAll

            case "-xc"
                iAction = kActionDelConnAll

            case "-c"
                iIndex = iIndex + 1
                strServer = oArgs(iIndex)

            case "-b"
                iIndex = iIndex + 1
                strPrinter = oArgs(iIndex)

            case "-f"
                iIndex = iIndex + 1
                strInfFile = oArgs(iIndex)

            case "-m"
                iIndex = iIndex + 1
                strDriver = oArgs(iIndex)

            case "-r"
                iIndex = iIndex + 1
                strPort = oArgs(iIndex)

            case "-p"
                iIndex = iIndex + 1
                strDriverPath = oArgs(iIndex)

            case "-?"
                Usage(true)
                exit function

            case else
                Usage(true)
                exit function

        end select

        iIndex = iIndex + 1

    wend

    if Err = kErrorSuccess then

        ParseCommandLine = kErrorSuccess

    else

        wscript.echo "Unable to parse command line, error 0x" & _
                     Hex(Err.Number) & ". " & Err.Description

        ParseCommandLine = kErrorFailure

    end if

end  function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo "Usage: prnmgr [-adl?][c] [-c server][-b printer][-m driver model]"
    wscript.echo "              [-p driver path][-r port][-f file]"
    wscript.echo "Arguments:"
    wscript.echo "-a     - add local printer"
    wscript.echo "-ac    - add printer connection"
    wscript.echo "-d     - delete local printer"
    wscript.echo "-dc    - delete printer connection"
    wscript.echo "-f     - inf file"
    wscript.echo "-l     - list printers"
    wscript.echo "-c     - server name"
    wscript.echo "-b     - printer name"
    wscript.echo "-m     - driver model"
    wscript.echo "-r     - port name"
    wscript.echo "-p     - driver path can be local or network path i.e. a:\ or \\server\share"
    wscript.echo "-x     - delete all local printers"
    wscript.echo "-xc    - delete all printer connections, cannot be used with the -c option"
    wscript.echo "-?     - display command usage"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "prnmgr -l -c \\server"
    wscript.echo "prnmgr -a -b ""printer"" -m ""driver"" -r ""lpt1:"""
    wscript.echo "prnmgr -d -b ""printer"" -c \\server"
    wscript.echo "prnmgr -ac -b ""\\server\printer"""
    wscript.echo "prnmgr -dc -b ""\\server\printer"""
    wscript.echo "prnmgr -x"
    wscript.echo "prnmgr -x -c \\server"
    wscript.echo "prnmgr -xc"

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

