'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' Abstract:
'
' PortMgr.vbs - Port operation script for Windows .NET Server 2003
'
' Usage:
'
' Usage: portmgr [-adl?] [-p port] [-c server][-n port number]
'                        [-t raw|lpr|local] [-e device name]"
'
' Examples
' PortMgr -d -c \\server -p c:\temp\foo.prn
' PortMgr -a -c \\server -p IP_1.2.3.4 -e 1.2.3.4 -t raw -n 9100
'
'----------------------------------------------------------------------

option explicit

'
' Debugging trace flags, to disable debug output trace message
' change gDebugFlag to true.
'
dim   gDebugFlag
const kDebugTrace = 1
const kDebugError = 2

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
const kActionAdd        = 0
const kActionDelete     = 1
const kActionList       = 2
const kActionUnknown    = 3
const kActionGet        = 4
const kActionSet        = 5

'
' Port Types
'
const kTcpRaw           = 1
const kTcpLPr           = 2
const kLocal            = 3
const kLocalDownLevel   = 4
const kLprMon           = 5
const kHPdlc            = 7
const kUnknown          = 8

const kStdTCPIP = "Standard TCP/IP Port"

const kErrorSuccess = 0
const KErrorFailure = 1

main

'
' Main execution starts here
'
sub main

    on error resume next

    dim iAction
    dim iRetval
    dim oParamDict

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(kMessage1 & vbCRLF & kMessage2 & vbCRLF & _
                          kMessage3 & vbCRLF & kMessage4 & vbCRLF & _
                          kMessage5 & vbCRLF & kMessage6 & vbCRLF)

        wscript.quit

    end if

    set oParamDict = CreateObject("Scripting.Dictionary")

    iRetval = ParseCommandLine(iAction, oParamDict)

    if iRetval = 0 then

        select case iAction

            case kActionAdd
                iRetval = AddPort(oParamDict)

            case kActionDelete
                iRetval = DeletePort(oParamDict.Item("ServerName"), oParamDict.Item("PortName"))

            case kActionList
                iRetval = ListPorts(oParamDict.Item("ServerName"))

            case kActionGet
                iRetVal = GetPort(oParamDict.Item("ServerName"), oParamDict.Item("PortName"))

            case kActionSet
                iRetVal = SetPort(oParamDict)

            case else
                Usage(true)
                exit sub

        end select

    end if

end sub

'
' Delete a port
'
function DeletePort(strServer, strPort)

    on error resume next

    dim oMaster, oPort
    dim iResult

    DebugPrint kDebugTrace, "In DeletePort "
    DebugPrint kDebugTrace, "Server = " & strServer
    DebugPrint kDebugTrace, "Port   = " & strPort

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")
    set oPort   = CreateObject("Port.Port.1")

    oPort.ServerName = strServer

    oPort.PortName   = strPort

    oMaster.PortDel oPort

    if Err.Number = 0 then

        wscript.echo "Success: Deleting port """ & strPort & """ "

        iResult = kErrorSuccess

    else

        wscript.echo "Unable to deletie port """ & strPort & """, error: 0x" _
                     & Hex(Err.Number) & ". " & Err.Description

        iResult = kErrorFailure

    end if

    DeletePort = iResult

end function

'
' Add a port
'
function AddPort(oParamDict)

    on error resume next

    dim oPort
    dim oMaster
    dim iResult
    dim PortType

    DebugPrint kDebugTrace, "In AddPort "
    DebugPrint kDebugTrace, "ServerName  = " & oParamDict.Item("ServerName")
    DebugPrint kDebugTrace, "PortName    = " & oParamDict.Item("PortName")
    DebugPrint kDebugTrace, "PortType    = " & oParamDict.Item("PortType")
    DebugPrint kDebugTrace, "PortNumber  = " & oParamDict.Item("PortNumber")
    DebugPrint kDebugTrace, "QueueName   = " & oParamDict.Item("QueueName")
    DebugPrint kDebugTrace, "Index       = " & oParamDict.Item("Index")
    DebugPrint kDebugTrace, "Community   = " & oParamDict.Item("CName")
    DebugPrint kDebugTrace, "HostAddress = " & oParamDict.Item("HostAddress")

    set oPort   = CreateObject("Port.Port.1")
    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    oPort.ServerName = oParamDict.Item("ServerName")
    oPort.PortName   = oParamDict.Item("PortName")
    PortType         = oParamDict.Item("PortType")

    '
    ' Update the port object with the settings corresponding
    ' to the port type of the port to be added
    '
    select case lcase(PortType)

            case "raw"
                 oPort.PortType    = kTcpRaw
                 oPort.HostAddress = oParamDict.Item("HostAddress")
                 oPort.PortNumber  = oParamDict.Item("PortNumber")

                 if oParamDict.Exists("SNMP") then

                     oPort.SNMP = oParamDict("SNMP")

                 end if

                 if oParamDict.Exists("SNMPDeviceIndex") then

                     oPort.SNMPDeviceIndex = oParamDict.Item("SNMPDeviceIndex")

                 end if

                 if oParamDict.Exists("CommunityName") then

                         oPort.CommunityName = oParamDict.Item("CommunityName")

                 end if

            case "lpr"
                 oPort.PortType    = kTcpLpr
                 oPort.HostAddress = oParamDict.Item("HostAddress")
                 oPort.QueueName   = oParamDict.Item("QueueName")

                 if oParamDict.Exists("SNMP") then

                     oPort.SNMP = oParamDict("SNMP")

                 end if

                 if oParamDict.Exists("SNMPDeviceIndex") then

                     oPort.SNMPDeviceIndex = oParamDict.Item("SNMPDeviceIndex")

                 end if

                 if oParamDict.Exists("CommunityName") then

                         oPort.CommunityName = oParamDict.Item("CommunityName")

                 end if

                 oPort.DoubleSpool = oParamDict.Item("DoubleSpool")

            case "local"
                 oPort.PortType = kLocal

            case "localdownlevel"
                 oPort.PortType = kLocalDownLevel

            case else
                 wscript.echo "Unable to add port, error: invalid port type """ & PortType & """ "
                 exit function
    end select

    '
    ' Try adding the port
    '
    oMaster.PortAdd oPort

    if Err.Number = kErrorSuccess then

        wscript.echo "Success: Adding port """ & oPort.PortName & """ "

        iResult = kErrorSuccess

    else

        wscript.echo "Error: Adding port """ & oPort.PortName & """, error: 0x" _
                     & Hex(Err.Number) & ". " & Err.Description

        iResult = kErrorFailure

    end if

    AddPort = iResult

end function

'
' List ports on a machine.
'
function ListPorts(strServer)

    on error resume next

    DebugPrint kDebugTrace, "In ListPorts"

    dim oMaster
    dim oPort
    dim oError
    dim iResult

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    for each oPort in oMaster.Ports(strServer)

        if Err = kErrorSuccess then

            wscript.echo ""

            wscript.echo "ServerName   " & oPort.ServerName

            wscript.echo "PortName     " & oPort.PortName

            '
            ' HPdlc and lpr mon ports don't have a MonitorName or Description set
            '
            if oPort.MonitorName <> "" then

                wscript.echo "MonitorName  " & oPort.MonitorName

            end if

            if oPort.Description <> "" then

                wscript.echo "Description  " & oPort.Description

            end if

            if oPort.PortType = kLprMon or oPort.PortType = kHPdlc then

                wscript.echo "HostAddress  " & oPort.HostAddress

                wscript.echo "Queue        " & oPort.QueueName

                wscript.echo "PortType     " & Description(oPort.PortType)

            end if

            if oPort.Description = kStdTCPIP then

                wscript.echo "Getting extended information about the TCP port"

                iResult = GetPort(strServer, oPort.PortName)

            end if

        else

            wscript.echo "Unable to list ports, error: 0x" & Hex(Err.Number) _
                         & ". " & Err.Description

            ListPorts = kErrorFailure

            exit function

        end if

    next

    wscript.echo "Success: Lising ports"

    ListPorts = kErrorSuccess

end function

'
' Gets the configuration of a port
'
function GetPort(strServerName, strPortName)

    on error resume next

    DebugPrint kDebugTrace, "In GetPort"

    dim oPort
    dim oMaster
    dim iResult

    set oPort   = CreateObject("Port.Port.1")
    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    '
    ' Get the configuration
    '
    oMaster.PortGet strServerName, strPortName, oPort

    if Err.Number = kErrorSuccess then

        wscript.echo "PortName     " & strPortName

        wscript.echo "PortType     " & Description(oPort.PortType)

        wscript.echo "Host         " & oPort.HostAddress

        if oPort.PortType = kTcpLpr or oPort.PortType = kLprMon then

            wscript.echo "QueueName    " & oPort.QueueName

        end if

        if oPort.PortType=kTcpRaw then

            wscript.echo "PortNumber   " & CStr(oPort.PortNumber)

        end if

        if oPort.PortType = kTcpRaw or oPort.PortType = kTcpLpr then

            if oPort.SNMP then

                wscript.echo "SNMP         Enabled"

                wscript.echo "SNMP Index   " & CStr(oPort.SNMPDeviceIndex)

                wscript.echo "Community    " & oPort.CommunityName

            else

                wscript.echo "SNMP         Disabled"

            end if

        end if

        if oPort.PortType = kTcpLpr then

            if oPort.DoubleSpool then

                wscript.echo     "Byte count   Enabled"

            else

                wscript.echo     "Byte count   Disabled"

            end if

        end if

        iResult = kErrorSuccess

    else

        wscript.echo "Unable to get port configuration, error: 0x" & _
                     Hex(Err.Number) & ". " & Err.Description

        Err.Clear

        iResult = kErrorFailure

    end if

    GetPort = iResult

end function

'
' Set the configuration of a port
'
function SetPort(oParamDict)

    on error resume next

    dim oPort
    dim oMaster
    dim iResult

    set oPort   = CreateObject("Port.Port.1")
    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    '
    ' Get the configuration of the port
    '
    oMaster.PortGet oParamDict.Item("ServerName"), oParamDict.Item("PortName"), oPort

    if Err.Number <> kErrorSuccess then

        wscript.echo "Unable to get port configuration, error: 0x" & Hex(Err.Number) _
                     & " " & Err.Description

        iResult = kErrorFailure

        exit function

    end if

    '
    ' Update the oPort object with the settings specified by the user
    '
    if oParamDict.Item("PortType") = "raw" then

        oPort.PortType = kTcpRaw

    elseif oParamDict.Item("PortType") = "lpr" then

        oPort.PortType = kTcpLpr

    end if

    oPort.HostAddress = oParamDict.Item("HostAddress")

    oPort.PortNumber  = oParamDict.Item("PortNumber")

    if oParamDict.Exists("QueueName") then

        oPort.QueueName = oParamDict.Item("QueueName")

    end if

    if oParamDict.Exists("SNMP") then

        oPort.SNMP = oParamDict.Item("SNMP")

    end if

    oPort.SNMPDeviceIndex = oParamDict.Item("SNMPDeviceIndex")

    if oParamDict.Exists("CommunityName") then

        oPort.CommunityName = oParamDict.Item("CommunityName")

    end if

    if oParamDict.Exists("DoubleSpool") then

        oPort.DoubleSpool = oParamDict.Item("DoubleSpool")

    end if

    '
    ' Set the port
    '
    oMaster.PortSet oPort

    if Err.Number = kErrorSuccess then

        wscript.echo "Success: Updating port " & oPort.PortName

        iResult = kErrorSuccess

    else

        wscript.echo "Unable to update port settings , error: 0x" & _
                     Hex(Err.Number) & ". " & Err.Description

        iResult = kErrorFailure

    end if

    SetPort = iResult

end function

'
' Get a string description for a port type
'
function Description(value)

    on error resume next

    select case value

           case kTcpRaw

                Description = "TCP RAW"

           case kTcpLpr

                Description = "TCP LPR"

           case kLocal

                Description = "Standard Local"

           case kLocalDownLevel

                Description = "Standard Local Down Level"

           case kLprMon

                Description = "LPR Mon"

           case kHPdlc

                Description = "HP DLC"

           case kUnknown

                Description = "Unknown Port"

           case Else

                Description = "Invalid PortType"

    end select

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
' Parse the command line into its components
'
function ParseCommandLine(iAction, oParamDict)

    on error resume next

    DebugPrint kDebugTrace, "In ParseCommandLine"

    dim oArgs
    dim iIndex

    iAction = kActionUnknown

    set oArgs = Wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-g"
                iAction = kActionGet

            case "-s"
                iAction = kActionSet

            case "-a"
                iAction = kActionAdd

            case "-d"
                iAction = kActionDelete

            case "-l"
                iAction = kActionList

            case "-2e"
                oParamDict.Add "DoubleSpool", true

            case "-2d"
                oParamDict.Add "DoubleSpool", false

            case "-c"
                iIndex = iIndex + 1
                oParamDict.Add "ServerName", oArgs(iIndex)

            case "-n"
                iIndex = iIndex + 1
                oParamDict.Add "PortNumber", oArgs(iIndex)

            case "-p"
                iIndex = iIndex + 1
                oParamDict.Add "PortName", oArgs(iIndex)

            case "-t"
                iIndex = iIndex + 1
                oParamDict.Add "PortType", oArgs(iIndex)

            case "-h"
                iIndex = iIndex + 1
                oParamDict.Add "HostAddress", oArgs(iIndex)

            case "-q"
                iIndex = iIndex + 1
                oParamDict.Add "QueueName", oArgs(iIndex)

            case "-i"
                iIndex = iIndex + 1
                oParamDict.Add "SNMPDeviceIndex", oArgs(iIndex)

            case "-y"
                iIndex = iIndex + 1
                oParamDict.Add "CommunityName", oArgs(iIndex)

            case "-me"
                oParamDict.Add "SNMP", true

            case "-md"
                oParamDict.Add "SNMP", false

            case "-?"
                Usage(True)
                exit function

            case else
                Usage(True)
                exit function

        end select

        iIndex = iIndex + 1

    wend

    if Err = kErrorSuccess then

        ParseCommandLine = kErrorSuccess

    else

        wscript.echo "Unable to parse command line, error 0x" _
                     & Hex(Err.Number) & ". " & Err.Description

        ParseCommandLine = kErrorFailure

    end if

end  function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo "Usage: portmgr [-adlgs?] [-p port] [-c server] [-n number]"
    wscript.echo "               [-t raw|lpr|local] [-h host address] [-q queue]"
    wscript.echo "               [-me | -md ] [-i SNMP index] [-y community] [-2e | -2d]"
    wscript.echo "Arguments:"
    wscript.echo "-a     - add a port"
    wscript.echo "-d     - delete the specified port"
    wscript.echo "-l     - list all ports"
    wscript.echo "-g     - get configuration for a TCP port"
    wscript.echo "-s     - set configuration for a TCP port"
    wscript.echo "-p     - port name"
    wscript.echo "-c     - server name"
    wscript.echo "-n     - port number, applies to TCP RAW ports"
    wscript.echo "-t     - port type"
    wscript.echo "-h     - IP address of the device"
    wscript.echo "-q     - queue name, applies to TCP LPR ports"
    wscript.echo "-m     - SNMP type. [e] enalbe, [d] disable"
    wscript.echo "-i     - SNMP index, if SNMP is enabled"
    wscript.echo "-y     - community name, if SNMP is enabled"
    wscript.echo "-2     - double spool, applies to TCP LPR ports.[e] enalbe, [d] disable"
    wscript.echo "-?     - display command usage"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "portmgr -l -c \\server"
    wscript.echo "portmgr -d -c \\server -p c:\temp\foo.prn"
    wscript.echo "portmgr -a -c \\server -p test -t local"
    wscript.echo "portmgr -a -c \\server -p IP_1.2.3.4 -h 1.2.3.4 -t raw -n 9100"
    wscript.echo "portmgr -s -c \\server -p IP_1.2.3.4 -me -y public -i 1 -n 9100"

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

