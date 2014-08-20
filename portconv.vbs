'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' Abstract:
'
' portconv.vbs - Script for converting lpr ports to tcp ports
'
' Usage:
' portconv [-ag?][-p port][-c source server][-i ip][-d destination server]
'
' Examples:
' portconv -g -i 1.2.3.4
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
const kActionAdd        = 1
const kActionGet        = 2
const kActionConvertAll = 3

const kErrorSuccess = 0
const KErrorFailure = 1

'
' Port Types
'
const kTcpRaw   = 1
const kTcpLPr   = 2
const kLocal    = 3
const kLprMon   = 5
const kHPdlc    = 7
const kUnknown  = 8

'
' The only supported conversion is from lpr mon to tcp
'
const kLprToTcp = 1

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strIPAddress
    dim strSourceServer
    dim strDestServer
    dim strPort

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(kMessage1 & vbCRLF & kMessage2 & vbCRLF & _
                          kMessage3 & vbCRLF & kMessage4 & vbCRLF & _
                          kMessage5 & vbCRLF & kMessage6 & vbCRLF)

        wscript.quit

    end if

    iRetval = ParseCommandLine(iAction, strSourceServer, strDestServer, strPort, strIPAddress)

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionAdd
                 iRetval = AddEquivalentTCPPort(strSourceServer, strDestServer, strPort)

            case kActionGet
                 iRetval = GetEquivalentTCPSettings(strIPAddress)

            case kActionConvertAll
                 iRetval = ConvertAll(strSourceServer, strDestServer)

            case else
                 Usage(true)

        end select

    end if

end sub

'
' Get the TCP equivalent of a lpr mon port
'
function GetEquivalentTCPSettings(strIPAddress)

    on error resume next

    DebugPrint kDebugTrace, "In GetEquivalentTCPSettings"

    dim oMaster
    dim oPort
    dim iResult

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")
    set oPort   = CreateObject("Port.Port.1")

    oPort.HostAddress = strIPAddress

    '
    ' oPort will contain the settings of the device
    '
    oMaster.PortConversion oPort, kLprToTcp

    if Err.Number <> kErrorSuccess then

        wscript.echo "Error converting the port, error 0x" & _
                     Hex(Err.Number) & ". " & Err.Description

        iResult = kErrorFailure

    else

        '
        ' Check if the device responded
        '
        if oPort.DeviceType <> "" then

            wscript.echo "DeviceType  " & oPort.DeviceType

        else

            wscript.echo "The device did not respond. The default port settings will be displayed"

        end if

        wscript.echo "Name        " & oPort.PortName
        wscript.echo "HostAddress " & oPort.HostAddress

        if oPort.PortType = kTcpRaw then

            wscript.echo "Protocol    RAW"

            wscript.echo "PortNumber  " & oPort.PortNumber

        else

            wscript.echo "Protocol    LPR"

            wscript.echo "Queue       " & oPort.QueueName

        end if

        if oPort.SNMP then

            wscript.echo "SNMP        Enabled"

        else

            wscript.echo "SNMP        Disabled"

        end if

        if oPort.DoubleSpool then

            wscript.echo "DoubleSpool Yes"

        else

            wscript.echo "DoubleSpool No"

        end if

        iResult = kErrrorSuccess

    end if

    GetEquivalentTCPSettings = iResult

end function

'
' Add an equivalent tcp port. strSource is the server where strPort is on.
' If strPort is a lpr mon port, a corresponding tcp port will be added to
' the destination server strDestServer
'
function AddEquivalentTCPPort(strSourceServer, strDestServer, strPort)

    on error resume next

    DebugPrint kDebugTrace, "In AddEquivalentTCPPort"

    dim oMaster
    dim oPort
    dim iResult

    iResult = kErrorFailure

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")
    set oPort   = CreateObject("Port.Port.1")

    '
    ' Get the configuration of the port to be converted
    '
    oMaster.PortGet strSourceServer, strPort, oPort

    if Err.Number <> kErrorSuccess then

        wscript.echo "Unable to get the configuration for the port """ & strPort & _
                     """, error 0x" & Hex( Err.Number ) & " " & Err.Description

    else

        '
        ' Check if it is lpr mon port
        '
        if oPort.PortType = kLprMon then

           '
           ' Attempt to get the equivalent tcp port
           '
           oMaster.PortConversion oPort, kLprToTcp

           if Err.Number <> kErrorSuccess then

               wscript.echo "Unable to convert the port, error: 0x" & Hex(Err.Number) _
                            & " " & Err.Description

           else

              '
              ' An empty DeviceType means the device did not respond or couldn't be identified
              '
              if oPort.DeviceType = "" then

                  wscript.echo "The device did not respond. A port with default settings will be added"

              end if

              oPort.ServerName = strDestServer

              '
              ' Add the equivalent port
              '
              oMaster.PortAdd oPort

              if Err.Number = kErrorSuccess then

                  wscript.echo "Success adding the TCP port: """ & oPort.PortName & """ on server " & strDestServer

                  iResult = kErrorSuccess

              else

                  wscript.echo "Unable to add the TCP port, error: 0x" & _
                               Hex(Err.Number) & ". " & Err.Description

              end if

           end if

        else

            wscript.echo "Error: This port is not lpr mon"

        end if

    end if

    AddEquivalentTCPPort = iResult

end function

'
' Convert all lpr mon ports from the source server onto the destination server
'
function ConvertAll(strSourceServer, strDestServer)

    on error resume next

    DebugPrint kDebugTrace, "In ConvertAll"

    dim oMaster
    dim oPort
    dim iTotal
    dim iLprCount
    dim iTcpCount

    '
    ' Total number of ports on the source server
    '
    iTotal = 0

    '
    ' Total number of lpr mon ports on the source server
    '
    iLprCount = 0

    '
    ' Total number of equivalent tcp ports added on the destination server
    '
    iTcpCount = 0

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    for each oPort in oMaster.Ports(strSourceServer)


        if Err.Number = kErrorSuccess then

            iTotal = iTotal + 1

            if oPort.PortType = kLprMon then

                iLprCount = iLprCount + 1

                oMaster.PortConversion oPort, kLprToTcp

                if Err.Number = kErrorSuccess then

                    '
                    ' Check if the device is responding
                    '
                    if oPort.DeviceType = "" then

                        wscript.echo "The device " & oPort.HostAddress & _
                                     " did not respond. Adding a port with default settings"

                        '
                        ' Enable LPR byte counting
                        '
                        oPort.DoubleSpool = true

                    else

                        wscript.echo oPort.HostAddress & " is " & oPort.DeviceType

                    end if

                    oPort.ServerName = strDestServer

                    '
                    ' Add the equivalent port
                    '
                    oMaster.PortAdd oPort

                    if Err.Number = kErrorSuccess then

                        iTcpCount = iTcpCount + 1

                        wscript.echo oPort.PortName & " was added"

                    else

                        wscript.echo "Unable to add """ & oPort.PortName & """, error: 0x" _
                                      & Hex(Err.Number) & ". " & Err.Description

                        Err.Clear

                    end if

                else

                    wscript.echo "Unable to convert port """ & oPort.PortName & """ , error: 0x" _
                                 & Hex(Err.Number) & ". " & Err.Description

                    Err.Clear

                end if

            end if

        else

            wscript.echo "Unable to list ports, error: 0x" & Hex(Err.Number) & ". " & Err.Description

            ConvertAll = kErrorFailure

            exit function

        end if

    next

    wscript.echo "Number of ports on the source server                    " & iTotal
    wscript.echo "Number of lpr mon ports on the source server            " & iLprCount
    wscript.echo "Number of tcp ports added to the the destination server " & iTcpCount

    ConvertAll = kErrorSuccess

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
function ParseCommandLine(iAction, strSourceServer, strDestServer, strPort, strIPAddress)

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
                iAction = kActionAdd

            case "-g"
                iAction = kActionGet

            case "-w"
                iAction = kActionConvertAll

            case "-i"
                iIndex = iIndex + 1
                strIPAddress = oArgs(iIndex)

            case "-p"
                iIndex = iIndex + 1
                strPort = oArgs(iIndex)

            case "-c"
                iIndex = iIndex + 1
                strSourceServer = oArgs(iIndex)

            case "-d"
                iIndex = iIndex + 1
                strDestServer = oArgs(iIndex)

            case "-?"
                Usage(true)
                exit function

            case else
                Usage(true)
                exit function

        end select

        iIndex = iIndex + 1

    wend

    if Err.Number <> kErrorSuccess then

        wscript.echo "Unable to parse command line, error 0x" & _
                     Hex(Err.Number) & " " & Err.Description

        ParseCommandLine = kErrorFailure

    else

        ParseCommandLine = kErrorSuccess

    end if

end  function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo "Usage: portconv [-agw?][-p port][-c source server]"
    wscript.echo "                       [-i ip][-d destination server]"
    wscript.echo "Arguments:"
    wscript.echo "-a     - adds the equivalent tcp port for an lpr port"
    wscript.echo "-g     - for an IP address, gets the preferred device settings"
    wscript.echo "-w     - convert all"
    wscript.echo "-p     - port name"
    wscript.echo "-c     - source server name"
    wscript.echo "-i     - ip address of the device to get the settings of"
    wscript.echo "-d     - destination server name, where the port will be added"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "portconv -g -i 1.2.3.4"
    wscript.echo "portconv -a -p 1.2.3.4:Queue -c \\server"
    wscript.echo "portconv -a -p 1.2.3.4:Queue -c \\server -d \\dest"
    wscript.echo "portconv -w -c \\server"
    wscript.echo "portconv -w -c \\server -d \\dest"
    wscript.echo "portconv -w -d \\dest"

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

