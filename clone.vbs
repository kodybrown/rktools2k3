'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' Abstract:
'
'    clone.vbs - printer server cloning script for Windows .NET Server 2003
'
' Usage:
'
'    clone [-dopfa?] [-c server-name]
'
' Examples:
'    clone -d
'    clone -o -c \\server
'    clone -p -c \\server
'    clone -f
'    clone -a
'
'----------------------------------------------------------------------

option explicit

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
' Default script names
'
const kDriverScript   =   "drv_clone.vbs"
const kPortScript     =   "port_clone.vbs"
const kPrinterScript  =   "prn_clone.vbs"
const kFormScript     =   "form_clone.vbs"

'
' The shell script for installing all components
'
const kInstallScript  =   "install.bat"

'
' Strings identifying environments
'
const kEnvironmentIntel   = "Windows NT x86"
const kEnvironmentMIPS    = "Windows NT R4000"
const kEnvironmentAlpha   = "Windows NT Alpha_AXP"
const kEnvironmentPowerPC = "Windows NT PowerPC"
const kEnvironmentWindows = "Windows 4.0"
const kEnvironmentUnknown = "unknown"

'
' Strings identifying architectures
'
const kArchIntel          = "Intel"
const kArchMIPS           = "MIPS"
const kArchAlpha          = "Alpha"
const kArchPowerPC        = "PowerPC"
const kArchUnknown        = "Unknown"

'
' Strings identifying driver versions
' Change these strings on localized builds
'
const kVersionWindows95   = "Windows 95 or 98"
const kVersion_NT31       = "Windows NT 3.1"
const kVersion35x         = "Windows NT 3.5 or 3.51"
const kVersion351         = "Windows NT 3.51"
const kVersion40          = "Windows NT 4.0"
const kVersion4050        = "Windows NT 4.0 or 2000"
const kVersion50          = "Windows 2000"

'
' Action codes
'
const kActionUnknown  = 0
const kActionDrivers  = 1
const kActionPorts    = 2
const kActionPrinters = 3
const kActionForms    = 4
const kActionAll      = 5

'
' Port types
'
const kTcpRaw   = 1
const kTcpLpr   = 2
const kLocal    = 3
const kLprMon   = 5
const kHPdlc    = 7
const kUnknown  = 8

const kStdTCPIP = "Standard TCP/IP Port"

'
' Printer attribute masks
'
const kPrinterNetwork = 16
const kPrinterLocal   = 64

'
'  Search patterns used in the regular expressions
'  to skip COMx: and LPTx: ports
'
const kComPortPattern = "^COM\d+:$"
const kLptPortPattern = "^LPT\d+:$"

'
' Persist constant
'
const kAllSettings    = 127

'
' Other string constants
'
const kTrueStr   =  "true"
const kFalseStr  =  "false"
const kLongLineStr  =    _
  "'--------------------------------------------------------------------------"

'
' Source server that will be cloned
'
dim strServerName
dim strPrefixServerName

'
' For writing the script file
'
dim oFileSystem
dim oScript

'
' Script name string
'
dim strScriptName
dim iAction

main

'
' Main execution starts here
'
sub main

   '
   ' Abort if the host is not cscript
   '
   if not IsHostCscript() then

       call wscript.echo(kMessage1 & vbCRLF & kMessage2 & vbCRLF & _
                         kMessage3 & vbCRLF & kMessage4 & vbCRLF & _
                         kMessage5 & vbCRLF & kMessage6 & vbCRLF)

       wscript.quit

   end if

   ParseCommandLine

   '
   ' Check if a machine name was passed
   '
   if (strServerName <> "") then

       strPrefixServerName = strServerName

   else

       strPrefixServerName =  strGetLocalMachineName()

   end if

   '
   ' Remove the "\\" in front of the machine name
   '
   strPrefixServerName = strGetNameStringOnly(strPrefixServerName) & "_"

   set oFileSystem = CreateObject ("Scripting.FileSystemObject")

   '
   ' According to the different action codes, create the corresponding script
   '
   if iAction = kActionDrivers or iAction = kActionAll then

       strScriptName  = strPrefixServerName & kDriverScript
       DriverCloneScript strScriptName, strServerName

   end if

   if iAction = kActionPorts or iAction = kActionAll then

       strScriptName  = strPrefixServerName & kPortScript
       PortCloneScript strScriptName, strServerName

   end if

   if iAction = kActionPrinters or iAction = kActionAll then

       strScriptName  = strPrefixServerName & kPrinterScript
       PrinterCloneScript strScriptName, strServerName

   end if

   if iAction = kActionForms or iAction = kActionAll then

       strScriptName  = strPrefixServerName & kFormScript
       FormCloneScript strScriptName, strServerName

   end if

   '
   '  Generate the script for launching all the cloning scripts
   '
   strScriptName = strPrefixServerName & kInstallScript
   InstallScript strScriptName, strPrefixServerName

end sub

'
'------------------------------------------------------------------------
'      Driver cloning script
'------------------------------------------------------------------------
'
sub DriverCloneScript(ByVal strScriptName, ByVal strServerName)

    on error resume next

    wscript.echo
    wscript.echo "Creating the driver cloning script..."

    '
    ' Open the script file
    '
    set oScript = oFileSystem.CreateTextFile(strScriptName,TRUE)

    '
    ' Write the header of the driver cloning script
    '
    DriverStartUp

    ws("    EchoLine kNormal, """" ")
    ws("    EchoLine kNormal, ""------------------------------"" ")
    ws("    EchoLine kNormal, ""Start installing drivers..."" ")

    '
    ' Enumerate all the drivers in server "strServerName", for each driver
    ' found, add a line in the script to call the AddDriver subroutine
    '
    dim oMaster
    dim oDriver
    dim iDriverCount

    iDriverCount = 0
    set oMaster  = CreateObject("PrintMaster.PrintMaster.1")
    for each oDriver in oMaster.Drivers(strServerName)

        if Err = 0 then

            '
            ' Add a call to "AddDriver" function in the script
            '
            iDriverCount = iDriverCount + 1
            ws("    AddDriver   strDestServer, _")
            ws("                """ & oDriver.ModelName     & """,  _")
            ws("                """ & GetVersion(oDriver.Version, oDriver.Environment) & """,  _")

            '
            ' Change Path to be default, i.e. ""
            '
            ws("                """",  _")
            ws("                """ & GetArchitecture(oDriver.Environment) & """,  _")

            '
            ' Change InfFile to be default, i.e. ""
            '
            ws("                """" ")

        else

            '
            ' Clean up
            '
            oScript.Close

            oFileSystem.DeleteFile strScriptName

            wscript.echo "Error: Listing ports, error: 0x" & Hex(Err.Number)
            if Err.Description <> "" then
                wscript.echo "Error description: " & Err.Description
            end if

            exit sub

        end if

    next

    wscript.echo "Success: Listing drivers on server " & strServerName

    '
    '  Write the summary script
    '
    wscript.echo "A total of " & CStr(iDriverCount) & " drivers are listed."

    ws("    EchoLine kNormal, ""Attempted to install a total of "" & CStr(iDriverCount) & "" driver(s)."" ")
    ws("    EchoLine kNormal, CStr(iSuccessCount) & "" driver(s) were successfully installed."" ")

    '
    '  Append other functions to the driver cloning script
    '
    DriverCleanUp

    '
    ' Close the script file
    '
    oScript.Close

    wscript.echo "The script file for cloning drivers is """ & strScriptName & """."

end sub

'
' Writing the script for "AddDriver" function
'
sub ScriptAddDriver

    '
    ' Insert the comment line before the function header
    '
    blank
    ws("'")
    ws("' Add a driver")
    ws("'")

    '
    ' The function header
    '
    ws("sub AddDriver(ByVal strServerName,         _")
    ws("              ByVal strModelName,          _")
    ws("              ByVal strDriverVersion,      _")
    ws("              ByVal strPath,               _")
    ws("              ByVal strDriverArchitecture, _")
    ws("              ByVal strInfFile             _")
    ws(")")
    blank

    '
    ' The function body
    '
    ws("    on error resume next")
    blank

    '
    ' Print out the information about the driver that is about to be installed
    '
    ws("    iDriverCount = iDriverCount + 1")
    ws("    EchoLine kVerbose, ""Driver:"" & CSTR(iDriverCount) ")
    ws("    EchoLine kVerbose, ""    ServerName         : "" & strServerName ")
    ws("    EchoLine kVerbose, ""    ModelName          : "" & strModelName ")
    ws("    EchoLine kVerbose, ""    DriverVersion      : "" & strDriverVersion ")
    ws("    EchoLine kVerbose, ""    DriverArchitecture : "" & strDriverArchitecture ")
    blank

    '
    ' The code that installs the driver
    '
    ws("    dim oMaster")
    ws("    dim oDriver")
    blank

    ws("    set oMaster = CreateObject(""PrintMaster.PrintMaster.1"")")
    ws("    set oDriver = CreateObject(""Driver.Driver.1"")")
    blank

    ws("    oDriver.ServerName         = strServerName")
    ws("    oDriver.ModelName          = strModelName")
    ws("    oDriver.DriverVersion      = strDriverVersion")
    ws("    oDriver.Path               = strPath")
    ws("    oDriver.DriverArchitecture = strDriverArchitecture")
    ws("    oDriver.InfFile            = strInfFile")
    blank

    ws("    oMaster.DriverAdd oDriver")
    blank

    ws("    if Err = 0 then")
    blank

    ws("        EchoLine kVerbose, ""Success: Driver "" & strModelName & "" added to server "" & strServerName ")
    ws("        iSuccessCount = iSuccessCount + 1")
    blank

    ws("    else")
    blank

    ws("        EchoLine kNormal, ""Error adding driver "" & strModelName & "", error: 0x"" & Hex(Err.Number)")
    ws("        if Err.Description <> """" then ")
    ws("            EchoLine kNormal,  ""       Error description: "" & Err.Description ")
    ws("        end if")
    ws("        Err.Clear")
    blank


    ws("    end if")
    blank

    ws("    EchoLine kVerbose, """"")
    blank

    ws("end sub")
    blank

end sub

'
' StartUp script for cloning drivers
'
sub DriverStartUp

    '
    ' Start creating the driver cloning script
    '
    CopyrightScript
    DriverAbstractScript

    '
    ' The script program starts
    '
    blank
    ws("option explicit")
    blank

    ws("'")
    ws("' Verbose Level")
    ws("'")
    ws("const kNormal    = 0")
    ws("const kVerbose   = 1")
    blank

    ws("'")
    ws("' Flag, set if the user doesn't want to replace the old forms")
    ws("'")
    ws("dim bKeepOriginalOnes")
    blank

    ws("dim strDestServer")
    blank

    ws("' The number of drivers to be installed")
    blank
    ws("dim iDriverCount")
    blank
    ws("' The number of drivers successfully installed")
    blank
    ws("dim iSuccessCount")
    blank
    ws("dim bVerbose")
    blank

    ws("main")
    blank
    ws("'")
    ws("' Main execution starts here")
    ws("'")

    ws("sub main")
    blank
    ws("    bVerbose = false")
    ws("    bKeepOriginalOnes = false")
    ws("    iDriverCount  = 0")
    ws("    iSuccessCount = 0")
    ws("    strDestServer = """"")
    ws("    ParseCommandLine")
    blank

end sub

'
' CleanUp script for cloning drivers
'
sub DriverCleanUp

    ws("end sub")

    '
    ' Append the subroutine "AddDriver"
    '
    ScriptAddDriver

    '
    ' Append the command line parsing script
    '
    ParseCommandLineScript

    '
    ' Append the Usage script
    '
    DriverUsageScript

    '
    ' Append the output macro
    '
    EchoLineScript

end sub

'
' Abstract for the driver cloning script
'
sub DriverAbstractScript

    ws("' Abstract:")
    ws("'")
    ws("' " & strScriptName & " - driver cloning script for Windows 2000")
    ws("'")
    oScript.WriteLine(kLongLineStr)

end sub

'
' The Usage script used in the driver cloning script
'
sub DriverUsageScript

    blank
    ws("'")
    ws("' Display command usage.")
    ws("'")
    ws("sub Usage(ByVal bExit)")
    blank

    ws("    EchoLine kNormal, ""Usage: " & strScriptName & " [-c Destination_Server] [-v]"" ")
    ws("    EchoLine kNormal, ""Arguments:"" ")
    ws("    EchoLine kNormal, ""    -c   - destination server name"" ")
    ws("    EchoLine kNormal, ""    -v   - verbose mode"" ")
    ws("    EchoLine kNormal, ""    -?   - display command usage"" ")
    ws("    EchoLine kNormal, """" ")
    ws("    EchoLine kNormal, ""Examples:"" ")
    ws("    EchoLine kNormal, ""    " & strScriptName & """ ")
    blank

    ws("    if bExit then")
    ws("        wscript.quit(1)")
    ws("    end if")
    blank

    ws("end sub")
    blank

end sub

'
'------------------------------------------------------------------------
'      Port cloning script
'------------------------------------------------------------------------
'
sub PortCloneScript(ByVal strScriptName, ByVal strServerName)

    on error resume next

    wscript.echo
    wscript.echo "Creating the port cloning script..."

    '
    ' Open the script file
    '
    set oScript = oFileSystem.CreateTextFile(strScriptName,TRUE)

    PortStartUp

    ws("    EchoLine kNormal, """" ")
    ws("    EchoLine kNormal, ""------------------------------"" ")
    ws("    EchoLine kNormal, ""Start installing ports..."" ")

    '
    ' Enumerate all the ports in server "strServerName", for each port found,
    ' add a line in the script to call the AddLocalPort subroutine
    '
    dim oMaster
    dim oPort
    dim iPortCount

    iPortCount = 0
    set oMaster  = CreateObject("PrintMaster.PrintMaster.1")
    for each oPort in oMaster.Ports(strServerName)

        if Err = 0 then

           if oPort.PortType = kLprMon or oPort.PortType = kHPdlc then

               '
               ' Skip these ports because PrnAdmin cannot add them
               '
               wscript.echo "Skipping port " & oPort.PortName & " (" & Description(oPort.PortType) & ")"

           else
               '
               '  Duplicate only local ports different from LPTx:, COMx:
               '
               if oPort.PortType         =  kLocal                           and _
                  bFindPortPattern(kComPortPattern, oPort.PortName) = false  and _
                  bFindPortPattern(kLptPortPattern, oPort.PortName) = false  then

                   '
                   ' First try deleting the existing port
                   '
                   ws("    DeletePort      strDestServer, _")
                   ws("                    """ & StuffQuote(oPort.PortName) & """")

                   '
                   ' Add the local port
                   '
                   iPortCount = iPortCount + 1
                   ws("    AddLocalPort    strDestServer, _")
                   ws("                    """ & StuffQuote(oPort.PortName)  & """")
                   blank

               else

                   '
                   ' Otherwise, clone Standard TCP/IP ports
                   '
                   if oPort.PortType = kTcpRaw      or _
                      oPort.PortType = kTcpLpr      or _
                      oPort.Description = kStdTCPIP then

                       '
                       ' Get the configuration of this TCP port
                       '
                       dim strPortNameBackup

                       strPortNameBackup = oPort.PortName

                       oMaster.PortGet strServerName, strPortNameBackup , oPort

                       if Err = 0 then

                           '
                           ' First try deleting the existing port
                           '
                           ws("    DeletePort      strDestServer, _")
                           ws("                    """ & StuffQuote(oPort.PortName) & """")

                           '
                           ' Add this standard TCP/IP port
                           '
                           iPortCount = iPortCount + 1

                           if oPort.PortType = kTcpRaw then

                               '
                               ' Add a call to "AddTCPRawPort"
                               '
                               ws("    AddTCPRawPort   strDestServer, _")
                               ws("                    """ & StuffQuote(oPort.PortName)      & """,  _")
                               ws("                    """ & oPort.HostAddress               & """,  _")
                               ws("                    "   & CStr(oPort.PortNumber)          & ",  _")
                               ws("                    "   & BoolStr(oPort.SNMP)             & ",  _")
                               ws("                    "   & CStr(oPort.SNMPDeviceIndex)     & ",  _")
                               ws("                    """ & oPort.CommunityName             & """")
                               blank

                           else

                               '
                               ' Add a call to "AddTCPLprPort"
                               '
                               ws("    AddTCPLprPort   strDestServer, _")
                               ws("                    """ & StuffQuote(oPort.PortName)      & """,  _")
                               ws("                    """ & oPort.HostAddress               & """,  _")
                               ws("                    "   & CStr(oPort.PortNumber)          & ",    _")
                               ws("                    """ & oPort.QueueName                 & """,  _")
                               ws("                    "   & BoolStr(oPort.SNMP)             & ",  _")
                               ws("                    "   & CStr(oPort.SNMPDeviceIndex)     & ",  _")
                               ws("                    """ & oPort.CommunityName             & """,  _")
                               ws("                    "   & BoolStr(oPort.DoubleSpool)      & "")
                               blank

                            end if

                        else

                            wscript.echo "Error getting configuration for port " & oPort.PortName & " (Standard TCP Port)"

                            Err.Clear

                        end if

                   else

                       wscript.echo "Skipping port " & oPort.PortName & " (" & Description(oPort.PortType) & ")"

                   end if

               end if

           end if

        else

            '
            ' Clean up
            '
            oScript.Close

            oFileSystem.DeleteFile strScriptName

            wscript.echo "Error: Listing ports, error: 0x" & Hex(Err.Number)
            if Err.Description <> "" then
                wscript.echo "Error description: " & Err.Description
            end if

            exit sub

        end if

    next

    if Err = 0 then

        wscript.echo "Success: Listing ports on Server " & strServerName

    else

        wscript.echo "Error: Listing ports, error: 0x" & Hex(Err.Number)
        if Err.Description <> "" then
            wscript.echo "Error description: " & Err.Description
        end if

        Err.Clear

    end if

    '
    ' Write the summary script
    '
    wscript.echo "A total of " & CStr(iPortCount) & " port(s) are listed."

    ws("    EchoLine kNormal, ""Attempted to install a total of "" & CStr(iPortCount) & "" port(s)."" ")
    ws("    EchoLine kNormal, CStr(iSuccessCount) & "" port(s) successfully installed,"" ")
    ws("    EchoLine kNormal, CStr(iExistCount) & "" port(s) already exist."" ")

    '
    ' Append other functions of port cloning script
    '
    PortCleanUp

    '
    ' Close the script file
    '
    oScript.Close

    wscript.echo "The script file for cloning ports is """ & strScriptName & """."

end sub

'
' Subroutine of "AddLocalPort"
'
sub ScriptAddLocalPort

    '
    ' Insert the comment line before the function header
    '
    blank
    ws("'")
    ws("' Add a local port")
    ws("'")

    '
    ' The function header
    '
    ws("sub AddLocalPort(ByVal strServerName,         _")
    ws("                 ByVal strPortName            _")
    ws(")")
    blank

    '
    ' The function body
    '
    ws("    on error resume next")
    blank

    '
    ' Print out the information about the port that is about to be installed
    '
    ws("    iPortCount = iPortCount + 1")
    ws("    EchoLine kVerbose, ""Port:"" & CStr(iPortCount) ")
    ws("    EchoLine kVerbose, ""    ServerName         : "" & strServerName ")
    ws("    EchoLine kVerbose, ""    PortName           : "" & strPortName ")
    ws("    EchoLine kVerbose, ""    PortType           : "  & Description(kLocal) & """")
    blank

    ws("    dim oMaster")
    ws("    dim oPort")
    blank

    ws("    set oMaster = CreateObject(""PrintMaster.PrintMaster.1"")")
    ws("    set oPort   = CreateObject(""Port.Port.1"")")
    blank

    ws("    oPort.ServerName         = strServerName")
    ws("    oPort.PortName           = strPortName")
    ws("    oPort.PortType           = kLocal" )
    blank

    ws("    oMaster.PortAdd oPort")
    blank

    ws("    if Err = 0 or Err.Number = &H800700B7 then")
    blank

    ws("        if Err = 0 then")
    ws("            EchoLine kVerbose, ""Success: Port "" & strPortName & "" (" & Description(kTcpRaw) & ") added to server "" & strServerName ")
    ws("            iSuccessCount = iSuccessCount + 1")
    ws("        else")
    ws("            EchoLine kVerbose, ""Port "" & strPortName & "" (" & Description(kTcpRaw) & ") already exists on server "" & strServerName ")
    ws("            iExistCount = iExistCount + 1")
    ws("        end if")
    blank

    ws("    else")
    blank

    ws("        EchoLine kNormal, ""Error: adding port "" & strPortName & "" (" & Description(kLocal) & "), error: 0x"" & Hex(Err.Number) ")
    ws("        if Err.Description <> """" then ")
    ws("            EchoLine kNormal,  ""       Error description: "" & Err.Description ")
    ws("        end if")
    ws("        Err.Clear")
    blank


    ws("    end if")
    blank

    ws("    EchoLine kVerbose, """"")
    blank

    ws("end sub")
    blank

end sub

'
' Subroutine of "AddTCPRawPort"
'
sub ScriptAddTCPRawPort

    '
    ' Insert the comment line before the function header
    '
    blank
    ws("'")
    ws("' Add a Tcp raw port")
    ws("'")

    '
    ' The function header
    '
    ws("sub AddTCPRawPort(ByVal strServerName,         _")
    ws("                  ByVal strPortName,           _")
    ws("                  ByVal strHostAddress,        _")
    ws("                  ByVal PortNumber,            _")
    ws("                  ByVal SNMP,                  _")
    ws("                  ByVal SNMPDeviceIndex,       _")
    ws("                  ByVal CommunityName          _")
    ws(")")
    blank

    '
    ' The function body
    '
    ws("    on error resume next")
    blank

    '
    ' Print out the information about the port that is about to be installed
    '
    ws("    iPortCount = iPortCount + 1")
    ws("    EchoLine kVerbose, ""Port:"" & CStr(iPortCount) ")
    ws("    EchoLine kVerbose, ""    ServerName         : "" & strServerName ")
    ws("    EchoLine kVerbose, ""    PortName           : "" & strPortName ")
    ws("    EchoLine kVerbose, ""    PortType           : "  & Description(kTcpRaw) & """")
    ws("    EchoLine kVerbose, ""    HostAddress        : "" & strHostAddress ")
    ws("    EchoLine kVerbose, ""    PortNumber         : "" & CStr(PortNumber) ")
    ws("    EchoLine kVerbose, ""    SNMP               : "" & BoolStr(SNMP) ")
    ws("    EchoLine kVerbose, ""    SNMPDeviceIndex    : "" & CStr(SNMPDeviceIndex) ")
    ws("    EchoLine kVerbose, ""    CommunityName      : "" & CommunityName ")
    blank

    '
    ' The code that installs the port
    '
    ws("    dim oMaster")
    ws("    dim oPort")
    blank

    ws("    set oMaster = CreateObject(""PrintMaster.PrintMaster.1"")")
    ws("    set oPort   = CreateObject(""Port.Port.1"")")
    blank

    ws("    oPort.ServerName         = strServerName")
    ws("    oPort.PortName           = strPortName")
    ws("    oPort.PortType           = kTcpRaw" )
    ws("    oPort.HostAddress        = strHostAddress")
    ws("    oPort.PortNumber         = PortNumber")
    ws("    oPort.SNMP               = SNMP")
    ws("    oPort.SNMPDeviceIndex    = SNMPDeviceIndex")
    ws("    oPort.CommunityName      = CommunityName")
    blank

    ws("    oMaster.PortAdd oPort")
    blank

    ws("    if Err = 0 or Err.Number = &H80070034  then")
    blank

    ws("        if Err = 0 then")
    ws("            EchoLine kVerbose, ""Success: Port "" & strPortName & "" (" & Description(kTcpRaw) & ") added to server "" & strServerName ")
    ws("            iSuccessCount = iSuccessCount + 1")
    ws("        else")
    ws("            EchoLine kVerbose, ""Port "" & strPortName & "" (" & Description(kTcpRaw) & ") already exists on server "" & strServerName ")
    ws("            iExistCount = iExistCount + 1")
    ws("        end if")
    blank

    ws("    else")
    blank

    ws("        EchoLine kNormal, ""Error: adding port "" & strPortName & "" (" & Description(kTcpRaw) & "), error: 0x"" & hex(Err.Number) ")
    ws("        if Err.Description <> """" then ")
    ws("            EchoLine kNormal,  ""       Error description: "" & Err.Description ")
    ws("        end if")
    ws("        Err.Clear")
    blank


    ws("    end if")
    blank

    ws("    EchoLine kVerbose, """"")
    blank

    ws("end sub")
    blank

end sub

'
' Subroutine of "AddTCPLprPort"
'
sub ScriptAddTCPLprPort

    '
    ' Insert the comment line before the function header
    '
    blank
    ws("'")
    ws("' Add a Tcp lpr port")
    ws("'")

    '
    ' The function header
    '
    ws("sub AddTCPLprPort(ByVal strServerName,         _")
    ws("                  ByVal strPortName,           _")
    ws("                  ByVal strHostAddress,        _")
    ws("                  ByVal PortNumber,            _")
    ws("                  ByVal QueueName,             _")
    ws("                  ByVal SNMP,                  _")
    ws("                  ByVal SNMPDeviceIndex,       _")
    ws("                  ByVal CommunityName,         _")
    ws("                  ByVal DoubleSpool            _")
    ws(")")
    blank

    '
    ' The function body
    '
    ws("    on error resume next")
    blank

    '
    ' Print out the information about the port that is about to be installed
    '
    ws("    iPortCount = iPortCount + 1")
    ws("    EchoLine kVerbose, ""Port:"" & CStr(iPortCount) ")
    ws("    EchoLine kVerbose, ""    ServerName         : "" & strServerName ")
    ws("    EchoLine kVerbose, ""    PortName           : "" & strPortName ")
    ws("    EchoLine kVerbose, ""    PortType           : "  & Description(kTcpLpr) & """")
    ws("    EchoLine kVerbose, ""    HostAddress        : "" & strHostAddress ")
    ws("    EchoLine kVerbose, ""    PortNumber         : "" & CStr(PortNumber) ")
    ws("    EchoLine kVerbose, ""    QueueName          : "" &  QueueName ")
    ws("    EchoLine kVerbose, ""    SNMP               : "" & BoolStr(SNMP) ")
    ws("    EchoLine kVerbose, ""    SNMPDeviceIndex    : "" & CStr(SNMPDeviceIndex) ")
    ws("    EchoLine kVerbose, ""    CommunityName      : "" & CommunityName ")
    ws("    EchoLine kVerbose, ""    DoubleSpool        : "" & BoolStr(DoubleSpool) ")
    blank

    '
    ' The code that installs the port
    '
    ws("    dim oMaster")
    ws("    dim oPort")
    blank

    ws("    set oMaster = CreateObject(""PrintMaster.PrintMaster.1"")")
    ws("    set oPort   = CreateObject(""Port.Port.1"")")
    blank

    ws("    oPort.ServerName         = strServerName")
    ws("    oPort.PortName           = strPortName")
    ws("    oPort.PortType           = kTcpLpr" )
    ws("    oPort.HostAddress        = strHostAddress")
    ws("    oPort.PortNumber         = PortNumber")
    ws("    oPort.QueueName          = QueueName")
    ws("    oPort.SNMP               = SNMP")
    ws("    oPort.SNMPDeviceIndex    = SNMPDeviceIndex")
    ws("    oPort.CommunityName      = CommunityName")
    ws("    oPort.DoubleSpool        = DoubleSpool")
    blank

    ws("    oMaster.PortAdd oPort")
    blank

    ws("    if Err = 0 or Err.Number = &H80070034 then")
    blank

    ws("        if Err = 0 then")
    ws("            EchoLine kVerbose, ""Success: Port "" & strPortName & "" (" & Description(kTcpRaw) & ") added to server "" & strServerName ")
    ws("            iSuccessCount = iSuccessCount + 1")
    ws("        else")
    ws("            EchoLine kVerbose, ""Port "" & strPortName & "" (" & Description(kTcpRaw) & ") already exists on server "" & strServerName ")
    ws("            iExistCount = iExistCount + 1")
    ws("        end if")
    blank

    ws("    else")
    blank

    ws("        EchoLine kNormal, ""Error: adding port "" & strPortName & "" (" & Description(kTcpLpr) & "), error: 0x"" & Hex(Err.Number) ")
    ws("        if Err.Description <> """" then ")
    ws("            EchoLine kVerbose,  ""       Error description: "" & Err.Description ")
    ws("        end if")
    ws("        Err.Clear")
    blank


    ws("    end if")
    blank

    ws("    EchoLine kVerbose, """"")
    blank

    ws("end sub")
    blank

end sub

'
' Subroutine of "DeletePort"
'
sub ScriptDeletePort

    '
    ' Insert the comment line before the function header
    '
    blank
    ws("'")
    ws("' Delete an existing port")
    ws("'")

    '
    ' The function header
    '
    ws("sub DeletePort(ByVal strServerName,         _")
    ws("               ByVal strPortName            _")
    ws(")")
    blank

    '
    ' The function body
    '
    ws("    on error resume next")
    blank

    '
    ' If the user asks for keeping the original port, then don't delete it
    '
    ws("    if bKeepOriginalOnes = true then")
    blank
    ws("        exit sub")
    blank
    ws("    end if")
    blank

    '
    ' Print out the information about the port that is about to be deleted
    '
    ws("    EchoLine kVerbose, ""  Deleting Port: """)
    ws("    EchoLine kVerbose, ""    ServerName         : "" & strServerName ")
    ws("    EchoLine kVerbose, ""    PortName           : "" & strPortName ")
    blank

    '
    ' The code that deletes the port
    '
    ws("    dim oMaster")
    ws("    dim oPort")
    blank

    ws("    set oMaster = CreateObject(""PrintMaster.PrintMaster.1"")")
    ws("    set oPort   = CreateObject(""Port.Port.1"")")
    blank

    ws("    oPort.ServerName = strServerName")
    ws("    oPort.PortName   = strPortName")
    blank

    ws("    oMaster.PortDel oPort")
    blank

    ws("    if Err = 0 then")
    blank

    ws("        EchoLine kVerbose, ""  Success: Delete Port"" ")
    blank

    ws("    else")
    blank

    ws("        EchoLine kVerbose, ""  Error deleting port. Error: 0x"" & hex(Err.Number)")
    ws("        if Err.Description <> """" then ")
    ws("            EchoLine kVerbose,  ""       Error description: "" & Err.Description ")
    ws("        end if")
    ws("        Err.Clear")
    blank

    ws("    end if")
    blank

    ws("    EchoLine kVerbose, """"")
    blank

    ws("end sub")
    blank

end sub

'
' StartUp script for cloning ports
'
sub PortStartUp

    '
    ' Start creating the port cloning script
    '
    CopyrightScript
    PortAbstractScript

    '
    ' The script program starts
    '
    blank
    ws("option explicit")
    blank

    ws("'")
    ws("' Verbose Level")
    ws("'")
    ws("const kNormal    = 0")
    ws("const kVerbose   = 1")
    blank

    ws("'")
    ws("' Port Types")
    ws("'")
    ws("const kTcpRaw   = 1")
    ws("const kTcpLpr   = 2")
    ws("const kLocal    = 3")
    ws("const kLprMon   = 5")
    ws("const kHPdlc    = 7")
    ws("const kUnknown  = 8")
    blank

    ws("'")
    ws("' Flag, set if users don't want to replace the old ports")
    ws("'")
    ws("dim bKeepOriginalOnes")
    blank

    ws("dim strDestServer")
    blank

    ws("' The number of ports to be installed")
    blank
    ws("dim iPortCount")
    blank
    ws("' The number of ports sucessfully installed or that already exist")
    blank
    ws("dim iSuccessCount")
    ws("dim iExistCount")
    blank
    ws("dim bVerbose")
    blank

    ws("main")
    blank
    ws("'")
    ws("' Main execution starts here")
    ws("'")

    ws("sub main")
    blank
    ws("    bVerbose = false")
    ws("    bKeepOriginalOnes=false")
    ws("    iPortCount  = 0")
    ws("    iSuccessCount = 0")
    ws("    iExistCount = 0")
    ws("    strDestServer = """"")
    ws("    ParseCommandLine")
    blank

end sub

'
' CleanUp script for cloning ports
'
sub PortCleanUp

    ws("end sub")

    '
    ' Append the subroutine "AddLocalPort"
    '
    ScriptAddLocalPort

    '
    ' Append the subroutine "AddTCPRawPort"
    '
    ScriptAddTCPRawPort

    '
    ' Append the subroutine "AddTCPLprPort"
    '
    ScriptAddTCPLprPort

    '
    ' Append the subroutine "DeletePort"
    '
    ScriptDeletePort

    '
    ' Append the function "BoolStr"
    '
    BoolStrScript

    '
    ' Append the command line parsing script
    '
    ParseCommandLineScript

    '
    ' Append the Usage script
    '
    PortUsageScript

    '
    ' Append the output macro
    '
    EchoLineScript

end sub

'
' Abstract for the port cloning script
'
sub PortAbstractScript

    ws("' Abstract:")
    ws("'")
    ws("' " & strScriptName & " - port cloning script for Windows 2000")
    ws("'")
    oScript.WriteLine(kLongLineStr)

end sub

'
' The Usage script used in the port cloning script
'
sub PortUsageScript

    blank
    ws("'")
    ws("' Display command usage.")
    ws("'")
    ws("sub Usage(ByVal bExit)")
    blank

    ws("    EchoLine kNormal, ""Usage: " & strScriptName & " [-c Destination_Server] [-kv]"" ")
    ws("    EchoLine kNormal, ""Arguments:"" ")
    ws("    EchoLine kNormal, ""    -c   - destination server name"" ")
    ws("    EchoLine kNormal, ""    -k   - keep the existing port with the same name"" ")
    ws("    EchoLine kNormal, ""           (WARNING: If this flag is set, the cloned server may not be identical to the original one."" ")
    ws("    EchoLine kNormal, ""    -v   - verbose mode"" ")
    ws("    EchoLine kNormal, ""    -?   - display command usage"" ")
    ws("    EchoLine kNormal, """" ")
    ws("    EchoLine kNormal, ""Examples:"" ")
    ws("    EchoLine kNormal, ""    " & strScriptName & """ ")
    blank

    ws("    if bExit then")
    ws("        wscript.quit(1)")
    ws("    end if")
    blank

    ws("end sub")
    blank

end sub

'
' Gets a string description for a port type
'
function Description(ByVal value)

    select case value

           case kTcpRaw

                Description = "TCP RAW Port"

           case kTcpLpr

                Description = "TCP LPR Port"

           case kLocal

                Description = "Standard Local Port"

           case kLprMon

                Description = "LPR Mon Port"

           case kHPdlc

                Description = "HP DLC Port"

           case kUnknown

                Description = "Unknown Port"

           case Else

                Description = "Invalid PortType"
    end select

end function

'
'------------------------------------------------------------------------
'      Printer cloning script
'------------------------------------------------------------------------
'
sub PrinterCloneScript(ByVal strScriptName, ByVal strServerName)

    on error resume next

    wscript.echo
    wscript.echo "Creating the printer cloning script..."

    dim strPersistFilename

    '
    ' Open the script file
    '
    set oScript = oFileSystem.CreateTextFile(strScriptName,TRUE)

    PrinterStartUp

    ws("    EchoLine kNormal, """" ")
    ws("    EchoLine kNormal, ""------------------------------"" ")
    ws("    EchoLine kNormal, ""Start installing printers..."" ")

    '
    ' Enumerate all the printers in server "strServerName",
    ' for each printer found, add a line in the script to
    ' call the AddPrinter subroutine
    '
    dim oMaster
    dim oPrinter
    dim iPrinterCount
    dim strPrinterName

    iPrinterCount = 0
    set oMaster  = CreateObject("PrintMaster.PrintMaster.1")
    for each oPrinter in oMaster.Printers(strServerName)

        if Err = 0 then
           '
           ' Check to see if we need to clone this printer
           ' We only clone local printers
           '

           if bIsLocal(strServerName, oPrinter.PrinterName) then

               '
               ' Remove the ServerName part of the PrinterName
               '
               strPrinterName = Strip(oPrinter.PrinterName)

               '
               ' Try deleting the existing printer
               '
               ws("    DeletePrinter  strDestServer,  _")
               ws("                   """ & StuffQuote(strPrinterName) & """")

               '
               ' Add a call to "AddPrinter"
               '
               iPrinterCount = iPrinterCount + 1
               ws("    AddPrinter  strDestServer, _")
               ws("                """ & StuffQuote(strPrinterName)  & """,  _")
               ws("                """ & oPrinter.DriverName   & """,  _")
               ws("                """ & StuffQuote(oPrinter.PortName)  & """,  _")
               ws("                """",  _")               ' Use default DriverPath
               ws("                """"    ")               ' Use default InfFile

               '
               ' Persist save the printer
               '
               strPersistFilename = strPrefixServerName & "per" & CStr(iPrinterCount) & "_clone.vbs"
               if SavePrinter( oPrinter.PrinterName, strPersistFilename ) = false then

                   wscript.echo "Error: skipping printer """ & oPrinter.PrinterName & """ persist save due to the error when getting the persist information"

               else

                   '
                   ' Add script for calling persist restore
                   '
                   ws("    RestorePrinter strDestServer, """ & strPrinterName & """, """ & strPersistFilename & """" )
                   wscript.echo "Printer Name: """ &  strPrinterName  & """    <==>    Persist filename: " & strPersistFilename

               end if

               blank

           else

               '
               ' It is not a local printer. Don't clone it.
               '
               wscript.echo "Skipping non-local printer: """ & oPrinter.PrinterName & """."

           end if

        else

            '
            ' Clean up
            '
            oScript.Close

            oFileSystem.DeleteFile strScriptName

            wscript.echo "Error: Listing printers, error: 0x" & Hex(Err.Number)
            if Err.Description <> "" then
               wscript.echo "Error description: " & Err.Description
            end if

            exit sub

        end if

    next

    if Err = 0 then

        wscript.echo "Success: Listing printers on Server " & strServerName

    else

        wscript.echo "Error: Listing printers, error: 0x" & Hex(Err.Number)
        if Err.Description <> "" then
            wscript.echo "Error description: " & Err.Description
        end if

        Err.Clear

    end if

    '
    ' Write the summary
    '
    wscript.echo "A total of " & CSTR(iPrinterCount) & " printers are listed."

    ws("    EchoLine kNormal, ""Attempted to install a total of "" & CStr(iPrinterCount) & "" printer(s)."" ")
    ws("    EchoLine kNormal, CStr(iSuccessCount) & "" printer(s) were successfully installed."" ")


    if strServerName = "" then

        '
        ' If the source server is local, get and set the default printer
        '
        dim strDefaultPrinterName
        set oMaster = CreateObject("PrintMaster.PrintMaster.1")
        strDefaultPrinterName = StuffQuote(oMaster.DefaultPrinter)

        if Err <> 0 then

            wscript.echo "Error: Getting default printer """ & oMaster.DefaultPrinter & """, error: 0x" & Hex(Err.Number)
            if Err.Description <> "" then
                wscript.echo "Error description: " & Err.Description
            end if

            Err.Clear

        else

            '
            ' Setting the default printer
            '
            blank
            ws("'")
            ws("' Set the default printer")
            ws("'   (do this only if the installation is on the local machine)")
            ws("'")
            ws("    if strDestServer = """" then ")
            blank
            ws("        dim oMaster")
            ws("        set oMaster = CreateObject(""PrintMaster.PrintMaster.1"") ")
            ws("        oMaster.DefaultPrinter = """ & strDefaultPrinterName & """" )
            blank
            ws("        if Err = 0 then")
            ws("            EchoLine kNormal, ""Success: Setting the default printer to """"" & strDefaultPrinterName & """"" "" " )
            ws("        else")
            ws("            EchoLine kNormal, ""Error: Setting default printer "" &  strDefaultPrinterName ")
            ws("        end if")
            blank
            ws("    end if")
            blank

        end if

    end if

    PrinterCleanUp

    '
    ' Close the script file
    '
    oScript.Close

    wscript.echo "The script file for cloning printers is """ & strScriptName & """."

end sub

'
' Subroutine of "AddPrinter"
'
sub ScriptAddPrinter

    '
    ' Insert the comment line before the function header
    '
    blank
    ws("'")
    ws("' Add a printer")
    ws("'")

    '
    ' The function header
    '
    ws("sub AddPrinter(ByVal strServerName,       _")
    ws("               ByVal strPrinterName,      _")
    ws("               ByVal strDriverName,       _")
    ws("               ByVal strPortName,         _")
    ws("               ByVal strDriverPath,       _")
    ws("               ByVal strInfFile           _")
    ws(")")
    blank

    '
    ' The function body
    '
    ws("    on error resume next")
    blank

    '
    ' Print out the information about the printer that is about to be installed
    '
    ws("    iPrinterCount = iPrinterCount + 1")
    ws("    EchoLine kVerbose, ""Printer:"" & CSTR(iPrinterCount)                  ")
    ws("    EchoLine kVerbose, ""    ServerName      : "" & strServerName          ")
    ws("    EchoLine kVerbose, ""    PrinterName     : "" & strPrinterName         ")
    ws("    EchoLine kVerbose, ""    DriverName      : "" & strDriverName          ")
    ws("    EchoLine kVerbose, ""    PortName        : "" & strPortName            ")
    blank

    '
    ' The codes that installs the printer
    '
    ws("    dim oMaster")
    ws("    dim oPrinter")
    blank

    ws("    set oMaster = CreateObject(""PrintMaster.PrintMaster.1"")")
    ws("    set oPrinter = CreateObject(""Printer.Printer.1"")")
    blank

    ws("    oPrinter.ServerName       = strServerName     ")
    ws("    oPrinter.PrinterName      = strPrinterName    ")
    ws("    oPrinter.DriverName       = strDriverName     ")
    ws("    oPrinter.PortName         = strPortName       ")
    ws("    oPrinter.DriverPath       = strDriverPath     ")
    ws("    oPrinter.InfFile          = strInfFile        ")
    blank

    ws("    oMaster.PrinterAdd oPrinter")
    blank

    ws("    if Err = 0 then")
    blank

    ws("        EchoLine kVerbose, ""Success: Printer "" & strPrinterName & "" added to server "" & strServerName ")
    ws("        iSuccessCount = iSuccessCount + 1")
    blank

    ws("    else")
    blank

    ws("        EchoLine kNormal, ""Error adding printer "" & strPrinterName & "", error: 0x"" & Hex(Err.Number) ")
    ws("        if Err.Description <> """" then ")
    ws("            EchoLine kVerbose,  ""       Error description: "" & Err.Description ")
    ws("        end if")
    ws("        Err.Clear")
    blank


    ws("    end if")
    blank

    ws("    EchoLine kVerbose, """"")
    blank

    ws("end sub")
    blank

end sub

'
' Save printer configuration
'
function SavePrinter(ByVal strPrinterName, ByVal strFileName)

    on error resume next

    dim oMaster
    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    oMaster.PrinterPersistSave strPrinterName, strFileName, kAllSettings

    if Err <> 0 then

        wscript.echo "Error saving the configuration of the printer """ & strPrinterName & """, error: 0x" & Hex(Err.Number)

        SavePrinter = false

    else

        SavePrinter = true

    end if

end function

'
' Script for printer persist restore
'
sub RestorePrinterScript()

    ws("'")
    ws("' Restore printer configuration")
    ws("'")
    blank

    ws("sub RestorePrinter(ByVal strServerName, ByVal strPrinterName, ByVal strFileName)")
    blank

    ws("    on error resume next")
    blank

    ws("    dim oMaster")
    ws("    set oMaster = CreateObject(""PrintMaster.PrintMaster.1"") ")
    blank

    ws("    oMaster.PrinterPersistRestore strFullName(strServerName, strPrinterName), strFileName, _")
    ws("      kAllSettings + kResolveName + kReslovePort + kResolveShare")
    blank

    ws("    if Err = 0 then")
    ws("        EchoLine kVerbose, ""Success restoring the configuration of the printer"" & strPrinterName ")
    ws("        EchoLine kVerbose, """" ")
    ws("    else")
    ws("        if Err.Number = kErrorNoDs then")
    ws("            Err.Clear")
    ws("            '")
    ws("            ' Try resoring without Printer Info 7")
    ws("            '")
    ws("           oMaster.PrinterPersistRestore strFullName(strServerName, strPrinterName), strFileName, _")
    ws("           kPersistNoDs + kResolveName + kReslovePort + kResolveShare")
    ws("        end if")
    ws("")
    ws("        if Err = 0 then")
    ws("            EchoLine kVerbose, ""Success restoring the configuration of the printer"" & strPrinterName ")
    ws("            EchoLine kVerbose, """" ")
    ws("        else")
    ws("            EchoLine kNormal, ""Error restoring the configuration of the printer "" & strPrinterName & "", error: 0x"" & Hex(Err.Number)")
    ws("            EchoLine kNormal, """" ")
    ws("            Err.Clear")
    ws("        end if")
    ws("    end if")
    blank

    ws("end sub")
    blank

end sub

'
' Subroutine of "DeletePrinter"
'
sub ScriptDeletePrinter

    '
    ' Insert the comment line before the function header
    '
    blank
    ws("'")
    ws("' Delete an existing printer")
    ws("'")

    '
    ' The function header
    '
    ws("sub DeletePrinter(ByVal strServerName,         _")
    ws("                  ByVal strPrinterName         _")
    ws(")")
    blank

    '
    ' The function body
    '
    ws("    on error resume next")
    blank

    '
    ' If the user asks for keeping the original printer, then don't delete it
    '
    ws("    if bKeepOriginalOnes = true then")
    blank
    ws("        exit sub")
    blank
    ws("    end if")
    blank

    '
    ' Print out the information about the printer that is about to be deleted
    '
    ws("    EchoLine kVerbose, ""  Deleting Printer: """)
    ws("    EchoLine kVerbose, ""    ServerName         : "" & strServerName ")
    ws("    EchoLine kVerbose, ""    PrinterName        : "" & strPrinterName ")
    blank

    '
    ' The code that deletes the printer
    '
    ws("    dim oMaster")
    ws("    dim oPrinter")
    blank

    ws("    set oMaster = CreateObject(""PrintMaster.PrintMaster.1"")")
    ws("    set oPrinter = CreateObject(""Printer.Printer.1"")")
    blank

    ws("    oPrinter.ServerName       = strServerName     ")
    ws("    oPrinter.PrinterName      = strPrinterName    ")
    blank

    ws("    oMaster.PrinterDel oPrinter")
    blank

    ws("    if Err = 0 then")
    blank

    ws("        EchoLine kVerbose, ""  Success: Delete Printer"" ")
    blank

    ws("    else")
    blank

    ws("        EchoLine kVerbose, ""  Error deleting printer "" & strPrinterName & "". Error: 0x"" & Hex(Err.Number)")
    ws("        if Err.Description <> """" then ")
    ws("            EchoLine kVerbose,  ""       Error description: "" & Err.Description ")
    ws("        end if")
    ws("        Err.Clear")
    blank

    ws("    end if")
    blank

    ws("    EchoLine kVerbose, """"")
    blank

    ws("end sub")
    blank

end sub

'
' StartUp script for cloning printers
'
sub PrinterStartUp

    '
    ' Start to create the printer cloning script
    '
    CopyrightScript
    PrinterAbstractScript

    '
    ' The script program starts
    '
    blank
    ws("option explicit")
    blank

    ws("'")
    ws("' Verbose Level")
    ws("'")
    ws("const kNormal    = 0")
    ws("const kVerbose   = 1")
    blank

    '
    '  Add the printer persist constants
    '
    ws("'")
    ws("' Constants for printer persist")
    ws("'")
    ws("const kAllSettings            = 127")
    ws("const kPrinterInfo7           = 4")
    ws("const kResolveName            = 256")
    ws("const kReslovePort            = 512")
    ws("const kResolveShare           = 1024")
    ws("'")
    ws("'")
    ws("' If the DS is not present, restore printer without Printer Info 7")
    ws("'")
    ws("const kPersistNoDs            = 123")
    ws("const kErrorNoDs              = &H8004000C")
    blank

    ws("'")
    ws("' Flag, set if users don't want to replace the old forms")
    ws("'")
    ws("dim bKeepOriginalOnes")
    blank

    ws("dim strDestServer")
    blank

    ws("' The number of printers to be installed")
    blank
    ws("dim iPrinterCount")
    blank
    ws("' The number of printers sucessfully installed")
    blank
    ws("dim iSuccessCount")
    blank
    ws("dim bVerbose")
    blank

    ws("main")
    blank
    ws("'")
    ws("' Main execution starts here")
    ws("'")

    ws("sub main")
    blank
    ws("    bVerbose = false")
    ws("    bKeepOriginalOnes=false")
    ws("    iPrinterCount  = 0")
    ws("    iSuccessCount = 0")
    ws("    strDestServer = """"")
    ws("    ParseCommandLine")
    blank

end sub

'
' CleanUp script for cloning printers
'
sub PrinterCleanUp

    ws("end sub")

    ' Append the subroutine "AddPrinter"
    ScriptAddPrinter

    '
    ' Append the subroutine "RestorePrinter"
    '
    RestorePrinterScript

    '
    ' Append the subroutine "DeletePrinter"
    '
    ScriptDeletePrinter

    '
    ' Append the script for creating the full name for the printer
    '
    strFullNameScript

    '
    ' Append the command line parsing script
    '
    ParseCommandLineScript

    '
    ' Append the Usage script
    '
    PrinterUsageScript

    '
    ' Append the output macro
    '
    EchoLineScript

end sub

'
' Script for creating the full printer name (containing ServerName and PrinterName)
' which is used in RestorePrinter
'
function strFullNameScript()

    ws("'")
    ws("' Function for creating the full printer name")
    ws("'")
    ws("function strFullName(ByVal strServerName, ByVal strPrinterName)")
    blank
    ws("    if strServerName = """" then")
    blank
    ws("        strFullName = strPrinterName")
    blank
    ws("    else")
    blank
    ws("        strFullName = strServerName & ""\"" & strPrinterName")
    blank
    ws("    end if")
    blank
    ws("end function")
    blank

end function

'
' Abstract for the printer cloning script
'
sub PrinterAbstractScript

    ws("' Abstract:")
    ws("'")
    ws("' " & strScriptName & " - printer cloning script for Windows 2000")
    ws("'")
    oScript.WriteLine(kLongLineStr)

end sub

'
' The Usage script used in the printer cloning script
'
sub PrinterUsageScript

    blank
    ws("'")
    ws("' Display command usage.")
    ws("'")
    ws("sub Usage(ByVal bExit)")
    blank

    ws("    EchoLine kNormal, ""Usage: " & strScriptName & " [-c Destination_Server] [-kv]"" ")
    ws("    EchoLine kNormal, ""Arguments:"" ")
    ws("    EchoLine kNormal, ""    -c   - destination server name"" ")
    ws("    EchoLine kNormal, ""    -k   - keep the existing printer with the same name"" ")
    ws("    EchoLine kNormal, ""           (WARNING: If this flag is set, the cloned server may not be identical to the original one."" ")
    ws("    EchoLine kNormal, ""    -v   - verbose mode"" ")
    ws("    EchoLine kNormal, ""    -?   - display command usage"" ")
    ws("    EchoLine kNormal, """" ")
    ws("    EchoLine kNormal, ""Examples:"" ")
    ws("    EchoLine kNormal, ""    " & strScriptName & """ ")
    blank

    ws("    if bExit then")
    ws("        wscript.quit(1)")
    ws("    end if")
    blank

    ws("end sub")
    blank

end sub

'
'------------------------------------------------------------------------
'      Form cloning script
'------------------------------------------------------------------------
'
sub FormCloneScript(ByVal strScriptName, ByVal strServerName)

    on error resume next

    wscript.echo
    wscript.echo   "Creating the form cloning script..."

    dim iHeight
    dim iWidth
    dim iTop
    dim iLeft
    dim iBottom
    dim iRight

    '
    ' Open the script file
    '
    set oScript = oFileSystem.CreateTextFile(strScriptName,TRUE)

    FormStartUp

    ws("    EchoLine kNormal, """" ")
    ws("    EchoLine kNormal, ""------------------------------"" ")
    ws("    EchoLine kNormal, ""Start installing forms..."" ")

    '
    ' Enumerate all the forms in server "strServerName",
    ' for each form found, add a line in the script to
    ' call the AddForm subroutine
    '
    dim oMaster
    dim oForm
    dim iFormCount

    iFormCount = 0
    set oMaster  = CreateObject("PrintMaster.PrintMaster.1")
    for each oForm in oMaster.Forms(strServerName)

        if Err = 0 then
           '
           ' Try deleting the existing form
           '
           ws("    DeleteForm  strDestServer,  _")
           ws("                """ & StuffQuote(oForm.Name) & """")

           oForm.GetSize iHeight, iWidth
           oForm.GetImageableArea  iTop, iLeft, iBottom, iRight

           iFormCount = iFormCount + 1
           ws("    AddForm     strDestServer, _")
           ws("                """ & StuffQuote(oForm.Name)  & """, _")
           ws("                "   & CStr(oForm.Flags)    & ", _")
           ws("                "   & CStr(iHeight)  & ", _")
           ws("                "   & CStr(iWidth)   & ", _")
           ws("                "   & CStr(iTop)     & ", _")
           ws("                "   & CStr(iLeft)    & ", _")
           ws("                "   & CStr(iBottom)  & ", _")
           ws("                "   & CStr(iRight)   )
           blank

        else

            '
            ' Clean up
            '
            oScript.Close

            oFileSystem.DeleteFile strScriptName

            wscript.echo "Error: Listing forms, error: 0x" & Hex(Err.Number)
            if Err.Description <> "" then
                wscript.echo "Error description: " & Err.Description
            end if

            exit sub

        end if

    next

    if Err = 0 then

        wscript.echo "Success: Listing forms on Server " & strServerName

    else

        wscript.echo "Error: Listing forms, error: 0x" & Hex(Err.Number)
        if Err.Description <> "" then
            wscript.echo "Error description: " & Err.Description
        end if

        Err.Clear

    end if

    wscript.echo  "A total of " & CSTR(iFormCount) & " forms are listed."

    ws("    EchoLine kNormal, ""Attempted to install a total of "" & CStr(iFormCount) & "" forms."" ")
    ws("    EchoLine kNormal, CStr(iSuccessCount) & "" forms successfully installed."" ")

    FormCleanUp

    '
    ' Close the script file
    '
    oScript.Close

    wscript.echo  "The script file for cloning forms is """ & strScriptName & """."

end sub

'
' Subroutine of "AddForm"
'
sub ScriptAddForm

    '
    ' Insert the comment line before the function header
    '
    blank
    ws("'")
    ws("' Add a Form")
    ws("'")

    '
    ' The function header
    '
    ws("sub AddForm(ByVal strServerName,         _")
    ws("            ByVal strName,               _")
    ws("            ByVal iFlags,                _")
    ws("            ByVal iHeight,               _")
    ws("            ByVal iWidth,                _")
    ws("            ByVal iTop,                  _")
    ws("            ByVal iLeft,                 _")
    ws("            ByVal iBottom,               _")
    ws("            ByVal iRight                 _")
    ws(")")
    blank

    '
    ' The function body
    '
    ws("    on error resume next")
    blank

    '
    ' Print out the information about the form that is about to be installed
    '
    ws("    iFormCount = iFormCount + 1")
    ws("    EchoLine kVerbose, ""Form:"" & CSTR(iFormCount) ")
    ws("    EchoLine kVerbose, ""    ServerName : "" & strServerName ")
    ws("    EchoLine kVerbose, ""    Name       : "" & strName ")
    ws("    EchoLine kVerbose, ""    Type       : "" & CStr(iFlags) ")
    ws("    EchoLine kVerbose, ""    Height     : "" & CStr(iHeight) ")
    ws("    EchoLine kVerbose, ""    Width      : "" & CStr(iWidth) ")
    ws("    EchoLine kVerbose, ""    Top        : "" & CStr(iTop) ")
    ws("    EchoLine kVerbose, ""    Left       : "" & CStr(iLeft) ")
    ws("    EchoLine kVerbose, ""    Bottom     : "" & CStr(iBottom) ")
    ws("    EchoLine kVerbose, ""    Right      : "" & CStr(iRight) ")
    blank

    '
    ' The code that installs the form
    '
    ws("    dim oMaster")
    ws("    dim oForm")
    blank

    ws("    set oMaster = CreateObject(""PrintMaster.PrintMaster.1"")")
    ws("    set oForm = CreateObject(""Form.Form.1"")")
    blank

    ws("    oForm.ServerName  = strServerName")
    ws("    oForm.Name        = strName")
    ws("    oForm.Flags       = iFlags")
    ws("    oForm.SetSize iHeight, iWidth")
    ws("    oForm.SetImageableArea iTop, iLeft, iBottom, iRight")
    blank

    ws("    oMaster.FormAdd oForm")
    blank

    ws("' If no error or error code is for ""existing form"" then succeed")
    ws("    if Err = 0 or Err.number = &H80070050 then")
    blank

    ws("        EchoLine kVerbose, ""Success: Form "" & strName & "" added to server "" & strServerName ")
    ws("        iSuccessCount = iSuccessCount + 1")
    blank

    ws("    else")
    blank

    ws("        EchoLine kNormal, ""Error: adding Form "" & strName & "", error: 0x"" & hex(Err.Number) ")
    ws("        if Err.Description <> """" then ")
    ws("            EchoLine kVerbose,  ""       Error description: "" & Err.Description ")
    ws("        end if")
    ws("        Err.Clear")
    blank


    ws("    end if")
    blank

    ws("    EchoLine kVerbose, """"")
    blank

    ws("end sub")
    blank

end sub

'
' Subroutine of "DeleteForm"
'
sub ScriptDeleteForm

    '
    ' Insert the comment line before the function header
    '
    blank
    ws("'")
    ws("' Delete an existing form")
    ws("'")

    '
    ' The function header
    '
    ws("sub DeleteForm(ByVal strServerName,         _")
    ws("               ByVal strFormName            _")
    ws(")")
    blank

    '
    ' The function body
    '
    ws("    on error resume next")
    blank

    '
    ' If the user asks for keeping the original form, then don't delete it
    '
    ws("    if bKeepOriginalOnes = true then")
    blank
    ws("        exit sub")
    blank
    ws("    end if")
    blank

    '
    ' Print out the information about the form that is about to be deleted
    '
    ws("    EchoLine kVerbose, ""  Deleting Form: """)
    ws("    EchoLine kVerbose, ""    ServerName         : "" & strServerName ")
    ws("    EchoLine kVerbose, ""    FormName           : "" & strFormName ")
    blank

    '
    ' The code that deletes the form
    '
    ws("    dim oMaster")
    ws("    dim oForm")
    blank

    ws("    set oMaster = CreateObject(""PrintMaster.PrintMaster.1"")")
    ws("    set oForm   = CreateObject(""Form.Form.1"")")
    blank

    ws("    oForm.Name       = strFormName")
    ws("    oForm.ServerName = strServerName")
    ws("    oMaster.FormDel  oForm")
    blank

    ws("    if Err = 0 then")
    blank

    ws("        EchoLine kVerbose, ""  Success: Delete Form"" & strFormName ")
    blank

    ws("    else")
    blank

    ws("        EchoLine kVerbose, ""  Error deleting form "" & strFormName & "". Error: 0x"" & hex(Err.Number)")
    ws("        if Err.Description <> """" then ")
    ws("            EchoLine kVerbose,  ""       Error description: "" & Err.Description ")
    ws("        end if")
    ws("        Err.Clear")
    blank

    ws("    end if")
    blank

    ws("    EchoLine kVerbose, """"")
    blank

    ws("end sub")
    blank

end sub

'
' StartUp script for cloning forms
'
sub FormStartUp

    '
    ' Start creating the form cloning script
    '
    CopyrightScript
    FormAbstractScript

    '
    ' The script program starts
    '
    blank
    ws("option explicit")
    blank

    ws("'")
    ws("' Verbose Level")
    ws("'")
    ws("const kNormal    = 0")
    ws("const kVerbose   = 1")
    blank

    ws("'")
    ws("' Flag, set if users don't want to replace the old forms")
    ws("'")
    ws("dim bKeepOriginalOnes")
    blank

    ws("dim strDestServer")
    blank

    ws("' The number of forms to be installed")
    blank
    ws("dim iFormCount")
    blank
    ws("' The number of forms sucessfully installed")
    blank
    ws("dim iSuccessCount")
    blank
    ws("dim bVerbose")
    blank

    ws("main")
    blank
    ws("'")
    ws("' Main execution starts here")
    ws("'")

    ws("sub main")
    blank
    ws("    bVerbose = false")
    ws("    bKeepOriginalOnes = false")
    ws("    iFormCount  = 0")
    ws("    iSuccessCount = 0")
    ws("    strDestServer = """"")
    ws("    ParseCommandLine")
    blank

end sub

'
' CleanUp script for cloning forms
'
sub FormCleanUp

    ws("end sub")

    '
    ' Append the subroutine "AddForm"
    '
    ScriptAddForm

    '
    ' Append the subroutine "DeleteForm"
    '
    ScriptDeleteForm

    '
    ' Append the commandline parsing script
    '
    ParseCommandLineScript

    '
    ' Append the Usage script
    '
    FormUsageScript

    '
    ' Append the output macro
    '
    EchoLineScript

end sub

'
' Abstract for the form cloning script
'
sub FormAbstractScript

    ws("' Abstract:")
    ws("'")
    ws("' " & strScriptName & " - Form cloning script for Windows 2000")
    ws("'")
    oScript.WriteLine(kLongLineStr)

end sub

'
' The Usage script used in the form cloning script
'
sub FormUsageScript

    blank
    ws("'")
    ws("' Display command usage.")
    ws("'")
    ws("sub Usage(ByVal bExit)")
    blank

    ws("    EchoLine kNormal, ""Usage: " & strScriptName & " [-c Destination_Server] [-kv]"" ")
    ws("    EchoLine kNormal, ""Arguments:"" ")
    ws("    EchoLine kNormal, ""    -c   - destination server name"" ")
    ws("    EchoLine kNormal, ""    -k   - keep the existing form that has the same name"" ")
    ws("    EchoLine kNormal, ""           (WARNING: If this flag is set, the cloned server may not be identical to the original one."" ")
    ws("    EchoLine kNormal, ""    -v   - verbose mode"" ")
    ws("    EchoLine kNormal, ""    -?   - display command usage"" ")
    ws("    EchoLine kNormal, """" ")
    ws("    EchoLine kNormal, ""Examples:"" ")
    ws("    EchoLine kNormal, ""    " & strScriptName & """ ")
    blank

    ws("    if bExit then")
    ws("        wscript.quit(1)")
    ws("    end if")
    blank

    ws("end sub")
    blank

end sub


'------------------------------------------------------------------------
'      Common scripts for the new generated cloning script
'------------------------------------------------------------------------

'
' Copyright header in the script
'
sub CopyrightScript()

    oScript.WriteLine(kLongLineStr)
    ws("'")
    ws("' Copyright (c) Microsoft Corporation 1999")
    ws("' All Rights Reserved")
    ws("'")

end sub

'
' The command line parsing script used in the generated cloning script
'
sub ParseCommandLineScript()

    blank
    ws("'")
    ws("' Command line parsing")
    ws("'")
    blank

    ws("sub ParseCommandLine()")
    blank

    ws("    dim oArgs")
    ws("    dim i")
    blank

    ws("    set oArgs = wscript.Arguments")
    blank

    ws("    while i < oArgs.Count")
    blank

    ws("       select case oArgs(i)")
    blank

    ws("           case ""-c"" ")
    ws("              i = i + 1")
    ws("              strDestServer = oArgs(i)")
    blank

    ws("           case ""-k"" ")
    ws("              bKeepOriginalOnes = true")
    blank

    ws("           case ""-v"" ")
    ws("              bVerbose = true")
    blank

    ws("           case ""-?"" ")
    ws("              Usage(true)")
    ws("              exit sub")
    blank

    ws("           case else")
    ws("              Usage(true)")
    ws("              exit sub")
    blank

    ws("         end select")
    blank

    ws("       i = i + 1")
    blank

    ws("     wend")
    blank

    ws("end sub")
    blank

end sub

'
' Script for converting a bool to a string
'
sub BoolStrScript

    ws("'")
    ws("' Transform a bool value to a string")
    ws("'")
    ws("function BoolStr(ByVal bValue)")
    blank
    ws("  if bValue then")
    blank
    ws("      BoolStr = " & kTrueStr)
    blank
    ws("  else")
    blank
    ws("      BoolStr = " & kFalseStr)
    blank
    ws("  end if ")
    blank
    ws("end function")
    blank

end sub

'
' Writing the script for debug output function
'
sub EchoLineScript

    ws("'")
    ws("' Print debug message according to the verbose level")
    ws("'")
    ws("sub EchoLine(ByVal Level, ByVal strALine)")
    blank
    ws("  if Level <> kVerbose or bVerbose = true then")
    blank
    ws("      wscript.echo strALine")
    blank
    ws("  end if ")
    blank
    ws("end sub")
    blank

end sub

'
' The function returns the name of the local machine
'
function strGetLocalMachineName()

    dim WSHShell
    dim WSHSysEnv

    set WSHShell = WScript.CreateObject("Wscript.Shell")
    set WSHSysEnv = WSHShell.Environment("Process")
    strGetLocalMachineName = WSHSysEnv("COMPUTERNAME")

end function

'
' Function to determine is a printer is local
'
function bIsLocal(ByVal strServerName, ByVal strPrinterName)

    on error resume next

    dim bRet

    dim oPrinter
    dim oMaster

    if strServerName <> "" then

        bIsLocal = true

        exit function

    end if

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")
    set oPrinter = CreateObject("Printer.Printer.1")

    oMaster.PrinterGet strServerName, strPrinterName, oPrinter

    if Err = 0 then

        if  ( oPrinter.Attributes and kPrinterLocal)   = kPrinterLocal  and  _
            ( oPrinter.Attributes and kPrinterNetwork) = 0              then

           '
           ' It is a local printer
           '
           bRet = true

        else

           '
           ' It is not local
           '
           bRet = false

        end if

    else

        '
        ' Error getting printer configuration, then assume it is local and try installing it
        '
        wscript.echo "Error: Get printer configuration for printer """ & strPrinterName & """" _
                     & " on server """ & strServerName & """"
        wscript.echo "Error: 0x" & Hex(Err.Number)
        if Err.Description <> "" then
            wscript.echo "Error description: " & Err.Description
        end if

        Err.Clear

        bRet = true

    end if

    bIsLocal = bRet

end function

'
' Function to remove the ServerName prefix in the printer name
'
function Strip(ByVal strOriginalPrinterName)

   dim strReturnPrinterName
   strReturnPrinterName=strOriginalPrinterName

   dim regEx
   set regEx = New RegExp
   regEx.Pattern =  "^\\\\[^\\]*\\"
   regEx.IgnoreCase = true

   '
   ' Remove the ServerName prefix by replacing  "\\*\" with ""
   '
   strReturnPrinterName=regEx.Replace(strReturnPrinterName, "")

   Strip=strReturnPrinterName

end function

'
' Function to remove the "\\" in front of the ServerName
'
function strGetNameStringOnly(ByVal strPrefixServerName)

   dim strReturn
   strReturn=strPrefixServerName

   dim regEx
   set regEx = New RegExp
   regEx.Pattern =  "^\\\\"
   regEx.IgnoreCase = true

   '
   ' Remove of "\\"
   '
   strReturn=regEx.Replace(strReturn, "")

   strGetNameStringOnly=strReturn

end function

'
' Function to change single " in the string to be double "s
'
function StuffQuote(ByVal strInput)

  Dim iIndex
  Dim strOutput
  strOutput = ""

  for iIndex = 1 to len(strInput)

    if mid(strInput, iIndex, 1) <> """" then

        ' This char is not a "

        strOutput = strOutput & mid(strInput, iIndex, 1)

    else

        ' It is a ", change it to be two "s

        strOutput = strOutput & """"""

    end if

  next

  StuffQuote = strOutput

end function


'------------------------------------------------------------------------
'      Helper functions for this program itself
'------------------------------------------------------------------------

'
' Transform a bool value to a string
'
function BoolStr(ByVal bValue)

  if bValue then

    BoolStr = kTrueStr

  else

    BoolStr = kFalseStr

  end if

end function

'
' Parse the command line into it's components
'
sub ParseCommandLine()

    dim oArgs
    dim i

    iAction = kActionUnknown

    set oArgs = wscript.Arguments

    if oArgs.Count = 0 then
        Usage(true)
        exit sub
    end if

    while i < oArgs.Count

        select case oArgs(i)

            case "-d"
                iAction = kActionDrivers

            case "-o"
                iAction = kActionPorts

            case "-p"
                iAction = kActionPrinters

            case "-f"
                iAction = kActionForms

            case "-a"
                iAction = kActionAll

            case "-c"
                i = i + 1
                strServerName = oArgs(i)

            case "-?"
                Usage(true)
                exit sub

            case else
                Usage(true)
                exit sub

        end select

        i = i + 1

    wend

end sub

'
' Display command usage.
'
sub Usage(ByVal bExit)

    wscript.echo "Usage: clone [-dopfa?] [-c server-name]"
    wscript.echo "Arguments:"
    wscript.echo "-d     - generate script for cloning the drivers"
    wscript.echo "-o     - generate script for cloning the ports"
    wscript.echo "-p     - generate script for cloning the printers"
    wscript.echo "-f     - generate script for cloning the forms"
    wscript.echo "-a     - generate script for cloning the drivers, ports, printers and forms"
    wscript.echo "-c     - source server name, default for local machine"
    wscript.echo "-?     - display command usage"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "clone -d"
    wscript.echo "clone -o -c \\server"
    wscript.echo "clone -p -c \\server"
    wscript.echo "clone -f"
    wscript.echo "clone -a"

    if bExit then
        wscript.quit(1)
    end if

end sub

'
' Function determining if the port name matches the pattern
'
function bFindPortPattern(strPattern, strString)

    dim RegEx
    set RegEx = New RegExp
    RegEx.Pattern = strPattern
    RegEx.IgnoreCase = true
    bFindPortPattern = RegEx.Test(strString)

end function

'
' Macro for writing a line in the script
'
sub ws(ByVal strALine)

   oScript.WriteLine(strALine)

end sub

sub blank()

   ws("")

end sub

'-------  End of Cloning scripts

'
'------------------------------------------------------------------------
'      Shell Installing script  (use for launching the cloning scripts)
'------------------------------------------------------------------------
'

sub InstallScript(ByVal strScriptName, ByVal strPrefixServerName)

    wscript.echo
    wscript.echo "Creating the installing script..."

    '
    ' Open the script file
    '
    set oScript = oFileSystem.CreateTextFile(strScriptName,TRUE)

    '
    ' Start creating the shell script
    '
    ws("@rem Abstract:")
    ws("@rem")
    ws("@rem " & strScriptName & " - shell script for installing all the server components")
    ws("@rem")
    blank

    ws("@echo off")
    blank

    ws("if ""%1"" == ""-?"" goto Usage")
    ws("if ""%1"" == ""/?"" goto Usage")
    blank

    dim strParameters
    strParameters = " %1 %2 %3 %4 "

    ws("cscript " & strPrefixServerName & kDriverScript   & strParameters)
    ws("cscript " & strPrefixServerName & kPortScript     & strParameters)
    ws("cscript " & strPrefixServerName & kPrinterScript  & strParameters)
    ws("cscript " & strPrefixServerName & kFormScript     & strParameters)
    blank

    ws("goto End")
    blank

    ws(":Usage")
    blank

    InstallUsageScript

    ws(":End")

    '
    ' Close the script file
    '
    oScript.Close

    wscript.echo "The script file for installing all server components is """ & strScriptName & """."

end sub

'
' The Usage script used in the install script
'
sub InstallUsageScript

    ws("    echo Usage: " & strScriptName & " [-kv?] [-c server-name]")
    ws("    echo Arguments: ")
    ws("    echo -k     - keep the existing component with the same name")
    ws("    echo -v     - verbose mode")
    ws("    echo -c     - destination server name")
    ws("    echo -?     - display command usage")
    ws("    echo.")
    ws("    echo Examples: ")
    ws("    echo    " & strScriptName )
    ws("    echo    " & strScriptName & " -v ")
    ws("    echo    " & strScriptName & " -c \\server ")
    ws("    echo    " & strScriptName & " -v -c \\server ")
    ws("    echo    " & strScriptName & " -k -v -c \\server ")
    ws("    echo    " & strScriptName & " -? ")
    blank

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

'
' Converts a driver environment string to a string
' representing the architecture of the driver.
'
function GetArchitecture(strEnvironment)

    dim strArchitecture

    if strEnvironment = kEnvironmentIntel then
        strArchitecture = kArchIntel
    elseif strEnvironment = kEnvironmentMIPS then
        strArchitecture = kArchMIPS
    elseif strEnvironment = kEnvironmentAlpha then
        strArchitecture = kArchAlpha
    elseif strEnvironment = kEnvironmentPowerPC then
        strArchitecture = kArchPowerPC
    elseif strEnvironment = kEnvironmentWindows then
        strArchitecture = kArchIntel
    else
        strArchitecture = kArchUnknown
    end if

    GetArchitecture = strArchitecture

end function

'
' Converts a driver environment string and a number to
' a string representing the driver version
'
function GetVersion(uVersion, strEnvironment)

    dim strVersion

    select case uVersion
    case 0:
        if strEnvironment = kEnvironmentWindows then
            strVersion = kVersionWindows95
        else
            strVersion = kVersionNT31

        end if

    case 1:
        if strEnvironment = kEnvironmentPowerPC then
            strVersion = kVersion351
        else
            strVersion = kVersion35x
        end if

    case 2:
        if strEnvironment = kEnvironmentPowerPC or _
           strEnvironment = kEnvironmentMIPS    or _
           strEnvironment = kEnvironmentAlpha   then
            strVersion = kVersion40
        else
            strVersion = kVersion4050
        end if

    case 3:
        strVersion = kVersion50

    case else:
        strVersion = kArchUnknown

    end select

    GetVersion = strVersion

end function
