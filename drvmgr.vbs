'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' Abstract:
'
' drvmgr.vbs - driver script for Windows .NET Server 2003
'
' Usage:
' drvmgr [-adl?] [-m model name] [-v Version] [-p Path]
'                [-c Server] [-t Architecture] [-i InfFile]
'
' Example:
' drvmgr -d -m "driver" -v "Windows XP and Windows .NET Server 2003" -t Intel
' drvmgr -l -c \\server
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
const kActionDel        = 2
const kActionDelAll     = 3
const kActionList       = 4

const kErrorSuccess     = 0
const kErrorFailure     = 1

'
' Strings identifying environments
'
const kEnvironmentIntel   = "Windows NT x86"
const kEnvironmentItanium = "Windows IA64"
const kEnvironmentMIPS    = "Windows NT R4000"
const kEnvironmentAlpha   = "Windows NT Alpha_AXP"
const kEnvironmentPowerPC = "Windows NT PowerPC"
const kEnvironmentWindows = "Windows 4.0"
const kEnvironmentUnknown = "unknown"

'
' Strings identifying architectures
'
const kArchIntel   = "Intel"
const kArchItanium = "Itanium"
const kArchMIPS    = "MIPS"
const kArchAlpha   = "Alpha"
const kArchPowerPC = "PowerPC"
const kArchUnknown = "Unknown"

'
' Strings identifying driver versions
' Change these strings on localized builds
'
const kVersionWindows95 = "Windows 95, Windows 98, and Windows Millennium Edition"
const kVersion_NT31     = "Windows NT 3.1"
const kVersion35x       = "Windows NT 3.5 or 3.51"
const kVersion351       = "Windows NT 3.51"
const kVersion40        = "Windows NT 4.0"
const kVersion50        = "Windows 2000, Windows XP and Windows .NET Server 2003"
const kVersion512       = "Windows XP and Windows .NET Server 2003"

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strServer
    dim strModel
    dim strPath
    dim strVersion
    dim strArchitecture
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

    iRetval = ParseCommandLine(iAction, strServer, strModel, strPath, strVersion, _
                               strArchitecture, strInfFile)

    if iRetval = kErrorSuccess  then

        select case iAction

            case kActionAdd
                iRetval = AddDriver(strServer, strModel, strPath, strVersion, _
                                    strArchitecture, strInfFile)

            case kActionDel
                iRetval = DelDriver(strServer, strModel, strVersion, strArchitecture)

            case kActionDelAll
                iRetval = DelAllDrivers(strServer)

            case kActionList
                iRetval = ListDrivers(strServer)

            case kActionUnknown
                Usage(true)
                exit sub

            case else
                Usage(true)
                exit sub

        end select

    end if

end sub

'
' Add a driver
'
function AddDriver(strServer, strModel, strPath, strVersion, strArchitecture, strInfFile)

    on error resume next

    DebugPrint kDebugTrace, "In AddDriver"

    dim oMaster
    dim oDriver
    dim iResult

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")
    set oDriver = CreateObject("Driver.Driver.1")

    oDriver.ModelName          = strModel
    oDriver.Path               = strPath
    oDriver.DriverArchitecture = strArchitecture
    oDriver.InfFile            = strInfFile
    oDriver.ServerName         = strServer
    odriver.DriverVersion      = strVersion

    oMaster.DriverAdd oDriver

    if Err.Number = kErrorSuccess then

        wscript.echo "Added driver """ & oDriver.ModelName & """"

        iResult = kErrorSuccess

    else

        wscript.echo "Unable to add driver """ & oDriver.ModelName & """, error: 0x" _
                     & Hex(Err.Number) & ". " & Err.Description

        iResult = kErrorFailure

    end if

    AddDriver = iResult

end function

'
' Delete a driver
'
function DelDriver(strServer, strModel, strVersion, strArchitecture)

    on error resume next

    DebugPrint kDebugTrace, "In DelDriver"

    dim oMaster
    dim oDriver
    dim iRetval

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")
    set oDriver = CreateObject("Driver.Driver.1")

    oDriver.ModelName          = strModel
    oDriver.DriverArchitecture = strArchitecture
    oDriver.ServerName         = strServer
    odriver.DriverVersion      = strVersion

    oMaster.DriverDel oDriver

    if Err.Number = kErrorSuccess then

        wscript.echo "Deleted driver """ & oDriver.ModelName & """"

        iRetval = kErrorSuccess

    else

        wscript.echo "Unable to delete driver """ & oDriver.ModelName  & """, error: 0x" _
                     & Hex(Err.Number) & ". " & Err.Description

        iRetval = kErrorFailure

    end if

    DelDriver = iRetval

end function

'
' Delete all drivers
'
function DelAllDrivers(strServer)

    on error resume next

    DebugPrint kDebugTrace, "In DelAllDrivers"

    dim oMaster
    dim oDriver
    dim iResult
    dim iTotal
    dim iTotalDeleted

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")
    set oDriver = CreateObject("Driver.Driver.1")

    iTotal = 0
    iTotalDeleted = 0

    for each oDriver in oMaster.Drivers(strServer)

        if Err.Number = kErrorSuccess then

            iTotal = iTotal + 1

            wscript.echo
            wscript.echo "Attempting to delete driver: " & oDriver.ModelName
            wscript.echo "               architecture: " & GetArchitecture(oDriver.Environment)
            wscript.echo "                    version: " & GetVersion(oDriver.Version, oDriver.Environment)
            wscript.echo "                from server: " & oDriver.ServerName

            oMaster.DriverDel oDriver

            if Err.Number = kErrorSuccess then

                wscript.echo "Success: Driver """ & oDriver.ModelName & """ was deleted"

                iTotalDeleted = iTotalDeleted + 1

            else

                wscript.echo "Unable to delete driver """ & oDriver.ModelName  & """, error: 0x" _
                             & Hex(Err.Number) & ". " & Err.Description

                Err.Clear

            end if

         else

             wscript.echo "Unable to delete drivers on server, error: 0x" & _
                          Hex(Err.Number) & ". " &  Err.Description

             DelAllDrivers = kErrorFailure

             exit function

         end if

    next

    wscript.echo "Number of drivers " & iTotal & ". Drivers deleted " & iTotalDeleted

    DelAllDrivers = kErrorSuccess

end function

'
' List drivers
'
function ListDrivers(strServer)

    on error resume next

    DebugPrint kDebugTrace, "In ListDriver"

    dim oMaster
    dim oDriver
    dim iResult
    dim iTotal
    dim vntDependentFiles

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    for each oDriver in oMaster.Drivers(strServer)

        if Err.Number = kErrorSuccess then

            wscript.echo ""
            wscript.echo "ServerName    : " & oDriver.ServerName
            wscript.echo "DriverName    : " & oDriver.ModelName
            wscript.echo "Version       : " & oDriver.Version
            wscript.echo "DriverVersion : " & GetVersion(oDriver.Version, oDriver.Environment)
            wscript.echo "DriverPath    : " & oDriver.Path
            wscript.echo "Environment   : " & oDriver.Environment
            wscript.echo "Architecture  : " & GetArchitecture(oDriver.Environment)
            wscript.echo "MonitorName   : " & oDriver.MonitorName
            wscript.echo "DataFile      : " & oDriver.DataFile
            wscript.echo "ConfigFile    : " & oDriver.ConfigFile
            wscript.echo "HelpFile      : " & oDriver.HelpFile

            vntDependentFiles = oDriver.DependentFiles

            '
            ' If there are no dependent files, the method will set DependentFiles to
            ' an empty variant, so we check if the variant is an array of variants
            '
            if VarType(vntDependentFiles) = (vbArray + vbVariant) then

                PrintDepFiles oDriver.DependentFiles

            end if

            Err.Clear

        else

            wscript.echo "Unable to list drivers, error: 0x" & Hex(Err.Number) & _
                         ". " & Err.Description

            ListDrivers = iErrorFailure

            exit function

        end if

    next

    wscript.echo "Success listing drivers"

    ListDrivers = kErrorSuccess

end function

'
' Prints the contents of an array of variants
'
sub PrintDepFiles(Param)

   on error resume next

   dim iIndex

   iIndex = LBound(Param)

   if Err.Number = 0 then

      wscript.echo "Dependent Files "

      for iIndex = LBound(Param) to UBound(Param)

          wscript.echo "                " & Param(iIndex)

      next

   else

        wscript.echo "Unable to print the dependent files, error 0x" & _
                     Hex(Err.Number ) & ". " & Err.Description

   end if

end sub

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
function ParseCommandLine(iAction, strServer, strModel, strPath, strVersion, strArchitecture, strInfFile)

    on error resume next

    DebugPrint kDebugTrace, "In the ParseCommandLine"

    dim oArgs
    dim iIndex

    iAction = kActionUnknown
    iIndex = 0

    set oArgs = wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-a"
                iAction = kActionAdd

            case "-d"
                iAction = kActionDel

            case "-l"
                iAction = kActionList

            case "-x"
                iAction = kActionDelAll

            case "-c"
                iIndex = iIndex + 1
                strServer = oArgs(iIndex)

            case "-m"
                iIndex = iIndex + 1
                strModel = oArgs(iIndex)

            case "-p"
                iIndex = iIndex + 1
                strPath = oArgs(iIndex)

            case "-v"
                iIndex = iIndex + 1
                strVersion = oArgs(iIndex)

            case "-t"
                iIndex = iIndex + 1
                strArchitecture = oArgs(iIndex)

            case "-i"
                iIndex = iIndex + 1
                strInfFile = oArgs(iIndex)

            case "-?"
                Usage(true)
                exit function

            case else
                Usage(true)
                exit function

        end select

        iIndex = iIndex + 1

    wend

    if Err.Number <> 0 then

        wscript.echo "Unable to parse command line, error 0x" & _
                     Hex(Err.Number) & ". " & Err.Description

        ParseCommandLine = kErrorFailure

    else

        ParseCommandLine = kErrorSuccess

    end if

end  function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo "Usage: drvmgr [-adlx?] [-m model] [-v version] [-p path]"
    wscript.echo "                       [-c server] [-t architecture] [-i inf file]"
    wscript.echo "Arguments:"
    wscript.echo "-a     - add the specified driver"
    wscript.echo "-c     - server name"
    wscript.echo "-d     - delete the specified driver"
    wscript.echo "-i     - inf file name"
    wscript.echo "-l     - list all drivers"
    wscript.echo "-m     - driver model name"
    wscript.echo "-p     - driver file path"
    wscript.echo "-t     - architecture"
    wscript.echo "-v     - version"
    wscript.echo "-x     - delete all drivers that are not in use"
    wscript.echo "-?     - display command usage"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "drvmgr -a -m ""driver"" -t Intel -v ""Windows 2000, Windows XP and Windows .NET Server 2003"""
    wscript.echo "drvmgr -a -m ""driver"" -t Intel -v ""Windows NT 4.0 or 2000"
    wscript.echo "drvmgr -a -m ""driver"" -t Intel -v ""Windows 95, Windows 98, and Windows Millennium Edition"" -p c:\drv\win9x"
    wscript.echo "drvmgr -a -m ""driver"" -t Itanium -v ""Windows XP and Windows .NET Server 2003"""
    wscript.echo "drvmgr -d -m ""driver"" -v ""Windows 2000, Windows XP and Windows .NET Server 2003"" -t Intel"
    wscript.echo "drvmgr -l -c \\server"
    wscript.echo "drvmgr -x -c \\server"

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
    elseif strEnvironment = kEnvironmentItanium then
        strArchitecture = kArchItanium
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
        strVersion = kVersion40

    case 3:
        if strEnvironment = kEnvironmentItanium then
            strVersion = kVersion512
        else
            strVersion = kVersion50
        end if

    case else:
        strVersion = kArchUnknown

    end select

    GetVersion = strVersion

end function
