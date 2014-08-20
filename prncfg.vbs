'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' Abstract:
'
' prncfg.vbs - printer configuration script for Windows .NET Server 2003
'
' Usage:
' prncfg [-gs?] [-b printer][-r port]
'               [-l location][-m comment][-s share][-f sep-file]
'               [-t data-type][-a attributes [+|-]value> etc.]
' Examples:
' prncfg.vbs -g -b \\server\printer
' prncfg.vbs -s -b printer -l "Office" -m "driver a"
' prncfg.vbs -s -b printer -h "Share" -a "attributes +shared attributes -direct"
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
const kActionSet        = 1
const kActionGet        = 2

const kErrorSuccess     = 0
const KErrorFailure     = 1

const kPrinterQueued          = 1
const kPrinterDirect          = 2
const kPrinterDefault         = 4
const kPrinterShared          = 8
const kPrinterNetwork         = 16
const kPrinterHidden          = 32
const kPrinterLocal           = 64
const kPrinterEnableDevq      = 128
const kPrinterKeepPrinterJobs = 256
const kPrinterDoCompleteFirst = 512
const kPrinterWorkOffline     = 1024
const kPrinterEnableBidi      = 2048
const kPrinterRawOnly         = 4096
const kPrinterPublished       = 8192

const kPrinterStatusPaused           = 1
const kPrinterStatusError            = 2
const kPrinterStatusPendingDeletion  = 4
const kPrinterStatusPapeJam          = 8
const kPrinterStatusPaperOut         = 16
const kPrinterStatusManualFeed       = 32
const kPrinterStatusPaperProblem     = 64
const kPrinterStatusOffline          = 128
const kPrinterStatusIOActive         = 256
const kPrinterStatusBusy             = 512
const kPrinterStatusPrinting         = 1024
const kPrinterStatusOuptutBinFull    = 2048
const kPrinterStatusNotAvailable     = 4096
const kPrinterStatusWaiting          = 8192
const kPrinterStatusProcessing       = 16834
const kPrinterStatusInitializing     = 32768
const kPrinterStatusWarmingUp        = 65536
const kPrinterStatusTonerLow         = 131072
const kPrinterStatusNoToner          = 262144
const kPrinterStatusPagePunt         = 524288
const kPrinterStatusUserIntervention = 1048576
const kPrinterStatusOutOfMemory      = 2097152
const kPrinterStatusDoorOpen         = 4194304
const kPrinterStatusServerUnknown    = 8388608
const kPrinterStatusPowerSave        = 16777216


main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strPrinter, strPort, strShare, strComment
    dim strLocation, Data, strSep, strNewName
    dim ParamDict, AttributeDictionary, StatusDictionary

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(kMessage1 & vbCRLF & kMessage2 & vbCRLF & _
                          kMessage3 & vbCRLF & kMessage4 & vbCRLF & _
                          kMessage5 & vbCRLF & kMessage6 & vbCRLF)

        wscript.quit

    end if

    set ParamDict           = CreateObject("Scripting.Dictionary")
    set AttributeDictionary = CreateObject("Scripting.Dictionary")
    set StatusDictionary    = CreateObject("Scripting.Dictionary")

    BuildAttributeDictionary  AttributeDictionary
    BuildStatusDictionary     StatusDictionary

    iRetval = ParseCommandLine(iAction, strPrinter, strPort, strShare, strComment, _
                               strLocation, Data, strSep, strNewName, ParamDict)

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionSet
                 iRetval = SetPrinter(strPrinter, strPort, strShare, strComment, _
                                      strLocation, Data, strSep, strNewName, ParamDict)

            case kActionGet
                 iRetval = GetPrinter(strPrinter, AttributeDictionary, StatusDictionary)

            case else
                 Usage(True)
                 exit sub

        end select

    end if

end sub

'
' Get printer configuration
'
function GetPrinter(strPrinterName, AttributeDictionary, StatusDictionary)

    on error resume next

    DebugPrint kDebugTrace, "In GetPrinter"

    dim oPrinter
    dim oMaster
    dim iRetval

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")
    set oPrinter = CreateObject("Printer.Printer.1")

    oMaster.PrinterGet "", strPrinterName, oPrinter

    if Err.Number = kErrorSuccess then

        wscript.echo "Success: getting printer config"

        wscript.echo
        wscript.echo "PrinterName:  " & oPrinter.PrinterName
        wscript.echo "ShareName:    " & oPrinter.ShareName
        wscript.echo "PortName:     " & oPrinter.PortName
        wscript.echo "DriverName    " & oPrinter.DriverName
        wscript.echo "Comment:      " & oPrinter.Comment
        wscript.echo "Location:     " & oPrinter.Location
        wscript.echo "SepFile:      " & oPrinter.Sepfile
        wscript.echo "PrintProc:    " & oPrinter.PrintProcessor
        wscript.echo "Datatype:     " & oPrinter.Datatype
        wscript.echo "Parameters:   " & oPrinter.Parameters

        BuildExplanationString AttributeDictionary, oPrinter.Attributes, "Attributes:   "

        wscript.echo "Priority:     " & CStr(oPrinter.Priority)
        wscript.echo "DefaultPri:   " & CStr(oPrinter.DefaultPriority)
        wscript.echo "StartTime:    " & CStr(oPrinter.StartTime)
        wscript.echo "UntilTime:    " & CStr(oPrinter.UntilTime)

        if oPrinter.Status = 0 then

            wscript.echo "Status:       Ready"

        else

            BuildExplanationString StatusDictionary, oPrinter.Status, "Status:       "

        end if

        wscript.echo "Jobcount:     " & CStr(oPrinter.Jobs)
        wscript.echo "AveragePPM    " & CStr(oPrinter.AveragePPM)
        wscript.echo

        iRetval = kErrorSuccess

    else

        wscript.echo "Unable to get the printer config, error: 0x" & _
                     Hex(Err.Number) & ". " & Err.Description

        iRetval = kErrorFailure

    end if

    GetPrinter = iRetval

end function

'
' Configure a printer
'
function SetPrinter(strPrinter, strPort, strShare, strComment, strLocation, Data, strSep, strNewName, AttrDict)

    on error resume next

    DebugPrint kDebugTrace, "In SetPrinter"

    dim oPrinter
    dim oMaster
    dim iRetval

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")
    set oPrinter = CreateObject("Printer.Printer.1")

    oMaster.PrinterGet "", strPrinter, oPrinter

    if strPort <> "" then

        oPrinter.PortName = strPort

    end if

    if strShare <> "" then

        oPrinter.ShareName = strShare

    end if

    if strLocation <> "" then

        oPrinter.Location = strLocation

    end if

    if strComment <> "" then

        oPrinter.Comment = strComment

    end if

    if Data <> "" then

        oPrinter.DataType = Data

    end if

    oPrinter.NewName = strNewName

    oPrinter.SepFile = strSep

    ' Field Queued
    '
    if AttrDict.Exists("queued") then

        oPrinter.Queued = AttrDict.Item("queued")

    end if

    ' Field Direct
    '
    if AttrDict.Exists("direct") then

        oPrinter.Direct = AttrDict.Item("direct")

    end if

    ' Field Default
    '
    if AttrDict.Exists("default") then

        oPrinter.Default = AttrDict.Item("default")

    end if

    ' Field Shared
    '
    if AttrDict.Exists("shared") then

        oPrinter.Shared = AttrDict.Item("shared")

    end if

    ' Field Hidden
    '
    if AttrDict.Exists("hidden") then

        oPrinter.Hidden = AttrDict.Item("hidden")

    end if

    ' Field EnableDevq
    '
    if AttrDict.Exists("enabledevq") then

        oPrinter.EnableDevq = AttrDict.Item("enabledevq")

    end if

    ' Field KeepPrintedJobs
    '
    if AttrDict.Exists("keepprintedjobs") then

        oPrinter.KeepPrintedJobs = AttrDict.Item("keepprintedjobs")

    end if

    ' Field DocompleteFirst
    '
    if AttrDict.Exists("docompletefirst") then

        oPrinter.DoCompleteFirst = AttrDict.Item("docompletefirst")

    end if

    ' Field workOffline
    '
    if AttrDict.Exists("workoffline") then

        oPrinter.WorkOffline = AttrDict.Item("workoffline")

    end if

    ' Field EnableBidi
    '
    if AttrDict.Exists("enablebidi") then

        oPrinter.EnableBidi = AttrDict.Item("enablebidi")

    end if

    ' Field RawOnly
    '
    if AttrDict.Exists("rawonly") then

        oPrinter.RawOnly = AttrDict.Item("rawonly")

    end if

    ' Field Published
    '
    if AttrDict.Exists("published") then

        oPrinter.Published = AttrDict.Item("published")

    end if

    oMaster.PrinterSet oPrinter

    if Err.Number = kErrorSuccess then

        wscript.echo "Success: configuring printer """ & strPrinter & """ "

        iRetval = kErrorSuccess

    else

        wscript.echo "Unable to configure printer """ & strPrinter & """, error: 0x"_
                     & Hex(Err.Number) & " " & Err.Description

        iRetval = kErrorFailure

    end if

    SetPrinter = iRetval

end function

'
' Builds a string description of the number
' The bits in the number have values associated in the dictionary
'
sub BuildExplanationString(oDict, Number, strInit)

    on error resume next

    dim strExpl
    dim AllKeys
    dim iIndex

    strExpl = strInit

    AllKeys = oDict.Keys

    for iIndex = 0 to oDict.Count -1

        if (Number and AllKeys(iIndex)) = AllKeys(iIndex) then

            strExpl = strExpl + oDict.Item(AllKeys(iIndex))

        end if

    next

    wscript.echo strExpl

end sub

'
' Initializes the AttributeDictionary
'
sub BuildAttributeDictionary(AttrExplanationDict)

    AttrExplanationDict.Add kPrinterQueued,          "Queued "
    AttrExplanationDict.Add kPrinterDirect,          "Direct "
    AttrExplanationDict.Add kPrinterDefault,         "Default "
    AttrExplanationDict.Add kPrinterShared,          "Shared "
    AttrExplanationDict.Add kPrinterNetwork,         "Network "
    AttrExplanationDict.Add kPrinterHidden,          "Hidden "
    AttrExplanationDict.Add kPrinterLocal,           "Local "
    AttrExplanationDict.Add kPrinterEnableDevq,      "EnableDevq "
    AttrExplanationDict.Add kPrinterKeepPrinterJobs, "KeepPrintedJobs "
    AttrExplanationDict.Add kPrinterDoCompleteFirst, "DoCompleteFirst "
    AttrExplanationDict.Add kPrinterWorkOffline,     "WorkOffLine "
    AttrExplanationDict.Add kPrinterEnableBidi,      "EnbleBiDi "
    AttrExplanationDict.Add kPrinterRawOnly,         "RawOnly "
    AttrExplanationDict.Add kPrinterPublished,       "Published "

end sub

'
' Initializes the AttributeDictionary
'
sub BuildStatusDictionary(StatusDict)

    StatusDict.Add kPrinterStatusPaused,          "Paused "
    StatusDict.Add kPrinterStatusError,           "Error "
    StatusDict.Add kPrinterStatusPendingDeletion, "PendingDeletion "
    StatusDict.Add kPrinterStatusPapeJam,         "PaperJam "
    StatusDict.Add kPrinterStatusPaperOut,        "PaperOut "
    StatusDict.Add kPrinterStatusManualFeed,      "ManualFeed "
    StatusDict.Add kPrinterStatusPaperProblem,    "PaperProblem "
    StatusDict.Add kPrinterStatusOffline,         "Offline "
    StatusDict.Add kPrinterStatusIOActive,        "IOActive "
    StatusDict.Add kPrinterStatusBusy,            "Busy "
    StatusDict.Add kPrinterStatusPrinting,        "Printing "
    StatusDict.Add kPrinterStatusOuptutBinFull,   "OutputBinFull "
    StatusDict.Add kPrinterStatusNotAvailable,    "NotAvailable "
    StatusDict.Add kPrinterStatusWaiting,         "Waiting "
    StatusDict.Add kPrinterStatusProcessing,      "Processing "
    StatusDict.Add kPrinterStatusInitializing,    "Initializing "
    StatusDict.Add kPrinterStatusWarmingUp,       "Warming Up "
    StatusDict.Add kPrinterStatusTonerLow,        "TonerLow "
    StatusDict.Add kPrinterStatusNoToner,         "NoToner "
    StatusDict.Add kPrinterStatusPagePunt,        "PagePunt "
    StatusDict.Add kPrinterStatusUserIntervention,"UserIntervention "
    StatusDict.Add kPrinterStatusOutOfMemory,     "OutOfMemory "
    StatusDict.Add kPrinterStatusDoorOpen,        "DoorOpen "
    StatusDict.Add kPrinterStatusServerUnknown,   "ServerUnknown "
    StatusDict.Add kPrinterStatusPowerSave,       "PowerSave "

end sub

'
' Prints the contents of the dictionary
'
sub PrintDictionary(oDict)

   dim KeyArray
   dim iIndex

   wscript.echo "Iterating the dictionary"

   KeyArray = oDict.Keys

   for iIndex = 0 to oDict.Count -1

       wscript.echo KeyArray(iIndex) & "  " & dict.Item(KeyArray(iIndex))

   next

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
function ParseCommandLine(iAction, strPrinter, strPort, strShare, strComment, strLocation, Data, strSep, strNewName, AttrDict)

    on error resume next

    DebugPrint kDebugTrace, "In the ParseCommandLine"

    dim oArgs
    dim iIndex

    iAction = kActionUnknown
    iIndex = 0

    set oArgs = wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-g"
                iAction = kActionGet

            case "-s"
                iAction = kActionSet

            case "-b"
                iIndex = iIndex + 1
                strPrinter = oArgs(iIndex)

            case "-r"
                iIndex = iIndex + 1
                strPort = oArgs(iIndex)

            case "-h"
                iIndex = iIndex + 1
                strShare = oArgs(iIndex)

            case "-m"
                iIndex = iIndex + 1
                strComment = oArgs(iIndex)

            case "-l"
                iIndex = iIndex + 1
                strLocation = oArgs(iIndex)

            case "-t"
                iIndex = iIndex + 1
                Data = oArgs(iIndex)

            case "-f"
                iIndex = iIndex + 1
                strSep = oArgs(iIndex)

            case "-w"
                iIndex = iIndex + 1
                strNewName = oArgs(iIndex)

            case "-queued"
                AttrDict.Add "queued", false

            case "+queued"
                AttrDict.Add "queued", true

            case "-direct"
                AttrDict.Add "direct", false

            case "+direct"
                AttrDict.Add "direct", true

            case "-default"
                AttrDict.Add "default", false

            case "+default"
                AttrDict.Add "default", true

            case "-shared"
                AttrDict.Add "shared", false

            case "+shared"
                AttrDict.Add "shared", true

            case "-hidden"
                AttrDict.Add "hidden", false

            case "+hidden"
                AttrDict.Add "hidden", true

            case "-enabledevq"
                AttrDict.Add "enabledevq", false

            case "+enabledevq"
                AttrDict.Add "enabledevq", true

            case "-keepprintedjobs"
                AttrDict.Add "keepprintedjobs", false

            case "+keepprintedjobs"
                AttrDict.Add "keepprintedjobs", true

            case "-docompletefirst"
                AttrDict.Add "docompletefirst", false

            case "+docompletefirst"
                AttrDict.Add "docompletefirst", true

            case "-workoffline"
                AttrDict.Add "workoffline", false

            case "+workoffline"
                AttrDict.Add "workoffline", true

            case "-enablebidi"
                AttrDict.Add "enablebidi", true

            case "+enablebidi"
                AttrDict.Add "enablebidi", true

            case "-rawonly"
                AttrDict.Add "rawonly", false

            case "+rawonly"
                AttrDict.Add "rawonly", true

            case "-published"
                AttrDict.Add "published", false

            case "+published"
                AttrDict.Add "published", true

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

        wscript.echo "Unable to parse command line, error 0x" & Hex(Err.Number) _
                     & ". " & Err.Description

        ParseCommandLine = kErrorFailure

    end if

end function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo "Usage: prncfg [-gs?] [-b printer][-r port][-w new printer name]"
    wscript.echo "              [-l location][-m comment][-h share name][-f sep file]"
    wscript.echo "              [-t datatype][<+|->shared][<+|->direct][<+\->default][<+|->published]"
    wscript.echo "              [<+|->rawonly][<+|->keepprintedjobs][<+|->queued]"
    wscript.echo "Arguments:"
    wscript.echo "-g     - get configuration"
    wscript.echo "-s     - set configuration"
    wscript.echo "-?     - display command usage"
    wscript.echo "-b     - printer name"
    wscript.echo "-r     - port name"
    wscript.echo "-w     - new printer name"
    wscript.echo "-l     - location string"
    wscript.echo "-h     - share name"
    wscript.echo "-f     - separator file string"
    wscript.echo "-t     - data type string"
    wscript.echo "-m     - comment string"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "prncfg.vbs -g -b \\server\printer"
    wscript.echo "prncfg.vbs -s -b Printer -l ""Building A/Floor 100/Office 1"""
    wscript.echo "prncfg.vbs -s -b Printer -h ""Share"" +shared -direct"
    wscript.echo "prncfg.vbs -s -b Printer +rawonly +keepprintedjobs"

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

