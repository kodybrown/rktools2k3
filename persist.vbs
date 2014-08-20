'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' Abstract:
'
' persist.vbs - script for saving and restoring printer configuration
'
' Usage:
' persist [-rs?] [-b printer-name][-f file-name]
'
' Examples:
' persist.vbs -s -b \\server\printer -f file.txt -all
' persist.vbs -r -b printer -f file.txt -sec
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
const kActionSave       = 1
const kActionRestore    = 2

const kErrorSuccess     = 0
const KErrorFailure     = 1

const kPrinterData      = 1
const kPrinterInfo2     = 2
const kPrinterInfo7     = 4
const kPrinterSec       = 8
const kUserDevmode      = 16
const kPrinterDevmode   = 32
const kColorProf        = 64
const kForceName        = 128
const kResolveName      = 256
const kResolvePort      = 512
const kResolveShare     = 1024
const kDontGenerateShare= 2048

'
' kPrinterData + kPrinterInfo2 + kPrinterDevmode
'
const kMinimumSettings  = 35

'
' kMinimumSettings + kPrinterInfo7 + kPrinterSec +
' kUserDevmode + kColorProf
'
const kAllSettings      = 127

main

'
' Main execution starts here
'
sub main

    dim iAction
    dim iRetval
    dim strPrinter
    dim strFile
    dim iFlags

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(kMessage1 & vbCRLF & kMessage2 & vbCRLF & _
                          kMessage3 & vbCRLF & kMessage4 & vbCRLF & _
                          kMessage5 & vbCRLF & kMessage6 & vbCRLF)

        wscript.quit

    end if

    iRetval = ParseCommandLine(iAction, strPrinter, strFile, iFlags)

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionSave
                 iRetval = SavePrinter(strPrinter, strFile, iFlags)

            case kActionRestore
                 iRetval = RestorePrinter(strPrinter, strFile, iFlags)

            case else
                 Usage(True)
                 exit sub

        end select

    end if

end sub

'
' Save printer configuration
'
function SavePrinter(strPrinter, strFile, iFlags)

    on error resume next

    DebugPrint kDebugTrace, "In SavePrinter"

    dim oMaster

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    oMaster.PrinterPersistSave strPrinter, strFile, iFlags

    if Err.Number = kErrorSuccess then

        wscript.echo "Success saving the configuration of the printer """ & strPrinter & """ "

        SavePrinter = kErrorSuccess

    else

        wscript.echo "Unable to save the configuration of the printer """ & strPrinter _
                     & """, error 0x" & Hex(Err.Number) & ". " & Err.Description

        SavePrinter = kErrorFailure

    end if

end function

'
' Restore printer configuration
'
function RestorePrinter(strPrinter, strFile, iFlags)

    on error resume next

    DebugPrint kDebugTrace, "In RestorePrinter"

    dim oMaster

    Set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    oMaster.PrinterPersistRestore strPrinter, strFile, iFlags

    if Err.Number = kErrorSuccess then

        wscript.echo "Success restoring the configuration of the printer """ & strPrinter & """ "

        RestorePrinter = kErrorSuccess

    else

        wscript.echo "Unable to restore the configuration of the printer """ & strPrinter _
                     & """, error: 0x" & Hex(Err.Number) & ". " & Err.Description

        RestorePrinter = kErrorFailure

    end if

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
function ParseCommandLine(iAction, strPrinter, strFile, iFlags)

    DebugPrint kDebugTrace, "In the ParseCommandLine"

    dim oArgs
    dim iIndex

    iFlags = 0

    iAction = kActionUnknown
    iIndex  = 0

    set oArgs = wscript.Arguments

    while iIndex < oArgs.Count

        select case oArgs(iIndex)

            case "-r"
                iAction = kActionRestore

            case "-s"
                iAction = kActionSave

            case "-b"
                iIndex = iIndex + 1
                strPrinter = oArgs(iIndex)

            case "-f"
                iIndex = iIndex + 1
                strFile = oArgs(iIndex)

            case "-data"
                iFlags = iFlags + kPrinterData

            case "-2"
                iFlags = iFlags + kPrinterInfo2

            case "-7"
                iFlags = iFlags + kPrinterInfo7

            case "-sec"
                iFlags = iFlags + kPrinterSec

            case "-udev"
                iFlags = iFlags + kUserDevmode

            case "-pdev"
                iFlags = iFlags + kPrinterDevmode

            case "-color"
                iFlags = iFlags + kColorProf

            case "-force"
                iFlags = iFlags + kForceName

            case "-resname"
                iFlags = iFlags + kResolveName

            case "-resport"
                iFlags = iFlags + kResolvePort

            case "-resshare"
                iFlags = iFlags + kResolveShare

            case "-noshare"
                iFlags = iFlags + kDontGenerateShare

            case "-min"
                iFlags = iFlags + kMinimumSettings

            case "-all"
                iFlags = iFlags + kAllSettings

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

end function

'
' Display command usage.
'
sub Usage(bExit)

    wscript.echo "Usage: persist [-rs?] [-b printer-name][-f file-name][-data]"
    wscript.echo "               [-2][-7][-sec][-udev][-pdev][-color][-force]"
    wscript.echo "               [-resname][-resport][-resshare][-noshare][-min][-all]"
    wscript.echo ""
    wscript.echo "Arguments:"
    wscript.echo "-r        - restore printer configuration"
    wscript.echo "-s        - save printer configuration"
    wscript.echo "-?        - display command usage"
    wscript.echo "-b        - printer name"
    wscript.echo "-data     - printer data"
    wscript.echo "-2        - Printer Info 2"
    wscript.echo "-7        - Printer Info 7"
    wscript.echo "-sec      - security"
    wscript.echo "-color    - color profile"
    wscript.echo "-force    - force name"
    wscript.echo "-udev     - user devmode"
    wscript.echo "-pdev     - printer devmode"
    wscript.echo "-resname  - resolve name"
    wscript.echo "-resport  - resolve port"
    wscript.echo "-resshare - resolve share"
    wscript.echo "-noshare  - don't share the printer"
    wscript.echo "-min      - minimum settings"
    wscript.echo "-all      - all settings"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "persist.vbs -s -b \\server\printer -f file.txt -all"
    wscript.echo "persist.vbs -r -b printer -f file.txt -sec"

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

