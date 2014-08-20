'----------------------------------------------------------------------
'
' Copyright (c) Microsoft Corporation 1998-2003
' All Rights Reserved
'
' Abstract:
'
' prndata.vbs - printer data configuration script for Windows .NET Server 2003
'
' Usage:
' prndata [-gsx?] [-n name][-k key][-v value][-t int|sz|msz|bin][-d data]"

'
' Examples:
' prndata.vbs -s -n \\server\printer -k TestKey -v TestValue -t msz -d "one" "two"
' prndata.vbs -s -n \\server\printer -k TestKey -v TestValue -t int -d 53
' prndata.vbs -g -n \\server\printer -k TestKey -v TestValue
' prndata.vbs -x -n \\server\printer -k TestKey -v TestValue
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
const kActionSet        = 1
const kActionGet        = 2
const kActionDel        = 3

const kErrorSuccess     = 0
const kErrorFailure     = 1

main

'
' Main execution start here
'
sub main

    on error resume next

    dim iAction
    dim iRetval
    dim oArgs
    dim strName
    dim strKey
    dim strValue
    dim strValueType
    dim DataArray()

    '
    ' Abort if the host is not cscript
    '
    if not IsHostCscript() then

        call wscript.echo(kMessage1 & vbCRLF & kMessage2 & vbCRLF & _
                          kMessage3 & vbCRLF & kMessage4 & vbCRLF & _
                          kMessage5 & vbCRLF & kMessage6 & vbCRLF)

        wscript.quit

    end if

    strKey = ""
    strValue = ""

    iRetval = ParseCommandLine(iAction, strName, strKey, strValue, strValueType, DataArray)

    if iRetval = kErrorSuccess then

        select case iAction

            case kActionGet

                 iRetval = GetPrinterData(strName, strKey, strValue)

            case kActionSet

                 iRetval = SetPrinterData(strName, strKey, strValue, strValueType, DataArray)

            case kActionDel

                 iRetval = DelPrinterData(strName, strKey, strValue)

            case else
                 Usage(True)
                 exit sub

        end select

    end if

end sub

'
' Get the printer data
'
function GetPrinterData(strName, strKey, strValue)

    on error resume next

    DebugPrint kDebugTrace, "In GetPrinterData"
    DebugPrint kDebugTrace, "Name      " & strName
    DebugPrint kDebugTrace, "Key       " & strKey
    DebugPrint kDebugTrace, "ValueName " & strValue

    dim oMaster
    dim PrinterData
    dim iDataType
    dim iIndex
    dim iRetval

    iRetval = kErrorFailure

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    PrinterData = oMaster.PrinterDataGet(strName, strKey, strValue)

    if Err = kErrorSuccess then

        iRetval = kErrorSuccess

        iDataType = VarType(PrinterData)

        wscript.echo "Name  """ & strName  & """"
        wscript.echo "Key   """ & strKey   & """"
        wscript.echo "Value """ & strValue & """"
        wscript.echo "Printer data:"

        '
        ' Check if the return value is a simple variable or an array
        '
        if iDataType = vbLong or iDataType = vbString then

            wscript.echo PrinterData

            '
            ' Check if array
            '
            elseif iDataType = (vbArray + vbVariant) then

                '
                ' Check if array of strings
                '
                if VarType(PrinterData(0)) = vbString then

                    wscript.echo PrinterData(LBound(PrinterData))

                    for iIndex = LBound(PrinterData) + 1 to UBound(PrinterData)

                       wscript.echo PrinterData(iIndex)

                    next

                '
                ' Check if array of bytes
                '
                elseif VarType(PrinterData(0)) = vbByte then

                    PrintBinaryArray PrinterData

                end if

            else

                wscript.echo "Invalid data returned " & iDataType

        end if

    else

        wscript.echo "Error getting " & strKey & " " & Hex(Err.Number) & ". " & Err.Description

    end if

    GetPrinterData = iRetval

end function

'
' SetPrinterData
'
function SetPrinterData(strName, strKey, strValue, strValueType, PrinterData)

    on error resume next

    DebugPrint kDebugTrace, "In SetPrinterData"
    DebugPrint kDebugTrace, "Name      " & strName
    DebugPrint kDebugTrace, "Key       " & strKey
    DebugPrint kDebugTrace, "ValueName " & strValue
    DebugPrint kDebugTrace, "ValueType " & strValueType

    dim oMaster
    dim vntVariant
    dim iIndex
    dim iRetval

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    '
    ' We need the veriant in the array to have a certain type
    '
    select case strValueType

        case "int"

            vntVariant = CLng(PrinterData(0))

        case "sz"

            vntVariant = PrinterData(0)

        case "msz"

            vntVariant = PrinterData

        case "bin"

            '
            ' Convert from array of strings to array of bytes
            '
            for iIndex = LBound(PrinterData) to UBound(PrinterData)

               PrinterData(iIndex) = CByte(PrinterData(iIndex))

            next

            vntVariant = PrinterData

        case else

            wscript.echo "Invalid type"

            exit function

    end select

    oMaster.PrinterDataSet strName, strKey, strValue, vntVariant

    if Err = kErrorSuccess then

        wscript.echo "Success setting printer data"

        iRetval = kErrorSuccess

    else

        wscript.echo "Error setting printer data, error " & Hex(Err.Number) & ". " & Err.Description

        iRetval = kErrorFailure

    end if

    SetPrinterData = iRetval

end function

'
' DelPrinterData
'
function DelPrinterData(strPrinter, strKey, strValue)

    on error resume next

    DebugPrint kDebugTrace, "In DelPrinterData"
    DebugPrint kDebugTrace, "Printer   " & strPrinter
    DebugPrint kDebugTrace, "Key       " & strKey
    DebugPrint kDebugTrace, "ValueName " & strValue

    dim oMaster
    dim iRetval

    set oMaster = CreateObject("PrintMaster.PrintMaster.1")

    '
    ' Check whether to delete a key or a value
    '
    if strValue <> "" then

        oMaster.PrinterDataDel strPrinter, strKey, strValue

    else

        oMaster.PrinterKeyDel strPrinter, strKey

    end if

    if Err = kErrorSuccess then

        wscript.echo "Success deleting printer data"

        iRetval = kErrorSuccess

    else

        wscript.echo "Error deleting printer data, error: " & Hex(Err.number) & ". " & Err.Description

        iRetval = kErrorFailure

    end if

    DelPrinterData = iRetval

end function

'
' Displays an array of bytes
'
sub PrintBinaryArray(DataArray)

    dim strString
    dim iIndex
    dim iCount
    dim iValue

    strString = ""

    for iIndex = LBound(DataArray) to UBound(DataArray)

        if iIndex <> 0 then

            if iIndex mod 16 = 0 then

                strString = strString + Chr(13) + Chr(10)

            elseif iIndex mod 8 = 0 then

                strString = strString + "-"

            else

                strString = strString + " "

            end if

        end if

        iValue = DataArray(iIndex)

        if iValue > 15 then

            strString = strString + hex(iValue)

        else

            strString = strString + "0" + hex(iValue)

        end if

    next

    wscript.echo strString

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
function ParseCommandLine(iAction, strName, strKey, strValueName, strValueType, DataArray)

    on error resume next

    DebugPrint kDebugTrace, "In the ParseCommandLine"

    dim oArgs
    dim i
    dim j

    iAction = kActionUnknown

    set oArgs = wscript.Arguments

    while i < oArgs.Count

        select case oArgs(i)

            case "-g"
                iAction = kActionGet

            case "-s"
                iAction = kActionSet

            case "-x"
                iAction = kActionDel

            case "-n"
                i = i + 1
                strName = oArgs(i)

            case "-k"
                i = i + 1
                strKey = oArgs(i)

            case "-v"
                i = i + 1
                strValueName = oArgs(i)

            case "-t"
                i = i + 1
                strValueType = oArgs(i)

            case "-d"
                i = i + 1

                j = 0

                '
                ' Add all values following -v to the array
                '
                while i < oArgs.Count

                    '
                    ' Increase the size of the array and keep all the values
                    '
                    redim preserve DataArray(j)

                    DataArray(j) = oArgs(i)

                i = i + 1

                j = j + 1

                wend

            case "-?"
                Usage(true)

            case else
                Usage(true)

        end select

        i = i + 1

    wend

    if Err = kErrorSuccess then

        ParseCommandLine = kErrorSuccess

    else

        ParseCommandLine = kErrorFailure

        wscript.echo "Error parsing the command line, error: 0x" & hex(Err.NUmber) & ". " &Err.Description

    end if

end function

'
' Display command usage.
'
sub Usage(ByVal bExit)

    wscript.echo "Usage: prndata [-gsx?] [-n name][-k key][-v value][-t int|sz|msz|bin][-d data]"
    wscript.echo "Arguments:"
    wscript.echo "-d     - data value: must be last option on the line"
    wscript.echo "-g     - get key value data"
    wscript.echo "-k     - key name"
    wscript.echo "-n     - server name or printer name"
    wscript.echo "-s     - set key vlaue data"
    wscript.echo "-t     - value type: int for integer, sz for string, msz for multi string, bin for binary data"
    wscript.echo "-v     - value name. Can be any one of the predefined values for print servers. Ex: ""DefaultSpoolDirectory"" "
    wscript.echo "         See GetPrinterData in the Platform Software Development Kit for more details"
    wscript.echo "-x     - delete a key or a value under a key"
    wscript.echo "-?     - display command usage"
    wscript.echo ""
    wscript.echo "Examples:"
    wscript.echo "prndata.vbs -s -n \\server\printer -k TestKey -v TestValue -t msz -d ""one"" ""two"""
    wscript.echo "prndata.vbs -s -n \\server\printer -k TestKey -v TestValue -t int -d 53"
    wscript.echo "prndata.vbs -g -n \\server\printer -k TestKey -v TestValue"
    wscript.echo "prndata.vbs -x -n \\server\printer -k TestKey -v TestValue"
    wscript.echo "prndata.vbs -x -n \\server\printer -k TestKey"
    wscript.echo "prndata.vbs -g -n \\server -v MajorVersion"

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

