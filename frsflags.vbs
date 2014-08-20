'****************************************************************************
'*
'*  Copyright (c) 2003 Microsoft Corporation.  All rights reserved.
'* 
'*  File:           frsflags.vbs
'*  Created:        09/10/2002
'*  Author:         Huseyin Dursun [HuseyinD]
'*
'*  Main Function:  Updating/displaying "fRSFlags" attribute of a given 
'*                  SYSVOL or DFS replica under a given domain. 
'*
'*
'*  cscript frsflag.vbs { -s | -u | -d } <ReplicaName>  <DomainName>
'*
'*
'****************************************************************************


Option Explicit 

Call Run()


'****************************************************************************
'*
'*  Function Run
'*
'****************************************************************************

Function Run()

   Dim oTarget
   Dim objArg
   Dim intArgCount
   Dim intFlag
   Dim Root
   Dim Replica
   Dim Domain
   Dim DomainController
   Dim LDAPStr
   Dim InstallOverrideStatus

'--
'--  Input validation and argument processing
'--
    Set objArg  = WScript.Arguments
    intArgCount = objArg.Count

 
    If ( intArgCount <> 3 ) Then
       Usage()
    ElseIf objArg(0) = "-s" Then
       intFlag = 1
    ElseIf objArg(0) = "-u" Then
       intFlag = 0
    ElseIf objArg(0) = "-d" Then
       intFlag = 2
    Else 
       Usage()
    End If

    Replica = objArg(1)
    Domain  = objArg(2)

'--
'-- Generate AD compatible domain name out of user input dotted domain name
'--
    DomainController = "DC=" & Replace(Domain, ".", ",DC=")    

'--
'-- Replace sysvol with the full sysvol replica name and generate LDAP
'-- search string . 
'-- 

    If (LCase(Replica) = "sysvol") Then

        Replica = "Domain System volume (SYSVOL Share)"

        LDAPStr =  "LDAP://" & Domain & "/CN=" & Replica                     & _
                   ",CN=File Replication Service,CN=System,"                 & _
                   DomainController
    Else

        '--
        '--  We need to extract the root from DFS link
        '--
        
        Root = mid( Replica, 1, InStr( 1, Replica, "|", 1) - 1 )

        LDAPStr =  "LDAP://" & Domain & "/CN=" & Replica & ",CN=" & Root     & _
                   ",CN=DFS Volumes,CN=File Replication Service,CN=System,"  & _
                   DomainController

    End If 
 
    WScript.Echo "  " & LDAPStr


'--
'-- Now, access the replica container in AD and get the value of fRSFlags
'--

    Set oTarget = GetObject(LDAPStr)

    If intFlag = 1 Then
       InstallOverrideStatus = " -- Install Override Enabled"
    ElseIf intFlag = 0 Then
       InstallOverrideStatus = " -- Install Override Disabled"
    Else
       InstallOverrideStatus = ""
    End If

    WScript.Echo "  Replica/Domain    : " & Replica & "/" & DomainController
    WScript.Echo "  Current FRS Flags : " & oTarget.fRSFlags 


'--
'-- If it is not a "show" request then update the attribute in DS
'--

    If NOT intFlag = 2 Then 
        oTarget.fRSFlags = intFlag
        oTarget.SetInfo
        WScript.Echo "  New FRS Flags     : " & oTarget.fRSFlags & InstallOverrideStatus 
    End If

   
End Function 


'****************************************************************************
'*
'*  Function Usage
'*
'****************************************************************************

Function Usage()


    WScript.Echo ""
    WScript.Echo "FRSFLAGS.VBS --- Copyright (c) 2003 Microsoft Corporation"
    WScript.Echo ""
    WScript.Echo "  Enables/Disables/Displays ""install override"" functionality of "
    WScript.Echo "  File Replication Service. Functionality is defined per replica  "
    WScript.Echo "  and replica name along with domain name is needed to set/unset  "
    WScript.Echo "  'fRSFlags' attribute. "
    WScript.Echo ""
    WScript.Echo ""
    WScript.Echo "  Usage: "  
    WScript.Echo ""
    WScript.Echo "     frsflags.vbs { -s | -u | -d } <ReplicaName> <DomainName>"
    WScript.Echo ""
    WScript.Echo "     -s    : Sets the flag for given replica"
    WScript.Echo "     -u    : Unsets the flag for given replica"
    WScript.Echo "     -d    : Displays current value of flag"
    WScript.Echo ""
    WScript.Echo "     <ReplicaName>   : Name of the replica can be either domain based DFS root or "  
    WScript.Echo "                       just SYSVOL"
    WScript.Echo "     <DomainName>    : Full domain name in dotted format. "
    WScript.Echo ""
    WScript.Echo ""
    WScript.Echo "  Examples: "
    WScript.Echo ""
    WScript.Echo "     cscript frsflags.vbs -s ""Root|MyDFS"" ""frs.corp.abc.com"" "
    WScript.Echo "     cscript frsflags.vbs -d ""sysvol"" ""frs.corp.abc.com"" "
    WScript.Echo ""
    WScript.Echo ""
    WScript.Echo "  Verification through ntfrsutl:"
    WScript.Echo ""
    WScript.Echo "     ""ntfrsutl sets"" displays value of Install Override flag under " 
    WScript.Echo "     ""RepSetObjFlags"" attribute. You'll see one of the following lines:"
    WScript.Echo ""
    WScript.Echo "        Disabled :  RepSetObjFlags: 00000000 Flags [<Flags Clear>]      "
    WScript.Echo "        Enabled  :  RepSetObjFlags: 00000001 Flags [InstallOverride ]   "
    WScript.Echo ""
    WScript.Echo ""
    WScript.Echo "  Notes: "
    WScript.Echo ""
    WScript.Echo "     o  For sysvol replica just enter ""sysvol"" (case-insensitive) "
    WScript.Echo "     o  If replica is not found in the AD then script will return (NULL) : 0x80005000 "
    WScript.Echo ""

    WScript.Quit(1)

End Function
