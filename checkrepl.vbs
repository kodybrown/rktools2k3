'---------------------------------------------------------
'
'  checkrepl.vbs  	
'
'	Monitor Windows 2000 
'	Active Directory Replication
'
'	mattraff 11/20/2002
'
'---------------------------------------------------------

Sub help()
    WScript.Echo ""
    WScript.Echo "Monitor replication and enumerate the"
    WScript.Echo "replication topology for a given DC."
    WScript.Echo ""
    WScript.Echo "usage:"
    WScript.Echo "======"  
    WScript.Echo "cscript checkrepl.vbs server"
    WScript.Echo ""
    WScript.Echo "eg:"
    WScript.Echo "==="     
    WScript.Echo "cscript checkrepl.vbs corpdc01"   
    WScript.Echo ""
    WScript.Echo "Note: This script requires that the Resource Kit tool, iadstools.dll be registered on the local machine."
   
    WScript.Quit
End Sub    
    

Set objArgs = WScript.Arguments

Select Case objArgs.Count
    Case 0
        help
    Case 1
        Select Case objArgs(0)
            Case "-?"
                help
               
            Case "?"
                help

            Case "/?"
                help              
                               
            Case Else

            	
        End Select
    Case Else
End Select


Dim site
Dim objSite
Dim DSACon
Dim nameList
Dim partitionResult
Dim nc
Dim outboundCount
Dim outboundPartners
Dim Partners

Dim server
Dim isPartialNC

Set Wshshell = Wscript.CreateObject("Wscript.shell")

server = objArgs(0)

Set DLL=CreateObject("IADsTools.DCFunctions")

'Use iads getSiteForServer to get the TS Server's site.
site=DLL.getSiteForServer(CStr(server),0)
If site = "" Then
  WScript.Echo ""
  WScript.Echo DLL.LastErrorText
  WScript.Quit
End If


'Get the number of connections For this server and the server names
DSACon=DLL.GetDSAConnections(CStr(server), CStr(site) , CStr(server), 0)
'Or... 
' If you want to direct the query to a specific server you can specify that server in the following line
'DSACon=DLL.GetDSAConnections("Server Name", CStr(site) , CStr(server), 0)

WScript.Echo ""  
WScript.Echo "Inbound Neighbors"
WScript.Echo "" 

For i =  1 to DSACon
  nameList=DLL.DSAConnectionServerName(i) 
  WSCript.Echo "  " & i & ")  " & CStr(nameList) 
Next

'Go through the connections one by one and get what info we can
'For i = 1 to DSACon

WScript.Echo "" 

'Get writable NC's
    partitionResult=DLL.GetNamingContexts(objArgs(0))
    For j=1 to partitionResult
    
      nc = DLL.NamingContextName(j)
      
      Partners=DLL.GetDirectPartners(CStr(server), DLL.NamingContextName(j))
      WScript.Echo ""
      WScript.Echo DLL.ConvertLDAPToDNS(DLL.NamingContextName(j))
      For k=1 to Partners
        WScript.Echo "    " & DLL.DirectPartnerName(k) & " via " & DLL.DirectPartnerTransportDN(k)

        'see if there's a failure code other than zero for any of the replication partners       
        If DLL.DirectPartnerFailReason(k) > 0 then
          wscript.echo "        Failure replicating partition " + DLL.NamingContextName(j)
          WScript.Echo "            @ " & DLL.DirectPartnerLastAttemptTime(k)
          WScript.Echo "            Error: " & DLL.ConvertErrorMsg(DLL.DirectPartnerFailReason(k))
          WScript.Echo "              " & DLL.DirectPartnerNumberFailures(k) & "consecutive failures."
          WScript.Echo "               Last successful attempt was @ " & DLL.DirectPartnerLastSuccessTime(k)  
          WScript.Echo "               Current through property update USN : " & DLL.DirectPartnerHighPU(k)                  
        Else
          wscript.echo "        Last successful attempt: " & DLL.DirectPartnerLastSuccessTime(k)
          WScript.Echo "            Current through property update USN : " & DLL.DirectPartnerHighPU(k)
        End if        
        
      Next      
      
      'Why was this connection generated?        
      'Reason=DLL.DSAConnectionReasonCode(i, j)
      'WSCript.Echo "         Reason for connection:  " & Reason   
    Next

  
    'Get partial NC's for GC's
    PartialPartitionResult=DLL.GetPartialNamingContexts(objArgs(0))
    WScript.Echo "" 

    ' Check to see if we even have any partial NC's to worry about
    If PartialPartitionResult = 0 Then

    Else
    
      For j=1 to PartialPartitionResult
        Partners=DLL.GetDirectPartners(CStr(server), DLL.NamingContextName(j))
        WScript.Echo ""
        WScript.Echo DLL.ConvertLDAPToDNS(DLL.NamingContextName(j))
        For k=1 to Partners
        WScript.Echo "    " & DLL.DirectPartnerName(k) & " via " & DLL.DirectPartnerTransportDN(k)
          
        'see if there's a failure code other than zero for any of the replication partners       
        If DLL.DirectPartnerFailReason(k) > 0 then
          wscript.echo "        Failure replicating partition " + DLL.NamingContextName(j)
          WScript.Echo "            @ " & DLL.DirectPartnerLastAttemptTime(k)
          WScript.Echo "            Error: " & DLL.ConvertErrorMsg(DLL.DirectPartnerFailReason(k))
          WScript.Echo "              " & DLL.DirectPartnerNumberFailures(k) & "consecutive failures."
          WScript.Echo "               Last successful attempt was @ " & DLL.DirectPartnerLastSuccessTime(k)           
          WScript.Echo "               Current through property update USN : " & DLL.DirectPartnerHighPU(k)          
        Else
          wscript.echo "        Last successful attempt: " & DLL.DirectPartnerLastSuccessTime(k)
          WScript.Echo "            Current through property update USN : " & DLL.DirectPartnerHighPU(k)
        End If
        
        Next  
        'Why was this connection generated?  
        'Reason=DLL.DSAConnectionReasonCode(i, j)
        'WSCript.Echo "         Reason for connection:  " & Reason           
      Next
      
     
    End If

'Next

'Enumerate Outbound replication partners.
outboundCount=DLL.GetChangeNotifications(CStr(server), CStr(nc), 0, 0)

WScript.Echo ""
WScript.Echo "Outbound Neighbors:"
WScript.Echo ""

For l = 1 to outboundCount

  WScript.Echo "  " & DLL.NotificationPartnerName(l) & " via " & DLL.NotificationPartnerTransport(l)
  WScript.Echo "        Object GUID: " & DLL.NotificationPartnerObjectGuid(l)  
  
Next