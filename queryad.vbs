'---------------------------------------------------------
'
'  adquery:  perform a forest wide search for an 
'	     object in AD
'
'---------------------------------------------------------

Sub help()
    WScript.Echo ""
    WScript.Echo "Search Active Directory for an object."
    WScript.Echo ""
    WScript.Echo "Usage:"
    Wscript.Echo "======"
    Wscript.Echo "cscript  queryad.vbs SearchFilter Name [GC]"
    WScript.Echo ""
    WScript.Echo "SearchFilter"
    WScript.Echo "============="
    WScript.Echo "-cn        Generic object search by CN"
    WScript.Echo "-c         Search for computer object"
    WScript.Echo "-u         Search for user object by Name"
    WScript.Echo "-ou        Search for an OU"
    WScript.Echo "-dl        Explode a Distribution List"  
    WScript.Echo "-attribute View a schema attribute by display name"       
    WScript.Echo ""    
    WScript.Echo "Name"
    WScript.Echo "===="
    WScript.Echo "Computer's Name, User's CN, OU, CN, or DL's Alias."
    WScript.Echo ""
    WScript.Echo "[GC]"
    WScript.Echo "===="  
    WScript.Echo "Optional - Point the script to query a specific GC."    
    WScript.Echo ""          
    WScript.Echo "examples:"
    WScript.Echo "========="
    WScript.Echo "cscript queryad.vbs -u " &  """" & "Joe Smith" & """"  
    WScript.Echo "cscript queryad.vbs -c computername globalcatalog" 
    WScript.Echo ""
    WScript.Quit
End Sub

Dim Con 
Dim oCommand 
Dim objArgs
Dim ADsObject
Dim sADsPath
Dim objName
Dim objClass
Dim objSchema
Dim classObject

On Error Resume Next

Set objArgs = WScript.Arguments

strName = objArgs(1)	

Select Case objArgs.Count
    Case 0
        help
    Case 1
        help
    Case 2
        Select Case objArgs(0)
            Case "-u"

            Case "-c"

            Case "-ou"

            Case Else
            	
		 help

        End Select
    Case Else
End Select

'--------------------------------------------------------
'Create the ADO connection object
'--------------------------------------------------------

Set Con = CreateObject("ADODB.Connection")
Con.Provider = "ADsDSOObject"
Con.Open "Active Directory Provider"

'Create ADO command object for the connection.
Set oCommand = CreateObject("ADODB.Command")
oCommand.ActiveConnection = Con
 
'Get the ADsPath for the domain to search. 
Set Root = GetObject("LDAP://rootDSE")

'---------------------------------------------------------
'Choose the NC you want to search and build the ADsPath
'---------------------------------------------------------

sDomain = root.Get("rootDomainNamingContext")

If objArgs(0) = "-attribute" Then
	sDomain = root.Get("schemaNamingContext")
End If
	
Set domain = GetObject("GC://" & sDomain)

sADsPath = "<" & domain.ADsPath & ">"
 
'--------------------------------------------------------
'Build the search filter
'--------------------------------------------------------

Select Case objArgs(0)
    Case "-c"
        sFilter = "(&(objectClass=computer)(cn=" & strName & "))"
        sAttribsToReturn = "distinguishedName"

    Case "-u"
        sFilter = "(&(objectCategory=person)(objectClass=user)(Name=" & strName & "))"
        sAttribsToReturn = "distinguishedName"

    Case "-ou"
       sFilter = "(&(objectClass=organizationalUnit)(ou=" & strName & "))"
       sAttribsToReturn = "distinguishedName"

    Case "-cn"
        sFilter = "(cn=" & strName & ")"
        sAttribsToReturn = "distinguishedName"

    Case "-dl"
        sFilter = "(&(dLMemDefault=1)(mailNickname=" & strName & "))"
        sAttribsToReturn = "distinguishedName"

End Select

sDepth = "subtree"

'---------------------------------------------------------
'Assemble and execute the query
'---------------------------------------------------------

oCommand.CommandText = sADsPath & ";" & sFilter & ";" & _
	sAttribsToReturn & ";" & sDepth

Set rs = oCommand.Execute

'---------------------------------------------------------
' Navigate the record set and get the object's DN
'---------------------------------------------------------

rs.MoveFirst
While Not rs.EOF
    For i = 0 To rs.Fields.Count - 1
    	If rs.Fields(i).Name = "distinguishedName" Then
	    Path = rs.Fields(i).Value
        End If        
    Next
    rs.MoveNext
Wend

WScript.Echo "Found " & rs.RecordCount & " objects in the forest"
Wscript.Echo ""

'Quit if nothing is found
If rs.RecordCount = 0 Then
	WScript.Quit
End If

'----------------------------------------------------------
' Bind to the object 
'----------------------------------------------------------

sADsPath = "GC://" & Path

'Did we explicity specify a server to get the info from?
If objArgs(2) > "" Then
	sADsPath = "GC://" & objArgs(2) & "/" & Path
End If	

Set ADsObject = GetObject(sADsPath)

'---------------------------------------------------------
' Display some basic object info
'---------------------------------------------------------

objName = ADsObject.Name
WScript.Echo "Name: " & objName

objClass = ADsObject.Class
WScript.Echo "Class: " & objClass

objSchema = ADsObject.Schema
WScript.Echo "Schema: " & objSchema

'---------------------------------------------------------
' Bind to the class schema object get a properties list
'---------------------------------------------------------

Set classObject = GetObject(ADsObject.Schema)

'---------------------------------------------------------
'Display mandatory properties
'---------------------------------------------------------

For Each PropertyName In classObject.MandatoryProperties
	sPropName = CStr(PropertyName) & ": "
	For Each PropertyValue In ADsObject.GetEx(PropertyName)
	If CStr(PropertyValue) > "" Then
		sText = sPropName & CStr(PropertyValue)
		WScript.Echo sText
	End If
	Next
Next

'---------------------------------------------------------
'Display optional properties
'---------------------------------------------------------

For Each PropertyName In classObject.OptionalProperties
  sPropName = CStr(PropertyName) & ": "

    For Each PropertyValue In ADsObject.GetEx(PropertyName)
	If CStr(PropertyValue) > "" Then	
	  sText = sPropName & CStr(PropertyValue)
	  WScript.Echo sText
	End If
    Next

Next

'---------------------------------------------------------
' Display any child objects
'---------------------------------------------------------

i = 0
WScript.Echo "Child objects:"
For Each Child in ADsObject
	i = i + 1
	objChild = Child
	sObject = Child.Name
	WScript.Echo "  " & Chr(28) & " " & Mid(sObject, 4)
	
 	For Each object in Child
		sGrandChild = child.Name
		WScript.Echo "    " & Chr(28) & " " & Mid(sGrandChild, 4)
		
	Next
Next

