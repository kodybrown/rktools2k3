<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" omit-xml-declaration="no" indent="no" media-type="text/html"/>
	<xsl:template match="SETSystemStateData">
	<html>
		<head>
			<title>Snapshot Data</title>
		</head>
		<body>
		    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">
	
		<tr>
			<td width="100%" colspan="1" bgcolor="#BD6F57"><font><A NAME="Contents">Contents</A></font></td>
		</tr>
		<tr><td><A HREF="#Header">Header</A></td></tr>
		<tr><td><A HREF="#PageFiles">Page File Information</A></td></tr>
		<tr><td><A HREF="#MemoryInfo">Memory Information</A></td></tr>
		<tr><td><A HREF="#PoolAllocationInformation">Pool Information</A></td></tr>
		<tr><td><A HREF="#ProcessSummary">Process Summary</A></td></tr>
		<tr><td><A HREF="#ProcessStartInfo">Process Start Information</A></td></tr>
		<tr><td><A HREF="#ProcessModuleInfo">Process Module Information</A></td></tr>
		<tr><td><A HREF="#ProcessThreadInfo">Process Thread Information</A></td></tr>
		<tr><td><A HREF="#KernelModuleInfo">Kernel Module Information</A></td></tr>
		<tr><td><A HREF="#PhyDiskInfo">Physical Disk Information</A></td></tr>
		<tr><td><A HREF="#Partition Info">Disk Partition Information</A></td></tr>
		<tr><td><A HREF="#LogDiskInfo">Logical Disk Information</A></td></tr>
		<tr><td><A HREF="#BiosInfo">BIOS Information</A></td></tr>
		<tr><td><A HREF="#ProcessorInfo">Processor Information</A></td></tr>
		<tr><td><A HREF="#NICInfo">NIC Information</A></td></tr>
		<tr><td><A HREF="#OSInfo">OS Information</A></td></tr>
		<tr><td><A HREF="#TimingInfo">Timing Information</A></td></tr>
		</table><br/>
		    <xsl:apply-templates/>
		</body>
		</html>
	
	</xsl:template>
	
	<xsl:template match="SETSystemStateData/Header">	
		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">
	    
		<tr>
			<td width="100%" colspan="2" bgcolor="#BD6F57"><font><A NAME="Header">Header Information  </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
		</tr>			
		
			
			<tr>
			<td>Reliability GUID</td>
			<td width="50%"><xsl:value-of select="ReliabilityGuid"/></td>
			</tr>		
			<tr>
			<td>Initiating Process</td>
			<td width="50%"><xsl:value-of select="InitiatingProcess"/></td>
			</tr>
			<tr>
			<td>Restart Date</td>
			<td width="50%"><xsl:value-of select="RestartDate"/></td>
			</tr>
			<tr>
			<td>Restart Time</td>
			<td width="50%"><xsl:value-of select="RestartTime"/></td>
			</tr>
			<tr>
			<td>Reason Code</td>
			<td width="50%"><xsl:value-of select="ReasonCode"/></td>
			</tr>
			<tr>
			<td>Reason Title</td>
			<td width="50%"><xsl:value-of select="ReasonTitle"/></td>
			</tr>
			<tr>
			<td>RestartType</td>
			<td width="50%"><xsl:value-of select="RestartType"/></td>
			</tr>
			<tr>
			<td>Uptime</td>
			<td width="50%"><xsl:value-of select="SystemUptime"/></td>
			</tr>
			<tr>
			<td>Comment</td>
			<td width="50%"><xsl:value-of select="Comment"/></td>
			</tr>
			
	
			</table>	<br></br>	
		
		
	</xsl:template>
	
	
	<xsl:template match="SETSystemStateData/PageFiles">	
		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="4" bgcolor="#BD6F57"><font><A NAME="PageFiles">Page File Information </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
			</tr>
			<tr>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Path</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Current Size</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Total</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Peak</font></td>			
		</tr>
			<xsl:for-each select="PageFile">
		    <tr>
			<td><xsl:value-of select="Path"/></td>
			<td ><xsl:value-of select="CurrentSize"/></td>
			<td ><xsl:value-of select="Total"/></td>
			<td ><xsl:value-of select="Peak"/></td>			
		</tr>
		</xsl:for-each>	
		</table>	<br></br>	
		    
    </xsl:template>
	
	<xsl:template match="SETSystemStateData/PoolInfo/AllocationInformation">
	<table border="1" cellpadding="1" cellspacing="0"  style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="6" bgcolor="#BD6F57"><font><A NAME="PoolAllocationInformation">Pool Information  </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
			</tr>
		<tr>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Pool Tag</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Pool Type</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF"># Allocs</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF"># Frees</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF"># Bytes</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Session ID</font></td>			
		</tr>
		<xsl:for-each select="TagEntry">
		    <tr>
			<td><xsl:value-of select="PoolTag"/></td>
			<td ><xsl:value-of select="PoolType"/></td>
			<td ><xsl:value-of select="NumAllocs"/></td>
			<td ><xsl:value-of select="NumFrees"/></td>
			<td ><xsl:value-of select="NumBytes"/></td>
			<td ><xsl:value-of select="SessionID"/></td>			
		</tr>
		</xsl:for-each>	
		</table>	<br></br>	
	</xsl:template>	
	
	<xsl:template match="SETSystemStateData/ProcessSummaries">
	<table border="1" cellpadding="1" cellspacing="0"  style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="10" bgcolor="#BD6F57"><font><A NAME="ProcessSummary">Process Summary  </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
			</tr>
		<tr>
			<td bgcolor="#588BC2"><font color="#FFFFFF">PID</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Name</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">User Time</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Kernel Time</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Working Set</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Page Faults</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Commmitted Bytes</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Priority</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Handle Count</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Thread Count</font></td>
		</tr>
		<xsl:for-each select="Process">
		    <tr>
			<td><xsl:value-of select="PID"/></td>
			<td ><xsl:value-of select="Name"/></td>
			<td ><xsl:value-of select="UserTime"/></td>
			<td ><xsl:value-of select="KernelTime"/></td>
			<td ><xsl:value-of select="WorkingSet"/></td>
			<td ><xsl:value-of select="PageFaults"/></td>
			<td ><xsl:value-of select="CommittedBytes"/></td>
			<td ><xsl:value-of select="Priority"/></td>
			<td ><xsl:value-of select="HandleCount"/></td>
			<td ><xsl:value-of select="ThreadCount"/></td>
		</tr>
		</xsl:for-each>	
		</table>	<br></br>	
	</xsl:template>	
	
	<xsl:template match="SETSystemStateData/ProcessStartInfo">
	<table border="1" cellpadding="1" cellspacing="0"  style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="10" bgcolor="#BD6F57"><font><A NAME="ProcessStartInfo">Process Start Information   </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
			</tr>
		<tr>
			<td bgcolor="#588BC2"><font color="#FFFFFF">PID</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Image Name</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Command Line</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Current Directory</font></td>			
		</tr>
		<xsl:for-each select="Process">
		    <tr>
			<td><xsl:value-of select="PID"/></td>
			<td ><xsl:value-of select="ImageName"/></td>
			<td ><xsl:value-of select="CmdLine"/></td>
			<td ><xsl:value-of select="CurrentDir"/></td>			
		</tr>
		</xsl:for-each>	
		</table>	<br></br>	
	</xsl:template>	
	
	<xsl:template match="SETSystemStateData/ProcessesThreadInfo">
	<table border="1" cellpadding="1" cellspacing="0"  style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="8" bgcolor="#BD6F57"><font><A NAME="ProcessThreadInfo">Process Thread Information  </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
			</tr>
		<tr>
			<td bgcolor="#588BC2"><font color="#FFFFFF">PID</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">TID</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Priority</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Context Switches</font></td>			
			<td bgcolor="#588BC2"><font color="#FFFFFF">Start Address</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">User Time</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Kernel Time</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">State</font></td>
		</tr>
		<xsl:for-each select="Process">
		    
		    <xsl:for-each select="Thread">
		    <tr>
			<td><xsl:value-of select="../PID"/></td>
			<td ><xsl:value-of select="TID"/></td>
			<td ><xsl:value-of select="Priority"/></td>
			<td ><xsl:value-of select="ContextSwitches"/></td>	
			<td><xsl:value-of select="StartAddress"/></td>
			<td ><xsl:value-of select="UserTime"/></td>
			<td ><xsl:value-of select="KernelTime"/></td>
			<td ><xsl:value-of select="State"/></td>
			</tr>
			</xsl:for-each>
		
		</xsl:for-each>	
		</table>	<br></br>	
	</xsl:template>	
	
	
	<xsl:template match="SETSystemStateData/ProcessesModuleInfo">
	<table border="1" cellpadding="1" cellspacing="0"  style="border-collapse: collapse" bordercolor="#111111" width="60%" id="AutoNumber1">	
		<tr>
			<td width="80%" colspan="5" bgcolor="#BD6F57"><font><A NAME="ProcessModuleInfo">Process Module Information  </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
			</tr>
		<tr>
			<td bgcolor="#588BC2"><font color="#FFFFFF">PID</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Load Address</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Image Size</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Entry Point</font></td>			
			<td bgcolor="#588BC2"><font color="#FFFFFF">File Name</font></td>			
		</tr>
		<xsl:for-each select="Process">
		    
		    <xsl:for-each select="Module">
		    <tr>
			<td><xsl:value-of select="../PID"/></td>
			<td ><xsl:value-of select="LoadAddr"/></td>
			<td ><xsl:value-of select="ImageSize"/></td>
			<td ><xsl:value-of select="EntryPoint"/></td>	
			<td><xsl:value-of select="FileName"/></td>			
			</tr>
			</xsl:for-each>		
		</xsl:for-each>	
		</table>	<br></br>	
	</xsl:template>	
	
	<xsl:template match="SETSystemStateData/KernelModuleInfo">
	<table border="1" cellpadding="1" cellspacing="0"  style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="80%" colspan="7" bgcolor="#BD6F57"><font><A NAME="KernelModuleInfo">Kernel Module Information  </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
			</tr>
		<tr>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Module Name</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Load Address</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Code</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Data</font></td>			
			<td bgcolor="#588BC2"><font color="#FFFFFF">Paged</font></td>	
			<td bgcolor="#588BC2"><font color="#FFFFFF">Date</font></td>			
			<td bgcolor="#588BC2"><font color="#FFFFFF">Time</font></td>			
		</tr>
		<xsl:for-each select="Module">		    		    
		    <tr>
			<td ><xsl:value-of select="ModuleName"/></td>
			<td ><xsl:value-of select="LoadAddress"/></td>
			<td ><xsl:value-of select="Code"/></td>
			<td ><xsl:value-of select="Data"/></td>	
			<td><xsl:value-of select="Paged"/></td>	
			<td ><xsl:value-of select="Date"/></td>	
			<td><xsl:value-of select="Time"/></td>	
			</tr>			
		</xsl:for-each>	
		</table>	<br></br>	
	</xsl:template>	
	<xsl:template match="OSInfo">	
		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="2" bgcolor="#BD6F57"><A NAME="OSInfo">OS Information  </A><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
		</tr>						
			<xsl:apply-templates/>
		</table>		<br></br>
    </xsl:template>
    <xsl:template match="CurrentBuild">		
			<tr>
			<td width="50%">Current Build</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="CurrentType">		
			<tr>
			<td width="50%">Current Type</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="CurrentVersion">		
			<tr>
			<td width="50%">Current Version</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="Path">		
			<tr>
			<td width="50%">Path</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="ProductName">		
			<tr>
			<td width="50%">Product Name</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="SoftwareType">		
			<tr>
			<td width="50%">Software Type</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="SourcePath">		
			<tr>
			<td width="50%">Source Path</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="SystemRoot">		
			<tr>
			<td width="50%">System Root</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="DebuggerEnabled">		
			<tr>
			<td width="50%">Debugger Enabled</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="Hotfix">		
			<tr>
			<td width="50%">Hotfix</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	
	
	<xsl:template match="Memory">	
		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="2" bgcolor="#BD6F57"><font><A NAME="MemoryInfo">Memory Information  </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
		</tr>						
		<tr>
		<td width="50%">Working Set Size</td>
		<td width="50%"><xsl:value-of select="WorkingSet"/></td>
		</tr>
		
			<xsl:apply-templates/>
		</table>		<br></br>
    </xsl:template>    
	<xsl:template match="PhysicalMemory">	
		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="2" bgcolor="#588BC2">Physical Memory</td>
		</tr>						
			<xsl:apply-templates/>
		</table>		
    </xsl:template>
    <xsl:template match="Total">		
			<tr>
			<td width="50%">Total</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="Available">		
			<tr>
			<td width="50%">Available</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="WorkingSet">						
	</xsl:template>	
	<xsl:template match="CommittedMemory">	
		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="2" bgcolor="#588BC2">Committed Memory</td>
		</tr>						
			<xsl:apply-templates/>
		</table>		
    </xsl:template>
    <xsl:template match="Total">		
			<tr>
			<td width="50%">Total</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="UserMode">		
			<tr>
			<td width="50%">User Mode</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="Limit">		
			<tr>
			<td width="50%">Limit</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="Peak">		
			<tr>
			<td width="50%">Peak</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="KernelMemory">	
		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="2" bgcolor="#588BC2">Kernel Memory</td>
		</tr>						
			<xsl:apply-templates/>
		</table>		
    </xsl:template>
    <xsl:template match="Nonpaged">		
			<tr>
			<td width="50%">Non Paged</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="Paged">		
			<tr>
			<td width="50%">Paged</td>
			<td width="50%"><xsl:apply-templates/></td>
			</tr>		
	</xsl:template>
	<xsl:template match="Pool">	
		<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="2" bgcolor="#588BC2">Memory Pool</td>
		</tr>						
			<xsl:apply-templates/>
		</table>		
    </xsl:template>
    
    
    
    <xsl:template match="SETSystemStateData/HardwareInfo/BiosInfo">	
		<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="2" bgcolor="#BD6F57"><A NAME="BiosInfo">BIOS Information  </A><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
		</tr>						
			<tr>
			<td width="50%">Identifier</td>
			<td width="50%"><xsl:value-of select="Identifier"/></td>
			</tr>	
			<tr>
			<td width="50%">System Bios Date</td>
			<td width="50%"><xsl:value-of select="SystemBiosDate"/></td>
			</tr>
			<tr>
			<td width="50%">System BIOS Version</td>
			<td width="50%"><xsl:value-of select="SystemBiosVersion"/></td>
			</tr>
			<tr>
			<td width="50%">Video Bios Date</td>
			<td width="50%"><xsl:value-of select="VideoBiosDate"/></td>
			</tr>
			<tr>
			<td width="50%">Video BIOS Version</td>
			<td width="50%"><xsl:value-of select="VideoBiosVersion"/></td>
			</tr>
		</table>	<br/>	
    </xsl:template>
    
    <xsl:template match="ProcessorInfo">
	<table border="1" cellpadding="1" cellspacing="0"  style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="100%" colspan="4" bgcolor="#BD6F57"><font><A NAME="ProcessorInfo">Processor Information  </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
			</tr>
		<tr>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Number</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Speed</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Identifier</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Vendor Information</font></td>									
		</tr>
		<xsl:for-each select="Processor">		    		    
		    <tr>
			<td ><xsl:value-of select="Number"/></td>
			<td ><xsl:value-of select="Speed"/></td>
			<td><xsl:value-of select="Identifier"/></td>
			<td><xsl:value-of select="VendorIdent"/></td>				
			</tr>			
		</xsl:for-each>	
		</table>	<br></br>	
	</xsl:template>	
    
    <xsl:template match="NICInfo">	
		<table border="1" cellpadding="1" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="80%" colspan="4" bgcolor="#BD6F57"><font><A NAME="NICInfo">NIC Information  </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
			</tr>
		<tr>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Description</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Service Name</font></td>											
		</tr>
		<xsl:for-each select="NIC">		    		    
		    <tr>
			<td ><xsl:value-of select="Description"/></td>
			<td ><xsl:value-of select="ServiceName"/></td>						
			</tr>			
		</xsl:for-each>	
		</table>	<br/>		
    </xsl:template>
    
    <xsl:template match="DiskInfo/PhysicalDisks">
	<table border="1" cellpadding="1" cellspacing="0"  style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="80%" colspan="10" bgcolor="#BD6F57"><font><A NAME="PhyDiskInfo">Physical Disk Information  </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
			</tr>
		<tr>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Disk ID</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Bytes Per Sector</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Sectors Per Track</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Tracks Per Cylinder</font></td>									
			<td bgcolor="#588BC2"><font color="#FFFFFF">Number of Cylinders</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Port Number</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Path ID</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Target ID</font></td>	
			<td bgcolor="#588BC2"><font color="#FFFFFF">LUN</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Manufacturer</font></td>	
		</tr>
		<xsl:for-each select="Disk">		    		    
		    <tr>
			<td ><xsl:value-of select="DiskID"/></td>
			<td ><xsl:value-of select="BytesPerSector"/></td>
			<td ><xsl:value-of select="SectorsPerTrack"/></td>
			<td ><xsl:value-of select="TracksPerCylinder"/></td>				
			<td ><xsl:value-of select="NumberOfCylinders"/></td>
			<td ><xsl:value-of select="PortNumber"/></td>
			<td ><xsl:value-of select="PathID"/></td>
			<td ><xsl:value-of select="TargetID"/></td>
			<td ><xsl:value-of select="LUN"/></td>
			<td ><xsl:value-of select="Manufacturer"/></td>
			</tr>			
		</xsl:for-each>	
		</table>	<br></br>	
	</xsl:template>	

	<xsl:template match="DiskInfo/PartitionByDiskInfo">
	<table border="1" cellpadding="1" cellspacing="0"  style="border-collapse: collapse" bordercolor="#111111" width="60%" id="AutoNumber1">	
		<tr>
			<td width="80%" colspan="5" bgcolor="#BD6F57"><font><A NAME="Partition Info">Disk Partition Information  </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
			</tr>
		<tr>
			<td bgcolor="#588BC2"><font color="#FFFFFF">DiskID</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">PartitionID</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Extent ID</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Starting Offset</font></td>			
			<td bgcolor="#588BC2"><font color="#FFFFFF">Partition Size</font></td>			
		</tr>
		<xsl:for-each select="Disk">
		    
		    <xsl:for-each select="Partitions">
		    		        
		        <xsl:for-each select="PartitionInfo"> 
		            
		            <xsl:for-each select="Extents"> 
		            
		                <xsl:for-each select="ExtentInfo"> 
		            
		    <tr>
			<td><xsl:value-of select="../../../../DiskID"/></td>
			<td ><xsl:value-of select="../../PartitionID"/></td>
			<td ><xsl:value-of select="ID"/></td>
			<td ><xsl:value-of select="StartingOffset"/></td>	
			<td><xsl:value-of select="PartitionSize"/></td>			
		    </tr>
			            </xsl:for-each>		
			        </xsl:for-each>		
		        </xsl:for-each>	
		    </xsl:for-each>		
		</xsl:for-each>	
		</table>	<br></br>	
	</xsl:template>	

	<xsl:template match="DiskInfo/LogicalDrives">
	<table border="1" cellpadding="1" cellspacing="0"  style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td width="80%" colspan="10" bgcolor="#BD6F57"><font><A Name="LogDiskInfo">Logical Disk Information     </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
			</tr>
		<tr>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Drive Path</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Free Space (bytes)</font></td>
			<td bgcolor="#588BC2"><font color="#FFFFFF">Total Space (bytes)</font></td>				
		</tr>
		<xsl:for-each select="LogicalDriveInfo">		    		    
		    <tr>
			<td ><xsl:value-of select="DrivePath"/></td>
			<td ><xsl:value-of select="FreeSpaceBytes"/></td>
			<td ><xsl:value-of select="TotalSpaceBytes"/></td>			
			</tr>			
		</xsl:for-each>	
		</table>	<br></br>	
	</xsl:template>	
	
	<xsl:template match="Timing">
	<table border="1" cellpadding="1" cellspacing="0"  style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber1">	
		<tr>
			<td  colspan="2" bgcolor="#BD6F57"><font><A NAME="TimingInfo">Timing Information   </A></font><A HREF="#Contents"><font color="#FFFFFF">[Back to Contents]</font></A></td>
	    </tr>
		<tr>
			<td>Summary Information</td>
			<td><xsl:value-of select="SummaryInfo"/></td>					
		</tr>
		<tr>
			<td>Pool Information</td>
			<td><xsl:value-of select="PoolInfo"/></td>					
		</tr>
		<tr>
			<td>Process Information</td>
			<td><xsl:value-of select="ProcessInfo"/></td>					
		</tr>
		<tr>
			<td>Process Startup Information</td>
			<td><xsl:value-of select="ProcessStartupInfo"/></td>					
		</tr>
		<tr>
			<td>Process Thread Information</td>
			<td><xsl:value-of select="ProcessThreadInfo"/></td>					
		</tr>
		<tr>
			<td>Process Module Information</td>
			<td><xsl:value-of select="ProcessModuleInfo"/></td>											
		</tr>
		<tr>
			<td>Drivers Loaded</td>
			<td><xsl:value-of select="LoadedDrivers"/></td>					
		</tr>
		<tr>
			<td>OS Information</td>
			<td><xsl:value-of select="OsInfo"/></td>					
		</tr>
		<tr>
			<td>Hot Fixes</td>
			<td><xsl:value-of select="HotFixes"/></td>					
		</tr>
		<tr>
			<td>Bios Information</td>
			<td><xsl:value-of select="BiosInfo"/></td>					
		</tr>
		<tr>
			<td>Hardware Information</td>
			<td><xsl:value-of select="HardwareInfo"/></td>					
		</tr>
		<tr>
			<td >Physical Disk Information</td>
			<td ><xsl:value-of select="PhysicalDiskInfo"/></td>					
		</tr>
		<tr>
			<td >Logical Drive Information</td>
			<td ><xsl:value-of select="LogicalDriveInfo"/></td>					
		</tr>
		<tr>
			<td >Extension DLL</td>
			<td ><xsl:value-of select="ExtensionDll"/></td>					
		</tr>
		</table>	<br></br>	
	</xsl:template>	
</xsl:stylesheet>
