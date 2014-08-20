@Rem -------------------------------------------
@Rem ! L-Bridge.Cmd   version 1.0
@Rem !
@Rem ! Purpose:  Script to push files from the Windows 2000 System Volume
@Rem !           to Export directory (Windows NT version 4, LMRepl).
@Rem !
@Rem ! Instructions: Select one Windows 2000 Domain controller.  
@Rem !               ** You may edit the L-Destination variable below or pass-in **
@Rem !               ** You may optionally swap XCopy for the utility of your choice **
@Rem !               Place this script on your Windows 2000 domain controller.
@Rem !               It is suggested you schedule the script to be run every 2 hours.
@Rem !
@Rem !    Original: December, 1998
@Rem ! Last Update: February, 1999
@Rem ! Last Author:
@Rem !
@Rem !    Comments:
@Rem !
@Rem !
@Rem -------------------------------------------
@Echo Off


:Variables
@Rem -------------------------------------------
@Rem ! Variables
@Rem !     You may edit or pass-in L-Destination
@Rem -------------------------------------------
Set L-Destination=%1

If !%L-Destination%=="!" Goto Instruction
Set L-Name=%0
Set L-Source=\\%USERDNSDOMAIN%\sysvol\%USERDNSDOMAIN%\scripts


@Echo ! ----------------------------------------
@Echo ! Now running: %L-Name%
@Echo !      Source: %L-Source%
@Echo ! Destination: %L-Destination%
@Echo !
Date /T
Time /T

Call :XCopy
@Rem Call :Robocopy

Time /T
@Echo ! ----------------------------------------
Goto End

:XCopy
@Rem -------------------------------------------
@Rem ! Sample section to use XCopy
@Rem !
@Rem !  Pros: Everybody has XCopy
@Rem !  Cons: Xcopy is blind to file deletes.
@Rem !
@Rem -------------------------------------------
@Rem Note: Remove Echo from line below to activate XCopy

Echo Xcopy %L-Source%  %L-Destination%  /s /D

@Rem  --  /S option tells xcopy to include sub-dirs.
@Rem  --  /D option tells xcopy to only copies newer files.
Exit

:RoboCopy
@Rem -------------------------------------------
@Rem ! Sample section to use Robocopy (from Resource Kit)
@Rem !
@Rem !      Pros: Much better copy utility.  Optionally handles Deletes.
@Rem !      Cons: Requires purchase of Windows 2000 Resource Kit.
@Rem ! Super Con: If /PURGE is used, you must first manually copy over
@Rem !            all your original scripts from NTv4, or you will LOSE them
@Rem !            all very quickly.
@Rem !
@Rem -------------------------------------------
@Rem Note: Remove Echo from line below to activate Robocopy

Echo Robocopy %L-Source%  %L-Destination%  /E /PURGE

@Rem  --  /E option tells robocopy to include sub-dirs
@Rem  --  /PURGE option tells robocopy to delete files no longer in source
Exit


:Instruction
@Echo !
@Echo ! %L-NAME% was improperly configured or improperly invoked.
@Echo !
@Echo ! Before using %l-name%, it is strongly suggested that you review the
@Echo ! planning and deployment guides for information on transitioning LMRepl
@Echo ! to the System volume.
@Echo !
@Echo ! The purpose of this command script is to bridge the two replication
@Echo ! architectures for logon scripts between an NTv4 domain controller and
@Echo ! a Windows 2000 domain controller colluding on a single domain (mixed mode).
@Echo !
@Echo ! In short, this command file copies (Xcopies or robocopies) logon scripts
@Echo !    from: %l-Source%
@Echo !      to: %l-destination%
@Echo ! If 'to:' is blank above, you've probably just found your problem.  Either
@Echo ! edit %L-Name% or pass in the destination machine and share.
@Echo !      ex: %l-name% \\mymachine\netlogon\scripts
@Echo !
@Echo !

Goto End

:End
Exit
