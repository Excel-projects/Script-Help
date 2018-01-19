@ECHO OFF

REM
REM	|--------------------------------------------------------------------------------------------------------------------
REM	| Purpose:			Generic Install
REM	|
REM	| Ver.	Date			Author			Details
REM	| 1.00	27-NOV-2017		Anthony Duguid		Initial version.
REM	|--------------------------------------------------------------------------------------------------------------------*/


REM
REM 	/E 	 = Copies directories and subdirectories, including empty ones. Same as /S /E. May be used to modify /T. 
REM	/D:m-d-y = Copies files changed on or after the specified date. 
REM		   If no date is given, copies only those files whose source time is newer than the destination time. 
REM	/K 	 = Copies attributes. Normal Xcopy will reset read-only attributes. 
REM 	/Q 	 = Does not display file names while copying. 
REM 	/R 	 = Overwrites read-only files. 
REM 	/Y 	 = Suppresses prompting to confirm you want to overwrite an existing destination file. 
REM

REM	Copy the install directory and sub-directories
REM	XCOPY ".\ScriptHelp.xlam" "%AppData%\Microsoft\AddIns\ScriptHelp.xlam" /E /K /Q /R /Y /D
   	XCOPY ".\ScriptHelp.xlam" "%AppData%\Microsoft\Excel\XLSTART\ScriptHelp.xlam" /E /K /Q /R /Y /D
REM	XCOPY ".\ScriptHelp.xlam" "%AppData%\Roaming\Microsoft\Excel\XLSTART\ScriptHelp.xlam" /E /K /Q /R /Y /D
