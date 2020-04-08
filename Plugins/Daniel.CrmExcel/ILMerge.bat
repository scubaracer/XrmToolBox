:: PostBuild command:
:: "$(ProjectDir)ILMerge.bat" $(ConfigurationName) "$(TargetDir)" $(TargetFileName)

ECHO Merging library files with plug-in project.
SET BUILDCONFIG=%1
SET TARGETDIR=%2
SET PROJECTBINARY=%3
SET KEYFILE=..\..\CrmPartnersKey.snk
SET ASSEMBLIESTOMERGE=%PROJECTBINARY%

SET TEMP_DLL=tmp_merge\Temp.dll
SET DE_ASSEMBLIESTOMERGE=%TEMP_DLL%

CD %TARGETDIR%

setlocal EnableDelayedExpansion

FOR %%A IN (*.dll) DO (
	SET FILE=%%A
	::check if FILE starts with (m)icrosoft.
	if not "!FILE:~1,9!" == "icrosoft." (
		::check if file equals PROJECTBINARY
		if not %%A == %PROJECTBINARY% (SET ASSEMBLIESTOMERGE=!ASSEMBLIESTOMERGE! %%A)
	)
)
FOR %%A IN (nl\*.dll) DO (
	SET ASSEMBLIESTOMERGE=!ASSEMBLIESTOMERGE! %%A
)
FOR %%A IN (de\*.dll) DO (
	SET DE_ASSEMBLIESTOMERGE=!DE_ASSEMBLIESTOMERGE! %%A
)

ECHO File name: %PROJECTBINARY%
ECHO Assemblies to merge: %ASSEMBLIESTOMERGE%
ECHO Target directory: %TARGETDIR%
ECHO Build configuration: %BUILDCONFIG%


IF EXIST tmp_merge RD /s /q tmp_merge
MD tmp_merge
IF EXIST ILMerge.log DEL /s /q ILMerge.log


SET REFASSEMBLIES=%PROGRAMFILES(X86)%\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.6.2
SET REFLIB=%PROGRAMFILES(X86)%\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.6.2\Facades
::C:\Windows\Microsoft.NET\Framework64\v4.0.30319

::%PROGRAMFILES(X86)%\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.6.2\Facades

:: try find ILMerge in PATH variable...
for %%X in (ILMerge.exe) do (set ILMERGE=%%~$PATH:X)
if not defined ILMERGE (
:: not found, set to program file path.
	::set ILMERGE="%PROGRAMFILES(X86)%\Microsoft\ILMerge\ILMerge.exe"
	set ILMERGE="..\..\..\..\packages\ILMerge.Tools.2.14.1208\tools\ILMerge.exe"
)

ECHO ------Merging------
ECHO ILMerge location: %ILMERGE%
ECHO Reference Assemblies: %REFASSEMBLIES%
ECHO Reference Lib: %REFLIB%
ECHO keyfile : %KEYFILE%
ECHO PROJECTBINARY : %PROJECTBINARY% 
ECHO ASSEMBLIESTOMERGE : %ASSEMBLIESTOMERGE%

%ILMERGE% /log:ILMerge.log /keyfile:"%KEYFILE%" /targetplatform:v4,"%REFASSEMBLIES%" /lib:"%REFLIB%" /out:tmp_merge\%PROJECTBINARY% %ASSEMBLIESTOMERGE% 

DEL tmp_merge\Temp.*

COPY tmp_merge\* . /Y

ECHO Cleaning up...
RD /s/q tmp_merge

::move /Y Daniel.CrmExcel.dll ..\..\..\..\XrmToolbox\bin\debug\Plugins\
COPY Daniel.CrmExcel.*  C:\Users\DJanse\AppData\Roaming\MscrmTools\XrmToolBox\Plugins\
ECHO ------Done------
