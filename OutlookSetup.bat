@echo off
CLS
ECHO.
ECHO =============================
ECHO Run shell as administrator
ECHO =============================

:init
setlocal DisableDelayedExpansion
set "batchPath=%~0"
for %%k in (%0) do set batchName=%%~nk
set "vbsGetPrivileges=%temp%\OEgetPriv_%batchName%.vbs"
setlocal EnableDelayedExpansion

:checkPrivileges
NET FILE 1>NUL 2>NUL
if '%errorlevel%' == '0' ( goto gotPrivileges ) else ( goto getPrivileges )

:getPrivileges
if '%1'=='ELEV' (echo ELEV & shift /1 & goto gotPrivileges)
ECHO.
ECHO **************************************
ECHO Invoking UAC to Escalate Privileges
ECHO **************************************

ECHO Set UAC = CreateObject^("Shell.Application"^) > "%vbsGetPrivileges%"
ECHO args = "ELEV " >> "%vbsGetPrivileges%"
ECHO For Each strArg in WScript.Arguments >> "%vbsGetPrivileges%"
ECHO args = args ^& strArg ^& " "  >> "%vbsGetPrivileges%"
ECHO Next >> "%vbsGetPrivileges%"
ECHO UAC.ShellExecute "!batchPath!", args, "", "runas", 1 >> "%vbsGetPrivileges%"
"%SystemRoot%\System32\WScript.exe" "%vbsGetPrivileges%" %*
exit /B

:gotPrivileges
setlocal & pushd .
cd /d %~dp0
if '%1'=='ELEV' (del "%vbsGetPrivileges%" 1>nul 2>nul  &  shift /1)

::::::::::::::::::::::::::::
::START OUTLOOK CONFIGURATION
::::::::::::::::::::::::::::

@echo off
taskkill /IM outlook.exe /F
IF EXIST reg query HKCU\Software\Microsoft\Office\16.0 (
reg.exe add HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles\MyEmail /f
IF EXIST reg query HKCU\Software\Microsoft\Office\15.0 (
reg.exe add HKCU\Software\Microsoft\Office\15.0\Outlook\Profiles\MyEmail /f
IF EXIST reg query HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles\MyEmail (
reg.exe add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook" /v "DefaultProfile" /t REG_SZ /d "MyEmail" /f
IF EXIST reg query HKCU\Software\Microsoft\Office\15.0\Outlook\Profiles\MyIntermediaEmail(
reg.exe add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\15.0\Outlook" /v "DefaultProfile" /t REG_SZ /d "MyEmail" /f
start outlook.exe
