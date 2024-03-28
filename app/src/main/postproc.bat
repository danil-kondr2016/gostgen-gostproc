@echo off

setlocal

@rem Specifying application directory

set DIRNAME=%~dp0
if "%DIRNAME%"=="" DIRNAME=.\
set APP_HOME=%DIRNAME%..
for /d %%i in (%APP_HOME%) do set APP_HOME=%%~fi

set CLASSPATH=%APP_HOME%\lib\postproc.jar
set CLASSPATH=%CLASSPATH%;%HOMEDRIVE%\Program Files\LibreOffice\program\classes\libreoffice.jar

if defined JAVA_HOME (
	%JAVA_HOME%\bin\java.exe ru.danilakondr.gostproc.Application %*
) else (
	java.exe ru.danilakondr.gostproc.Application %*
)

if %errorlevel% NEQ 0 (
	echo Failed to launch application
	exit /b 1
)

exit /b 0
