@echo off

set __JAVA_MIN_VERSION=18
set OUT_DIR=.\bin

echo The Shloma-style Java application building script
echo Written by Danila A. Kondratenko
echo (c) 2024
echo.

setlocal enableextensions
setlocal enabledelayedexpansion
echo Building md2writer project

if "x%JAVA_HOME%"=="x" (
	echo Error: %%JAVA_HOME%% has not been specified.
	exit /b 1
)

if not exist "%JAVA_HOME%" (
	echo Error: %JAVA_HOME% don't exists.
	exit /b 1
)

set PATH=%JAVA_HOME%\bin;%PATH%
java -version 2> nul
if %ERRORLEVEL% NEQ 0 (
	echo Error: failed to determine Java version.
	exit /b 1
)

for /f "tokens=3" %%g in ('java -version 2^>^&1 ^| findstr /i "version"') do (
	set JAVA_VERSION=%%g
)

set JAVA_VERSION=%JAVA_VERSION:"=%
echo Found JDK %JAVA_VERSION% at %JAVA_HOME%

for /f "delims=.-_ tokens=1-2" %%v in ("%JAVA_VERSION%") do (
	if /I "%%v" EQU "1" (
		set JAVA_MAJOR=%%w
	) else (
		set JAVA_MAJOR=%%v
	)
)

if %JAVA_MAJOR% LSS %__JAVA_MIN_VERSION% (
	echo Error: minimal version for Java is %__JAVA_MIN_VERSION%, your version is %JAVA_VERSION%.
	exit /b 1
)

call lopath 2> nul
if errorlevel 1 (
	echo Error: failed to find LibreOffice.
	exit /b 1
)

for /f "delims=; tokens=2,3 usebackq" %%i in (`call lopath 2^>^&1`) do (
	set LIBREOFFICE_VERSION=%%i
	set LIBREOFFICE_HOME=%%j
)

for /f "delims=. tokens=1,2" %%i in ("%LIBREOFFICE_VERSION%") do (
	set LIBREOFFICE_MAJOR=%%i
	set LIBREOFFICE_MINOR=%%j
)

echo Found LibreOffice %LIBREOFFICE_VERSION% at %LIBREOFFICE_HOME%

if %LIBREOFFICE_MAJOR% LSS 7 (
	echo Error: Minimal supported version of LibreOffice is 7.0. 
	echo Your version is %LIBREOFFICE_MAJOR%.%LIBREOFFICE_MINOR%
	exit /b 1
)

if not exist "%LIBREOFFICE_HOME%\classes\libreoffice.jar" (
	echo Failed to find libreoffice.jar classes.
	exit /b 1
)

set CLASSPATH=.;%LIBREOFFICE_HOME%\classes\libreoffice.jar

echo.

echo Creating build directories
mkdir bin >nul	

echo Building class files
for /r src\ %%i in (*.java) do (
	echo %%i
	javac -sourcepath ./src -d %OUT_DIR%  %%i
	if !ERRORLEVEL! NEQ 0 (
		echo Compilation of source file %%~nxi has been failed.
		exit /b 1
	)
)

echo Building JAR file
jar cf postproc.jar -C %OUT_DIR% .
if %ERRORLEVEL% NEQ 0 (
	echo JAR file building has been failed
	exit /b 1
)

echo Creating postproc.bat
echo @java -cp "%CLASSPATH%;postproc.jar" ru.danilakondr.md2writer.Application %%* > postproc.bat

echo Building process has been finished successfully.
