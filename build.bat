@echo off

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
	echo Error: %JAVA_HOME% not exist as file or directory
	exit /b 1
)

echo JDK found at %JAVA_HOME%
set PATH=%JAVA_HOME%\bin;%PATH%
java -version 2> nul
if %ERRORLEVEL% NEQ 0 (
	echo Error: failed to determine Java version
	exit /b 1
)

for /f "tokens=3" %%g in ('java -version 2^>^&1 ^| findstr /i "version"') do (
	set JAVA_VERSION=%%g
)

set JAVA_VERSION=%JAVA_VERSION:"=%
echo Found Java version %JAVA_VERSION%

for /f "delims=.-_ tokens=1-2" %%v in ("%JAVA_VERSION%") do (
	if /I "%%v" EQU "1" (
		set JAVA_MAJOR=%%w
	) else (
		set JAVA_MAJOR=%%v
	)
)

if %JAVA_MAJOR% LSS 17 (
	echo Error: minimal version for Java for this project is 17
	exit /b 1
)

echo.

echo Searching for LibreOffice
if "x%LIBREOFFICE_HOME%"=="x" (
	for /f "skip=2 tokens=3*" %%i in ('reg query HKLM\SOFTWARE\LibreOffice\Uno\InstallPath /ve') do (
		set LIBREOFFICE_HOME=%%j
	)
	echo Found LibreOffice at: !LIBREOFFICE_HOME!
	if "x!LIBREOFFICE_HOME!"=="x" (
		for /f "skip=2 tokens=3*" %%i in ('reg query HKLM\SOFTWARE\Wow6432Node\LibreOffice\Uno\InstallPath /ve') do (
			set LIBREOFFICE_HOME=%%j
		)
	)
	if "x!LIBREOFFICE_HOME!"=="x" (
		echo Error: LibreOffice not found.
		exit /b 1
	)
) else (
	echo Checking LIBREOFFICE_HOME: %LIBREOFFICE_HOME%
	if not exist "%LIBREOFFICE_HOME%" (
		echo Error: %LIBREOFFICE_HOME% not found.
		exit /b 1
	)
)

soffice --version > nul 
if %ERRORLEVEL% NEQ 0 (
	echo Error: Failed to check LibreOffice version.
	exit /b 1
)

set PATH=%LIBREOFFICE_HOME%;%PATH%
for /f "tokens=1,2" %%i in ('soffice --version ^|^ findstr "LibreOffice"') do (
	set LIBREOFFICE_VERSION=%%j
)

for /f "delims=. tokens=1,2" %%i in ("%LIBREOFFICE_VERSION%") do (
	set LIBREOFFICE_MAJOR=%%i
	set LIBREOFFICE_MINOR=%%j
)

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
	javac -sourcepath ./src -d ./bin  %%i
	if !ERRORLEVEL! NEQ 0 (
		echo Compilation of source file %%~nxi has been failed.
		exit /b 1
	)
)

echo Building JAR file
jar cf postproc.jar -C build .
if %ERRORLEVEL% NEQ 0 (
	echo JAR file building has been failed
	exit /b 1
)

echo Building process has been finished successfully.