@echo off

setlocal enabledelayedexpansion

for /f "skip=2 tokens=3*" %%i in ('reg query HKLM\SOFTWARE\LibreOffice\Uno\InstallPath /ve') do (
	set LIBREOFFICE_HOME=%%j
)

if "x!LIBREOFFICE_HOME!"=="x" (
	for /f "skip=2 tokens=3*" %%i in ('reg query HKLM\SOFTWARE\Wow6432Node\LibreOffice\Uno\InstallPath /ve') do (
		set LIBREOFFICE_HOME=%%j
	)
	if "x!LIBREOFFICE_HOME!"=="x" (
		echo Error: LibreOffice not found.
		exit /b 1
	)
) else (
	if not exist "%LIBREOFFICE_HOME%" (
		echo Error: %LIBREOFFICE_HOME% not found.
		exit /b 1
	)
)

set PATH=%LIBREOFFICE_HOME%;%PATH%

soffice --version > nul 
if %ERRORLEVEL% NEQ 0 (
	echo Error: Failed to check LibreOffice version.
	exit /b 1
)

for /f "tokens=1,2" %%i in ('soffice --version ^|^ findstr "LibreOffice"') do (
	set LIBREOFFICE_VERSION=%%j
)

for /f "delims=. tokens=1,2" %%i in ("%LIBREOFFICE_VERSION%") do (
	set LIBREOFFICE_MAJOR=%%i
	set LIBREOFFICE_MINOR=%%j
)

if %LIBREOFFICE_MAJOR% LSS 7 (
	echo Error: minimal version of LibreOffice is 7.0.
	echo Your version is %LIBREOFFICE_VERSION%.
	exit /b 1
)

if not exist "%LIBREOFFICE_HOME%\classes\libreoffice.jar" (
	echo Error: failed to find libreoffice.jar.
	exit /b 1
)

echo Found libreoffice.jar at %LIBREOFFICE_HOME%\classes

mkdir app\libs 2> nul
copy /y "%LIBREOFFICE_HOME%\classes\libreoffice.jar" app\libs > nul
