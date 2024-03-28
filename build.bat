@echo off

setlocal enabledelayedexpansion

echo Copying LibreOffice libraries
call locopy

echo Creating output directory
if exist out (
	dir /a:d out >nul 2>nul
	if errorlevel 1 (
		del /q out
	) else (
		rd /s /q out
	)
)

mkdir out
mkdir out\bin
mkdir out\lib

call gradle -version >nul 2>&1
if %errorlevel% EQU 0 (
	call gradle build
) else (
	call gradlew build
)

copy app\build\libs\app.jar out\lib\postproc.jar
copy app\src\main\postproc.bat out\bin\postproc.bat
