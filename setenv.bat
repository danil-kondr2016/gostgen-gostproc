@echo off

set JAVA_HOME=%cd%\env\jdk
set GRADLE_HOME=%cd%\env\gradle

set PATH=%JAVA_HOME%\bin;%GRADLE_HOME%\bin;%PATH%
