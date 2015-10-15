@echo off
if "%1"=="" goto arg_error
cscript c:\utilities\unzip.vbs %1
goto exit
:arg_error
echo Usage: unzip filename.zip
:exit
