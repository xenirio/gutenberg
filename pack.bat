@echo off
pushd %~dp0

set MSBUILD="C:\Program Files (x86)\MSBuild\14.0\Bin\MSBuild.exe"
if not exist %MSBUILD% (
    echo [ERROR] MSBuild Required, https://www.microsoft.com/en-us/download/details.aspx?id=48159
    goto error_exit
)
set NUEGT="%~dp0nuget.exe"
if not exist %NUEGT% (
    powershell -Command "Invoke-WebRequest https://dist.nuget.org/win-x86-commandline/latest/nuget.exe -OutFile nuget.exe"
)

%NUEGT% pack %1 -OutputDirectory .\artifacts

:end
echo Packing successfully.
EXIT 0

:error_exit
pause
EXIT 1