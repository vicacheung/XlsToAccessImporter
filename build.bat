@echo off
setlocal
set CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" (
  echo Could not find C# compiler: %CSC%
  exit /b 1
)

if exist bin rmdir /s /q bin
mkdir bin >nul 2>nul

powershell -ExecutionPolicy Bypass -File "%~dp0generate-icon.ps1"
if errorlevel 1 exit /b 1

"%CSC%" /nologo /target:winexe /out:bin\XlsToAccessImporter.exe /win32icon:app.ico /platform:anycpu /reference:System.dll /reference:System.Core.dll /reference:System.Data.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll *.cs
if errorlevel 1 exit /b 1

echo Build succeeded: bin\XlsToAccessImporter.exe
endlocal
