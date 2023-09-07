@echo off

setlocal EnableDelayedExpansion

del swon-analyzer.exe

::: Build and Sign the exe
python build.py

::: Signing needs to be more dynamic...
signtool sign /a /s MY /n "Open Source Developer, Darren Richer" /fd SHA256 /t http://time.certum.pl /v swon-analyzer.exe

::: Clean up build artifacts
rmdir /q/s build

::: Build the installer

makensis swon-analyzer.nsi

::: Sign the installer

signtool sign /a /s MY /n "Open Source Developer, Darren Richer" /fd SHA256 /t http://time.certum.pl /v swon-install.exe
