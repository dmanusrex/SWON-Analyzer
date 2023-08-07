@echo off

setlocal EnableDelayedExpansion

del swon-analyzer.exe

::: Build and Sign the exe
python build.py

::: Signing needs to be more dynamic...
signtool sign /a /s MY /n "Open Source Developer, Darren Richer" /fd SHA256 /t http://time.certum.pl /v swon-analyzer.exe

::: Clean up build artifacts
rmdir /q/s build

::: Create zip file
del swon-analyzer.zip
powershell Compress-Archive swon-analyzer.exe swon-analyzer.zip

::: Sign release artifact
::: del swon-analyzer.zip.asc
::: gpg --detach-sign --armor --local-user EB0F2232 swon-analyzer.zip
