@echo off

setlocal EnableDelayedExpansion

del swon-analyzer.exe

::: Get the most recent git version tag
for /F "tokens=* usebackq" %%i IN (`git describe --tags --match "v*"` ) do (
  set version=%%i
)

::: Replace 'unreleased' with the version tag
move /y version.py version.py.save
for /F "tokens=* usebackq" %%i IN (version.py.save) do (
   set z=%%i
   echo !z:unreleased=%version%! >> version.py
)

::: Build and Sign the exe
python build.py

::: Signing needs to be more dynamic...
"C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe" sign /a /s MY /n "Open Source Developer, Darren Richer" /t http://time.certum.pl /v swon-analyzer.exe

::: After building the exe, put the original version file back
move /y version.py.save version.py

::: Clean up build artifacts
rmdir /q/s build

::: Create zip file
del swon-analyzer.zip
powershell Compress-Archive swon-analyzer.exe swon-analyzer.zip

::: Sign release artifact
::: del swon-analyzer.zip.asc
::: gpg --detach-sign --armor --local-user EB0F2232 swon-analyzer.zip
