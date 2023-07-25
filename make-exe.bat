@echo off

setlocal EnableDelayedExpansion

del swon-analyzer.exe

::: Set up the environment
::: python -m venv venv
::: call venv\Scripts\activate.bat
::: python -m pip install --upgrade pip
::: pip install --upgrade -r requirements.txt

::: Get the most recent git version tag
::: for /F "tokens=* usebackq" %%i IN (`git describe --tags --match "v*"` ) do (
:::    set version=%%i
::: )
::: set version="0.01"

::: Replace 'unreleased' with the version tag
::: move /y version.py version.py.save
::: for /F "tokens=* usebackq" %%i IN (version.py.save) do (
:::    set z=%%i
:::    echo !z:unreleased=%version%! >> version.py
:::)


::: pyinstaller --onefile ^
:::    --noconsole ^
:::    --distpath=. ^
:::    --workpath=build ^
:::   --add-data swon-analyzer.ico;. ^
:::    --icon swon-analyzer.ico ^
:::    --name swon-analyzer ^
:::    club_analyze.py

::: Sign the exe
python build.py

"C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe" sign /a /s MY /sha1 19AC2E43A5A5F506D7B6A7233FB005434F91B252 /t http://timestamp.digicert.com /v "C:\test code\swon-analyzer.exe"

::: After building the exe, put the original version file back
::: move /y version.py.save version.py

::: Clean up build artifacts
::: del swon-analyzer.spec
rmdir /q/s build

::: Create zip file
del swon-analyzer.zip
powershell Compress-Archive swon-analyzer.exe swon-analyzer.zip

::: Sign release artifact
del swon-analyzer.zip.asc
gpg --detach-sign --armor --local-user EB0F2232 swon-analyzer.zip
