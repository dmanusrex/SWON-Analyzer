# Club Analyzer - https://github.com/dmanusrex/SWON-Analyzer
# Copyright (C) 2023 - Darren Richer
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
# IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
# DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
# OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE
#

"""Python script to build Swim Ontario Analyzer executable"""

import os
import shutil
import subprocess

import PyInstaller.__main__
import PyInstaller.utils.win32.versioninfo as vinfo
import semver  # type: ignore
import swon_version

#from version import ANALYZER_VERSION

print("Starting build process...\n")

# Remove any previous build artifacts
try:
    shutil.rmtree("build")
except FileNotFoundError:
    pass


# Determine current git tag
git_ref = (
    subprocess.check_output('git describe --tags --match "v*" --long', shell=True)
    .decode("utf-8")
    .rstrip()
)
ANALYZER_VERSION = swon_version.git_semver(git_ref)

print(f"Building SWON Analyzer, version: {ANALYZER_VERSION}")
version = semver.version.Version.parse(ANALYZER_VERSION)

with open("version.py", "w") as f:
    f.write('"""Version information"""\n\n')
    f.write(f'ANALYZER_VERSION = "{ANALYZER_VERSION}"\n')

    f.flush()
    f.close()

# Create file info to embed in executable
v = vinfo.VSVersionInfo(
    ffi=vinfo.FixedFileInfo(
        filevers=(version.major, version.minor, version.patch, 0),
        prodvers=(version.major, version.minor, version.patch, 0),
        mask=0x3F,
        flags=0x0,
        OS=0x4,
        fileType=0x1,
        subtype=0x0,
    ),
    kids=[
        vinfo.StringFileInfo(
            [
                vinfo.StringTable(
                    "040904e4",
                    [
                        # https://docs.microsoft.com/en-us/windows/win32/menurc/versioninfo-resource
                        # Required fields:
                        vinfo.StringStruct("CompanyName", "Swim Ontario"),
                        vinfo.StringStruct("FileDescription", "Sanctions Analyzer"),
                        vinfo.StringStruct("FileVersion", ANALYZER_VERSION),
                        vinfo.StringStruct("InternalName", "club_analyze"),
                        vinfo.StringStruct("ProductName", "Sanctions Analyzer"),
                        vinfo.StringStruct("ProductVersion", ANALYZER_VERSION),
                        vinfo.StringStruct("OriginalFilename", "swon-analyzer.exe"),
                        # Optional fields
                        vinfo.StringStruct(
                            "LegalCopyright", "(c) Swim Ontario"
                        ),
                    ],
                )
            ]
        ),
        vinfo.VarFileInfo(
            [
                # 1033 -> Engligh; 1252 -> charsetID
                vinfo.VarStruct("Translation", [1033, 1252])
            ]
        ),
    ],
)
with open("swon-analyzer.fileinfo", "w") as f:
    f.write(str(v))
    f.flush()
    f.close()

print("Invoking PyInstaller to generate executable...\n")

# Build it
PyInstaller.__main__.run(["--distpath=.", "--workpath=build", "swon-analyzer.spec"])
