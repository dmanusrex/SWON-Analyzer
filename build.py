# Wahoo! Results - https://github.com/JohnStrunk/wahoo-results
# Copyright (C) 2022 - John D. Strunk
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as published
# by the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

"""Python script to build Wahoo! Results executable"""

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
