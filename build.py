"""Python script to build  executable"""

import os
import shutil
import subprocess

import PyInstaller.__main__
import PyInstaller.utils.win32.versioninfo as vinfo
import semver  # type: ignore

from swon_version import ANALYZER_VERSION

print("Starting build process...\n")

# Remove any previous build artifacts
try:
    shutil.rmtree("build")
except FileNotFoundError:
    pass

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
